defmodule ExVEx.OOXML.Table do
  @moduledoc """
  Parse, mutate, and serialize an `xl/tables/table*.xml` part.

  A table is a rectangular `ref` on one worksheet with a header row
  (`header_row_count` of `1` or `0`), zero or one totals row
  (`totals_row_count`), and one `%Column{}` per column of the range.
  Column names must equal the text of the header cells; enforcing that
  is the caller's job (see `ExVEx.Workbook.Tables`).

  The struct keeps the root attribute list and child list it was parsed
  from. Serialization rewrites only the modelled attributes and the
  `<autoFilter>`, `<sortState>`, `<tableColumns>`, and `<tableStyleInfo>`
  children; everything else is emitted as read.
  """

  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.SqrefShift
  alias ExVEx.OOXML.Table.{Column, StyleInfo}
  alias ExVEx.Utils.Range

  @main_ns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  @type node_tuple :: {String.t(), list(), list()}

  @type t :: %__MODULE__{
          id: pos_integer(),
          name: String.t(),
          display_name: String.t(),
          ref: Range.t(),
          header_row_count: 0 | 1,
          totals_row_count: 0 | 1,
          totals_row_shown: boolean(),
          columns: [Column.t()],
          style: StyleInfo.t() | nil,
          auto_filter: node_tuple() | nil,
          sort_state: node_tuple() | nil,
          attrs: [{String.t(), String.t()}],
          children: [node_tuple() | String.t()]
        }

  @enforce_keys [:id, :name, :display_name, :ref]
  defstruct [
    :id,
    :name,
    :display_name,
    :ref,
    header_row_count: 1,
    totals_row_count: 0,
    totals_row_shown: true,
    columns: [],
    style: nil,
    auto_filter: nil,
    sort_state: nil,
    attrs: [],
    children: []
  ]

  @doc """
  Builds a new table. Required options: `:id`, `:name`, `:ref`,
  `:columns` (a list of names). Optional: `:style` (`%StyleInfo{}` or
  `nil`), `:header_row` (default `true`).
  """
  @spec new(keyword()) :: t()
  def new(opts) do
    name = Keyword.fetch!(opts, :name)
    ref = Keyword.fetch!(opts, :ref)
    header_row_count = if Keyword.get(opts, :header_row, true), do: 1, else: 0

    columns =
      opts
      |> Keyword.fetch!(:columns)
      |> Enum.with_index(1)
      |> Enum.map(fn {column_name, id} -> Column.new(id, column_name) end)

    %__MODULE__{
      id: Keyword.fetch!(opts, :id),
      name: name,
      display_name: name,
      ref: ref,
      header_row_count: header_row_count,
      totals_row_shown: false,
      columns: columns,
      style: Keyword.get(opts, :style, %StyleInfo{}),
      auto_filter: if(header_row_count == 1, do: {"autoFilter", [{"ref", ""}], []}),
      attrs: [{"xmlns", @main_ns}],
      children: []
    }
    |> sync_auto_filter()
  end

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"table", attrs, children}} -> from_tree(attrs, children)
      {:ok, _other} -> {:error, :not_a_table_file}
      {:error, reason} -> {:error, reason}
    end
  end

  @spec serialize(t()) :: binary()
  def serialize(%__MODULE__{} = table) do
    attrs =
      table.attrs
      |> put_attr("id", Integer.to_string(table.id))
      |> put_attr("name", table.name)
      |> put_attr("displayName", table.display_name)
      |> put_attr("ref", Range.to_string(table.ref))
      |> put_count_attr("headerRowCount", table.header_row_count, 1)
      |> put_count_attr("totalsRowCount", table.totals_row_count, 0)
      |> put_flag_attr("totalsRowShown", table.totals_row_shown)

    children =
      table.children
      |> replace_child("autoFilter", table.auto_filter)
      |> replace_child("sortState", table.sort_state)
      |> replace_child("tableColumns", columns_node(table.columns))
      |> replace_child("tableStyleInfo", table.style && StyleInfo.to_node(table.style))

    {"table", attrs, children}
    |> to_saxy()
    |> Saxy.encode!(version: "1.0", encoding: "UTF-8", standalone: true)
  end

  @spec column_names(t()) :: [String.t()]
  def column_names(%__MODULE__{columns: columns}), do: Enum.map(columns, & &1.name)

  @spec auto_filter_ref(t()) :: String.t() | nil
  def auto_filter_ref(%__MODULE__{auto_filter: nil}), do: nil

  def auto_filter_ref(%__MODULE__{auto_filter: {"autoFilter", attrs, _}}) do
    case List.keyfind(attrs, "ref", 0) do
      {_, ref} -> ref
      nil -> nil
    end
  end

  @spec header_row(t()) :: pos_integer() | nil
  def header_row(%__MODULE__{header_row_count: 0}), do: nil
  def header_row(%__MODULE__{ref: %Range{top_left: {top, _}}}), do: top

  @spec totals_row(t()) :: pos_integer() | nil
  def totals_row(%__MODULE__{totals_row_count: 0}), do: nil
  def totals_row(%__MODULE__{ref: %Range{bottom_right: {bottom, _}}}), do: bottom

  @doc "The rows holding data: `{first, last}`, or `nil` when the table has no data rows."
  @spec data_rows(t()) :: {pos_integer(), pos_integer()} | nil
  def data_rows(%__MODULE__{} = table) do
    {top, _} = table.ref.top_left
    {bottom, _} = table.ref.bottom_right
    first = top + table.header_row_count
    last = bottom - table.totals_row_count
    if first <= last, do: {first, last}, else: nil
  end

  @spec put_ref(t(), Range.t()) :: t()
  def put_ref(%__MODULE__{} = table, %Range{} = ref) do
    %{table | ref: ref}
    |> sync_auto_filter()
    |> drop_sort_state_outside_ref()
  end

  @spec put_totals_row_count(t(), 0 | 1) :: t()
  def put_totals_row_count(%__MODULE__{} = table, count) when count in [0, 1] do
    shown = table.totals_row_shown or count == 1
    %{table | totals_row_count: count, totals_row_shown: shown} |> sync_auto_filter()
  end

  @spec put_style(t(), StyleInfo.t() | nil) :: t()
  def put_style(%__MODULE__{} = table, style), do: %{table | style: style}

  @spec rename(t(), String.t()) :: t()
  def rename(%__MODULE__{} = table, name) when is_binary(name) do
    %{table | name: name, display_name: name}
  end

  @spec rename_column(t(), String.t(), String.t()) :: t()
  def rename_column(%__MODULE__{columns: columns} = table, old, new) do
    new_columns =
      Enum.map(columns, fn column ->
        if same_name?(column.name, old), do: Column.rename(column, new), else: column
      end)

    %{table | columns: new_columns}
  end

  @doc """
  Replaces the column list with one column per name, in order. Existing
  columns are matched by position and keep their ids and attributes;
  extra names become new columns; surplus columns are dropped.
  """
  @spec put_column_names(t(), [String.t()]) :: t()
  def put_column_names(%__MODULE__{columns: columns} = table, names) when is_list(names) do
    next_id = next_column_id(columns)

    new_columns =
      names
      |> Enum.with_index()
      |> Enum.map(fn {name, index} ->
        case Enum.at(columns, index) do
          nil -> Column.new(next_id + index - length(columns), name)
          column -> Column.rename(column, name)
        end
      end)

    %{table | columns: new_columns}
  end

  @spec put_column(t(), String.t(), Column.t()) :: t()
  def put_column(%__MODULE__{columns: columns} = table, name, %Column{} = replacement) do
    new_columns =
      Enum.map(columns, fn column ->
        if same_name?(column.name, name), do: replacement, else: column
      end)

    %{table | columns: new_columns}
  end

  @spec fetch_column(t(), String.t()) :: {:ok, Column.t()} | :error
  def fetch_column(%__MODULE__{columns: columns}, name) do
    case Enum.find(columns, &same_name?(&1.name, name)) do
      nil -> :error
      column -> {:ok, column}
    end
  end

  @spec column_index(t(), String.t()) :: {:ok, non_neg_integer()} | :error
  def column_index(%__MODULE__{columns: columns}, name) do
    case Enum.find_index(columns, &same_name?(&1.name, name)) do
      nil -> :error
      index -> {:ok, index}
    end
  end

  @doc """
  Applies a structural row/column shift. Returns `:deleted` when the
  shift removes every row or every column of the table.

  Rows: an insert at or above the header moves the whole table; an
  insert inside grows it. A delete that removes every data row leaves
  the header plus one blank data row and drops the totals row. A delete
  that removes the header row or the totals row drops the totals row.

  Columns: an insert at the first column moves the table; an insert
  inside adds placeholder columns named `ColumnN`. A delete inside drops
  the affected columns and their filter settings.
  """
  @spec shift(t(), MutShift.t()) :: {:ok, t()} | :deleted
  def shift(%__MODULE__{} = table, %MutShift{axis: :row} = shift) do
    {top, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right

    case shift_span(top, bottom, shift) do
      :deleted ->
        :deleted

      {new_top, new_bottom} ->
        totals_count = totals_after_row_shift(table, top, bottom, new_top, new_bottom, shift)
        header = table.header_row_count
        minimum_bottom = new_top + header + totals_count
        final_bottom = max(new_bottom, minimum_bottom)

        ref = %Range{top_left: {new_top, left}, bottom_right: {final_bottom, right}}

        {:ok,
         %{table | ref: ref, totals_row_count: totals_count}
         |> shift_sort_state(shift)
         |> sync_auto_filter()}
    end
  end

  def shift(%__MODULE__{} = table, %MutShift{axis: :col} = shift) do
    {top, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right

    case shift_span(left, right, shift) do
      :deleted ->
        :deleted

      {new_left, new_right} ->
        ref = %Range{top_left: {top, new_left}, bottom_right: {bottom, new_right}}

        {:ok,
         %{table | ref: ref}
         |> shift_columns(left, right, shift)
         |> shift_filter_columns(left, right, shift)
         |> shift_sort_state(shift)
         |> sync_auto_filter()}
    end
  end

  defp shift_span(first, last, %MutShift{delta: delta, at: at, count: count}) when delta > 0 do
    cond do
      at <= first -> {first + count, last + count}
      at <= last -> {first, last + count}
      true -> {first, last}
    end
  end

  defp shift_span(first, last, %MutShift{at: at, count: count}) do
    span_end = at + count - 1

    if at <= first and span_end >= last do
      :deleted
    else
      {shift_deleted_start(first, at, count), shift_deleted_end(last, at, count)}
    end
  end

  defp shift_deleted_start(index, at, count) when index >= at + count, do: index - count
  defp shift_deleted_start(index, at, _count) when index >= at, do: at
  defp shift_deleted_start(index, _at, _count), do: index

  defp shift_deleted_end(index, at, count) when index >= at + count, do: index - count
  defp shift_deleted_end(index, at, _count) when index >= at, do: at - 1
  defp shift_deleted_end(index, _at, _count), do: index

  defp totals_after_row_shift(%__MODULE__{totals_row_count: 0}, _, _, _, _, _), do: 0

  defp totals_after_row_shift(table, top, bottom, new_top, new_bottom, %MutShift{delta: delta})
       when delta > 0 do
    data_rows_after(table, new_top, new_bottom, 1) |> keep_totals_if_data(top, bottom)
  end

  defp totals_after_row_shift(table, top, bottom, new_top, new_bottom, shift) do
    header_hit? = table.header_row_count == 1 and in_span?(top, shift)
    totals_hit? = in_span?(bottom, shift)
    data_after = data_rows_after(table, new_top, new_bottom, 1)

    if header_hit? or totals_hit? or data_after <= 0, do: 0, else: 1
  end

  defp keep_totals_if_data(data_after, _top, _bottom) when data_after > 0, do: 1
  defp keep_totals_if_data(_data_after, _top, _bottom), do: 0

  defp data_rows_after(table, new_top, new_bottom, totals) do
    new_bottom - new_top + 1 - table.header_row_count - totals
  end

  defp in_span?(index, %MutShift{at: at, count: count}), do: index >= at and index < at + count

  defp shift_columns(table, left, _right, %MutShift{delta: delta, at: at, count: count})
       when delta > 0 do
    if at <= left do
      table
    else
      insert_placeholder_columns(table, at - left, count)
    end
  end

  defp shift_columns(%__MODULE__{columns: columns} = table, left, _right, shift) do
    kept =
      columns
      |> Enum.with_index()
      |> Enum.reject(fn {_column, index} -> in_span?(left + index, shift) end)
      |> Enum.map(&elem(&1, 0))

    %{table | columns: kept}
  end

  defp insert_placeholder_columns(%__MODULE__{columns: columns} = table, offset, count) do
    next_id = next_column_id(columns)
    {before, after_columns} = Enum.split(columns, offset)
    taken = columns |> Enum.map(&String.downcase(&1.name)) |> MapSet.new()

    {new_columns, _taken} =
      Enum.map_reduce(0..(count - 1)//1, taken, fn index, taken ->
        name = placeholder_name(offset + index + 1, taken)
        {Column.new(next_id + index, name), MapSet.put(taken, String.downcase(name))}
      end)

    %{table | columns: before ++ new_columns ++ after_columns}
  end

  @doc "Returns `ColumnN`, bumping `N` until the name is not in `taken` (lower-cased names)."
  @spec placeholder_name(pos_integer(), MapSet.t(String.t())) :: String.t()
  def placeholder_name(position, taken) do
    candidate = "Column#{position}"

    if MapSet.member?(taken, String.downcase(candidate)),
      do: placeholder_name(position + 1, taken),
      else: candidate
  end

  defp shift_filter_columns(%__MODULE__{auto_filter: nil} = table, _left, _right, _shift),
    do: table

  defp shift_filter_columns(
         %__MODULE__{auto_filter: {tag, attrs, children}} = table,
         left,
         _right,
         shift
       ) do
    new_children =
      Enum.flat_map(children, fn
        {"filterColumn", fc_attrs, fc_children} ->
          shift_filter_column(fc_attrs, fc_children, left, shift)

        other ->
          [other]
      end)

    %{table | auto_filter: {tag, attrs, new_children}}
  end

  defp shift_filter_column(attrs, children, left, shift) do
    with {_, raw} <- List.keyfind(attrs, "colId", 0),
         {col_id, ""} <- Integer.parse(raw) do
      case MutShift.apply_index(shift, left + col_id) do
        :unchanged ->
          [{"filterColumn", attrs, children}]

        :deleted ->
          []

        {:ok, new_index} ->
          [
            {"filterColumn", put_attr(attrs, "colId", Integer.to_string(new_index - left)),
             children}
          ]
      end
    else
      _ -> [{"filterColumn", attrs, children}]
    end
  end

  defp shift_sort_state(%__MODULE__{sort_state: nil} = table, _shift), do: table

  defp shift_sort_state(%__MODULE__{sort_state: node} = table, shift) do
    case shift_ref_node(node, shift) do
      nil -> %{table | sort_state: nil}
      shifted -> %{table | sort_state: shifted}
    end
  end

  defp shift_ref_node({tag, attrs, children}, shift) do
    case List.keyfind(attrs, "ref", 0) do
      {_, ref} ->
        case SqrefShift.shift(ref, shift) do
          "" -> nil
          new_ref -> {tag, put_attr(attrs, "ref", new_ref), shift_ref_children(children, shift)}
        end

      nil ->
        {tag, attrs, shift_ref_children(children, shift)}
    end
  end

  defp shift_ref_children(children, shift) do
    Enum.flat_map(children, fn
      {_, _, _} = node -> List.wrap(shift_ref_node(node, shift))
      other -> [other]
    end)
  end

  defp drop_sort_state_outside_ref(%__MODULE__{sort_state: nil} = table), do: table

  defp drop_sort_state_outside_ref(%__MODULE__{sort_state: {_, attrs, _}, ref: ref} = table) do
    with {_, sort_ref} <- List.keyfind(attrs, "ref", 0),
         {:ok, sort_range} <- Range.parse(sort_ref),
         true <-
           Range.contains?(ref, sort_range.top_left) and
             Range.contains?(ref, sort_range.bottom_right) do
      table
    else
      _ -> %{table | sort_state: nil}
    end
  end

  defp sync_auto_filter(%__MODULE__{auto_filter: nil} = table), do: table

  defp sync_auto_filter(%__MODULE__{auto_filter: {tag, attrs, children}} = table) do
    {top, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right
    filter_bottom = max(bottom - table.totals_row_count, top)

    filter_ref =
      Range.to_string(%Range{top_left: {top, left}, bottom_right: {filter_bottom, right}})

    %{table | auto_filter: {tag, put_attr(attrs, "ref", filter_ref), children}}
  end

  defp next_column_id(columns) do
    columns |> Enum.map(& &1.id) |> Enum.max(fn -> 0 end) |> Kernel.+(1)
  end

  defp same_name?(a, b), do: String.downcase(a) == String.downcase(b)

  defp from_tree(attrs, children) do
    with {:ok, id} <- int_attr(attrs, "id"),
         {:ok, ref} <- range_attr(attrs, "ref") do
      name = attr(attrs, "name") || attr(attrs, "displayName") || ""

      {:ok,
       %__MODULE__{
         id: id,
         name: name,
         display_name: attr(attrs, "displayName") || name,
         ref: ref,
         header_row_count: count_attr(attrs, "headerRowCount", 1),
         totals_row_count: count_attr(attrs, "totalsRowCount", 0),
         totals_row_shown: attr(attrs, "totalsRowShown") not in ["0", "false"],
         columns: parse_columns(children),
         style: children |> find_child("tableStyleInfo") |> maybe(&StyleInfo.from_node/1),
         auto_filter: find_child(children, "autoFilter"),
         sort_state: find_child(children, "sortState"),
         attrs: attrs,
         children: children
       }}
    end
  end

  defp parse_columns(children) do
    case find_child(children, "tableColumns") do
      {"tableColumns", _, items} ->
        for {"tableColumn", _, _} = node <- items, do: Column.from_node(node)

      nil ->
        []
    end
  end

  defp columns_node(columns) do
    {"tableColumns", [{"count", Integer.to_string(length(columns))}],
     Enum.map(columns, &Column.to_node/1)}
  end

  @child_order ["autoFilter", "sortState", "tableColumns", "tableStyleInfo", "extLst"]

  defp replace_child(children, tag, nil), do: Enum.reject(children, &match?({^tag, _, _}, &1))

  defp replace_child(children, tag, node) do
    if Enum.any?(children, &match?({^tag, _, _}, &1)) do
      Enum.map(children, fn
        {^tag, _, _} -> node
        other -> other
      end)
    else
      insert_in_schema_order(children, tag, node)
    end
  end

  defp insert_in_schema_order(children, tag, node) do
    successors = @child_order |> Enum.drop_while(&(&1 != tag)) |> Enum.drop(1)

    index =
      Enum.find_index(children, fn
        {name, _, _} -> name in successors
        _ -> false
      end) || length(children)

    List.insert_at(children, index, node)
  end

  defp to_saxy({tag, attrs, children}) do
    Saxy.XML.element(tag, attrs, Enum.map(children, &to_saxy_child/1))
  end

  defp to_saxy_child({_, _, _} = node), do: to_saxy(node)
  defp to_saxy_child(text) when is_binary(text), do: Saxy.XML.characters(text)
  defp to_saxy_child(other), do: other

  defp find_child(children, tag), do: Enum.find(children, &match?({^tag, _, _}, &1))

  defp maybe(nil, _fun), do: nil
  defp maybe(value, fun), do: fun.(value)

  defp attr(attrs, key) do
    case List.keyfind(attrs, key, 0) do
      {_, value} -> value
      nil -> nil
    end
  end

  defp int_attr(attrs, key) do
    case Integer.parse(attr(attrs, key) || "") do
      {value, ""} -> {:ok, value}
      _ -> {:error, {:invalid_table_attribute, key}}
    end
  end

  defp range_attr(attrs, key) do
    case Range.parse(attr(attrs, key) || "") do
      {:ok, range} -> {:ok, range}
      :error -> {:error, {:invalid_table_attribute, key}}
    end
  end

  defp count_attr(attrs, key, default) do
    case Integer.parse(attr(attrs, key) || "") do
      {value, ""} -> value
      _ -> default
    end
  end

  defp put_attr(attrs, key, value) do
    case List.keyfind(attrs, key, 0) do
      nil -> attrs ++ [{key, value}]
      _ -> List.keyreplace(attrs, key, 0, {key, value})
    end
  end

  defp put_count_attr(attrs, key, value, default) do
    case {List.keyfind(attrs, key, 0), value} do
      {nil, ^default} -> attrs
      _ -> put_attr(attrs, key, Integer.to_string(value))
    end
  end

  defp put_flag_attr(attrs, key, true), do: List.keydelete(attrs, key, 0)
  defp put_flag_attr(attrs, key, false), do: put_attr(attrs, key, "0")
end
