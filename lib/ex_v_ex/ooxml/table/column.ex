defmodule ExVEx.OOXML.Table.Column do
  @moduledoc """
  One `<tableColumn>` of a table part.

  `id` is unique within the table and never reused; `name` must equal
  the text of the column's header cell. `attrs` and `children` hold the
  element's full attribute list and child elements, so attributes and
  children not modelled here (`dataDxfId`, `queryTableFieldId`,
  `<xmlColumnPr>`, …) pass through untouched.
  """

  @type node_tuple :: {String.t(), list(), list()}

  @type t :: %__MODULE__{
          id: pos_integer(),
          name: String.t(),
          attrs: [{String.t(), String.t()}],
          children: [node_tuple() | String.t()]
        }

  @enforce_keys [:id, :name]
  defstruct [:id, :name, attrs: [], children: []]

  @totals_functions %{
    sum: "sum",
    average: "average",
    count: "count",
    count_nums: "countNums",
    max: "max",
    min: "min",
    std_dev: "stdDev",
    var: "var",
    none: "none",
    custom: "custom"
  }

  @type totals_function ::
          :sum | :average | :count | :count_nums | :max | :min | :std_dev | :var | :none | :custom

  @spec new(pos_integer(), String.t()) :: t()
  def new(id, name) when is_integer(id) and is_binary(name) do
    %__MODULE__{id: id, name: name, attrs: [{"id", Integer.to_string(id)}, {"name", name}]}
  end

  @spec from_node(node_tuple()) :: t()
  def from_node({"tableColumn", attrs, children}) do
    %__MODULE__{
      id: attrs |> attr("id") |> String.to_integer(),
      name: attr(attrs, "name") || "",
      attrs: attrs,
      children: children
    }
  end

  @spec to_node(t()) :: node_tuple()
  def to_node(%__MODULE__{} = column) do
    attrs =
      column.attrs
      |> put_attr("id", Integer.to_string(column.id))
      |> put_attr("name", column.name)

    {"tableColumn", attrs, column.children}
  end

  @spec rename(t(), String.t()) :: t()
  def rename(%__MODULE__{} = column, name) when is_binary(name), do: %{column | name: name}

  @spec calculated_formula(t()) :: String.t() | nil
  def calculated_formula(%__MODULE__{children: children}) do
    child_text(children, "calculatedColumnFormula")
  end

  @spec totals_row_function(t()) :: totals_function() | nil
  def totals_row_function(%__MODULE__{attrs: attrs}) do
    case attr(attrs, "totalsRowFunction") do
      nil -> nil
      value -> @totals_functions |> Enum.find({nil, nil}, fn {_, v} -> v == value end) |> elem(0)
    end
  end

  @spec totals_row_label(t()) :: String.t() | nil
  def totals_row_label(%__MODULE__{attrs: attrs}), do: attr(attrs, "totalsRowLabel")

  @spec totals_row_formula(t()) :: String.t() | nil
  def totals_row_formula(%__MODULE__{children: children}) do
    child_text(children, "totalsRowFormula")
  end

  @doc """
  Sets the totals-row behaviour of the column. Accepts a
  `t:totals_function/0`, `{:custom, formula}`, `{:label, text}`, or
  `nil` to clear every totals-row attribute and child.
  """
  @spec put_totals(t(), totals_function() | {:custom, String.t()} | {:label, String.t()} | nil) ::
          t()
  def put_totals(%__MODULE__{} = column, nil), do: clear_totals(column)

  def put_totals(%__MODULE__{} = column, {:label, text}) when is_binary(text) do
    %{clear_totals(column) | attrs: put_attr(column.attrs, "totalsRowLabel", text)}
    |> drop_totals_attr("totalsRowFunction")
  end

  def put_totals(%__MODULE__{} = column, {:custom, formula}) when is_binary(formula) do
    cleared = clear_totals(column)

    %{
      cleared
      | attrs: put_attr(cleared.attrs, "totalsRowFunction", "custom"),
        children: cleared.children ++ [{"totalsRowFormula", [], [formula]}]
    }
  end

  def put_totals(%__MODULE__{} = column, function) when is_map_key(@totals_functions, function) do
    cleared = clear_totals(column)
    %{cleared | attrs: put_attr(cleared.attrs, "totalsRowFunction", @totals_functions[function])}
  end

  @doc "The SUBTOTAL function number Excel uses for a totals-row function."
  @spec subtotal_code(totals_function()) :: pos_integer() | nil
  def subtotal_code(:average), do: 101
  def subtotal_code(:count), do: 103
  def subtotal_code(:count_nums), do: 102
  def subtotal_code(:max), do: 104
  def subtotal_code(:min), do: 105
  def subtotal_code(:std_dev), do: 107
  def subtotal_code(:sum), do: 109
  def subtotal_code(:var), do: 110
  def subtotal_code(_), do: nil

  defp clear_totals(%__MODULE__{} = column) do
    column
    |> drop_totals_attr("totalsRowFunction")
    |> drop_totals_attr("totalsRowLabel")
    |> Map.update!(
      :children,
      &Enum.reject(&1, fn node -> match?({"totalsRowFormula", _, _}, node) end)
    )
  end

  defp drop_totals_attr(%__MODULE__{attrs: attrs} = column, key) do
    %{column | attrs: List.keydelete(attrs, key, 0)}
  end

  defp put_attr(attrs, key, value) do
    case List.keyfind(attrs, key, 0) do
      nil -> attrs ++ [{key, value}]
      _ -> List.keyreplace(attrs, key, 0, {key, value})
    end
  end

  defp attr(attrs, key) do
    case List.keyfind(attrs, key, 0) do
      {_, value} -> value
      nil -> nil
    end
  end

  defp child_text(children, tag) do
    case Enum.find(children, &match?({^tag, _, _}, &1)) do
      {^tag, _, inner} -> inner |> Enum.filter(&is_binary/1) |> Enum.join("")
      nil -> nil
    end
  end
end
