defmodule ExVEx.Formula.StructuredReference do
  @moduledoc """
  Rewrites the column and table names inside `:structured_ref` tokens.

  A structured reference body (the text between the outermost brackets
  of `Table1[...]`) is a comma or colon separated list of items. Each
  item is either a special selector (`#Headers`, `#Data`, `#Totals`,
  `#All`, `#This Row`), the this-row marker `@`, a column name written
  bare (`Amount`) or bracketed (`[Unit Price]`), or a nested bracket
  group. Inside a column name the characters `[`, `]`, `#` and `'` are
  escaped with a leading `'`.

  Only the named column or table is rewritten; every other character of
  the body is emitted exactly as it was read.
  """

  alias ExVEx.Formula.Token
  alias ExVEx.Utils.Coordinate

  @type item ::
          {:text, String.t()}
          | {:column, String.t(), :bare | :bracketed}
          | {:group, [item()]}

  @doc """
  Renames column `old` to `new` in every structured reference that
  targets `table`. Implicit references (`[@Col]`, no table prefix) are
  renamed only when `inside: true` is given, meaning the formula lives
  inside `table`'s own range. Matching is case-insensitive.
  """
  @spec rename_column([Token.t()], String.t(), String.t(), String.t(), keyword()) :: [Token.t()]
  def rename_column(tokens, table, old, new, opts \\ []) do
    inside? = Keyword.get(opts, :inside, false)
    Enum.map(tokens, &rename_column_in_token(&1, table, old, new, inside?))
  end

  @doc "Renames every qualified reference to table `old` so it targets `new`."
  @spec rename_table([Token.t()], String.t(), String.t()) :: [Token.t()]
  def rename_table(tokens, old, new) do
    Enum.map(tokens, fn
      %Token{kind: :structured_ref, table: table} = token when is_binary(table) ->
        if same_name?(table, old), do: %{token | table: new, text: nil}, else: token

      token ->
        token
    end)
  end

  @doc "Returns the column names a structured reference token mentions, in order."
  @spec column_names(Token.t()) :: [String.t()]
  def column_names(%Token{kind: :structured_ref, body: body}) do
    body |> parse_body() |> collect_columns([]) |> Enum.reverse()
  end

  @doc """
  Converts a structured reference token into the A1 range it denotes on
  `table`, or returns the token unchanged when it cannot be resolved.

  `sheet_prefix` is prepended to the range (pass `""` when the formula is
  on the table's own sheet). `formula_row` is the row of the cell holding
  the formula, used for `@` / `#This Row` references.
  """
  @spec to_range(Token.t(), map(), String.t(), pos_integer() | nil) :: Token.t()
  def to_range(
        %Token{kind: :structured_ref, body: body} = token,
        table,
        sheet_prefix,
        formula_row
      ) do
    items = parse_body(body)
    columns = items |> collect_columns([]) |> Enum.reverse()
    specials = collect_specials(items)

    with {:ok, {first_row, last_row}} <- row_span(specials, table, formula_row),
         {:ok, {first_col, last_col}} <- column_span(columns, table) do
      text = sheet_prefix <> range_text(first_row, first_col, last_row, last_col)
      %Token{kind: :literal, text: text}
    else
      :error -> token
    end
  end

  def to_range(token, _table, _sheet_prefix, _formula_row), do: token

  defp collect_specials(items) do
    Enum.flat_map(items, fn
      {:text, "@"} -> ["#this row"]
      {:text, "#" <> _ = special} -> [String.downcase(special)]
      {:group, inner} -> collect_specials(inner)
      _ -> []
    end)
  end

  defp row_span([], table, _formula_row), do: rows_or_error(table.data_rows)

  defp row_span(specials, table, formula_row) do
    spans = Enum.map(specials, &special_rows(&1, table, formula_row))

    if Enum.any?(spans, &(&1 == :error)) or spans == [] do
      :error
    else
      {:ok,
       {spans |> Enum.map(&elem(&1, 0)) |> Enum.min(),
        spans |> Enum.map(&elem(&1, 1)) |> Enum.max()}}
    end
  end

  defp rows_or_error(nil), do: :error
  defp rows_or_error(span), do: {:ok, span}

  defp special_rows("#all", table, _), do: {table.top, table.bottom}
  defp special_rows("#data", table, _), do: table.data_rows || :error

  defp special_rows("#headers", table, _),
    do: if(table.header_row, do: {table.header_row, table.header_row}, else: :error)

  defp special_rows("#totals", table, _),
    do: if(table.totals_row, do: {table.totals_row, table.totals_row}, else: :error)

  defp special_rows("#this row", _table, row) when is_integer(row), do: {row, row}
  defp special_rows(_, _, _), do: :error

  defp column_span([], table), do: {:ok, {table.left, table.right}}

  defp column_span(names, table) do
    indexes = Enum.map(names, &column_offset(table, &1))

    if Enum.any?(indexes, &is_nil/1) do
      :error
    else
      {:ok, {table.left + Enum.min(indexes), table.left + Enum.max(indexes)}}
    end
  end

  defp column_offset(table, name) do
    wanted = String.downcase(name)
    Enum.find_index(table.columns, &(String.downcase(&1) == wanted))
  end

  defp range_text(row, col, row, col), do: cell_text(row, col)
  defp range_text(r1, c1, r2, c2), do: cell_text(r1, c1) <> ":" <> cell_text(r2, c2)

  defp cell_text(row, col), do: Coordinate.to_string({row, col})

  defp rename_column_in_token(%Token{kind: :structured_ref} = token, table, old, new, inside?) do
    if targets_table?(token.table, table, inside?) do
      new_body = token.body |> parse_body() |> rename_items(old, new) |> serialize_items()
      if new_body == token.body, do: token, else: %{token | body: new_body, text: nil}
    else
      token
    end
  end

  defp rename_column_in_token(token, _table, _old, _new, _inside?), do: token

  defp targets_table?(nil, _table, inside?), do: inside?
  defp targets_table?(token_table, table, _inside?), do: same_name?(token_table, table)

  defp same_name?(a, b), do: String.downcase(a) == String.downcase(b)

  defp rename_items(items, old, new) do
    Enum.map(items, fn
      {:column, name, form} ->
        if same_name?(name, old), do: {:column, new, form}, else: {:column, name, form}

      {:group, inner} ->
        {:group, rename_items(inner, old, new)}

      other ->
        other
    end)
  end

  defp collect_columns(items, acc) do
    Enum.reduce(items, acc, fn
      {:column, name, _form}, acc -> [name | acc]
      {:group, inner}, acc -> collect_columns(inner, acc)
      _other, acc -> acc
    end)
  end

  @spec parse_body(String.t()) :: [item()]
  defp parse_body(body), do: body |> parse_items([]) |> Enum.reverse()

  defp parse_items(<<>>, acc), do: acc

  defp parse_items(<<?[, rest::binary>>, acc) do
    {inner, rest} = take_bracketed(rest, 1, <<>>)
    parse_items(rest, [bracket_item(inner) | acc])
  end

  defp parse_items(<<ch, rest::binary>>, acc) when ch in [?,, ?:, ?@] do
    parse_items(rest, [{:text, <<ch>>} | acc])
  end

  defp parse_items(<<?#, _::binary>> = input, acc) do
    {special, rest} = take_until_delimiter(input, <<>>)
    parse_items(rest, [{:text, special} | acc])
  end

  defp parse_items(input, acc) do
    {raw, rest} = take_until_delimiter(input, <<>>)
    parse_items(rest, [{:column, unescape(raw), :bare} | acc])
  end

  defp bracket_item(inner) do
    case parse_body(inner) do
      [{:column, name, :bare}] -> {:column, name, :bracketed}
      items -> {:group, items}
    end
  end

  defp take_bracketed(<<>>, _depth, buf), do: {buf, <<>>}

  defp take_bracketed(<<?', ch::utf8, rest::binary>>, depth, buf) do
    take_bracketed(rest, depth, buf <> <<?', ch::utf8>>)
  end

  defp take_bracketed(<<?], rest::binary>>, 1, buf), do: {buf, rest}

  defp take_bracketed(<<?], rest::binary>>, depth, buf),
    do: take_bracketed(rest, depth - 1, buf <> "]")

  defp take_bracketed(<<?[, rest::binary>>, depth, buf),
    do: take_bracketed(rest, depth + 1, buf <> "[")

  defp take_bracketed(<<ch::utf8, rest::binary>>, depth, buf) do
    take_bracketed(rest, depth, buf <> <<ch::utf8>>)
  end

  defp take_until_delimiter(<<>>, buf), do: {buf, <<>>}

  defp take_until_delimiter(<<?', ch::utf8, rest::binary>>, buf) do
    take_until_delimiter(rest, buf <> <<?', ch::utf8>>)
  end

  defp take_until_delimiter(<<ch, _::binary>> = input, buf) when ch in [?,, ?:, ?[, ?]] do
    {buf, input}
  end

  defp take_until_delimiter(<<ch::utf8, rest::binary>>, buf) do
    take_until_delimiter(rest, buf <> <<ch::utf8>>)
  end

  defp serialize_items(items) do
    Enum.map_join(items, "", fn
      {:text, text} -> text
      {:column, name, :bare} -> serialize_bare(name)
      {:column, name, :bracketed} -> "[" <> escape(name) <> "]"
      {:group, inner} -> "[" <> serialize_items(inner) <> "]"
    end)
  end

  defp serialize_bare(name) do
    if Regex.match?(~r/^[A-Za-z0-9_][A-Za-z0-9_ ]*[A-Za-z0-9_]$|^[A-Za-z0-9_]$/, name) do
      name
    else
      "[" <> escape(name) <> "]"
    end
  end

  @escaped [?[, ?], ?#, ?']

  defp escape(name) do
    for <<ch::utf8 <- name>>, into: "" do
      if ch in @escaped, do: <<?', ch::utf8>>, else: <<ch::utf8>>
    end
  end

  defp unescape(raw), do: unescape(raw, <<>>)

  defp unescape(<<>>, acc), do: acc
  defp unescape(<<?', ch::utf8, rest::binary>>, acc), do: unescape(rest, acc <> <<ch::utf8>>)
  defp unescape(<<ch::utf8, rest::binary>>, acc), do: unescape(rest, acc <> <<ch::utf8>>)
end
