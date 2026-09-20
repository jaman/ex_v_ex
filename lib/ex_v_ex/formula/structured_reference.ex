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
