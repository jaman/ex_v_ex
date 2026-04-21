defmodule ExVEx.Formula.Tokenizer do
  @moduledoc """
  Tokenises an Excel formula string into a flat list of
  `ExVEx.Formula.Token` records.

  The goal isn't to build an AST — Excel formulas are too complex for
  that in a pure-Elixir budget — but to flatten the input into a stream
  where cell and range references are structured and everything else
  passes through as literal text. That's enough to rewrite references
  on row/column insert without understanding operator precedence or
  function semantics.

  Supported reference forms:

    * Plain cells:            `A1`, `$A$1`, `A$1`, `$A1`
    * Ranges:                 `A1:B5`, `$A$1:$B$5`
    * Row ranges:             `1:5`, `$1:5`
    * Column ranges:          `A:C`
    * Sheet prefix:           `Sheet1!A1`, `'Data Sheet'!A1`
    * 3D sheet spans:         `Sheet1:Sheet3!A1`

  String literals (`"text"`) and quoted sheet names (`'Sheet Name'`)
  are preserved verbatim.
  """

  alias ExVEx.Formula.{Reference, Token}

  @spec tokenize(String.t()) :: [Token.t()]
  def tokenize(formula) when is_binary(formula) do
    formula |> tokenize(<<>>, []) |> coalesce_literals()
  end

  defp tokenize(<<>>, <<>>, acc), do: Enum.reverse(acc)

  defp tokenize(<<>>, literal, acc) do
    Enum.reverse([%Token{kind: :literal, text: literal} | acc])
  end

  # String literal "…"
  defp tokenize(<<?", rest::binary>>, literal, acc) do
    {string, rest} = consume_string(rest, <<"\"">>)
    acc = flush_literal(literal, acc)
    tokenize(rest, <<>>, [%Token{kind: :literal, text: string} | acc])
  end

  # Quoted sheet name 'Sheet Name'! — but only as a reference prefix,
  # not as a standalone literal. We detect by checking the ! follows.
  defp tokenize(<<?', _::binary>> = input, literal, acc) do
    case parse_quoted_sheet_ref(input) do
      {:ok, token, rest} ->
        acc = flush_literal(literal, acc)
        tokenize(rest, <<>>, [token | acc])

      :error ->
        # No valid reference after the quoted sheet, treat the quote as a
        # literal char and continue
        <<ch::utf8, rest::binary>> = input
        tokenize(rest, literal <> <<ch::utf8>>, acc)
    end
  end

  defp tokenize(<<ch::utf8, _::binary>> = input, literal, acc)
       when ch in ?A..?Z or ch in ?a..?z or ch == ?$ or ch in ?0..?9 do
    case parse_reference_or_sheet_prefixed(input) do
      {:ok, token, rest} ->
        acc = flush_literal(literal, acc)
        tokenize(rest, <<>>, [token | acc])

      :error ->
        <<c, rest::binary>> = input
        tokenize(rest, literal <> <<c>>, acc)
    end
  end

  defp tokenize(<<ch::utf8, rest::binary>>, literal, acc) do
    tokenize(rest, literal <> <<ch::utf8>>, acc)
  end

  defp flush_literal(<<>>, acc), do: acc
  defp flush_literal(text, acc), do: [%Token{kind: :literal, text: text} | acc]

  defp consume_string(<<?", ?", rest::binary>>, buf), do: consume_string(rest, buf <> <<"\"\"">>)
  defp consume_string(<<?", rest::binary>>, buf), do: {buf <> <<"\"">>, rest}

  defp consume_string(<<ch::utf8, rest::binary>>, buf),
    do: consume_string(rest, buf <> <<ch::utf8>>)

  defp consume_string(<<>>, buf), do: {buf, <<>>}

  # Parses 'Sheet Name'! or 'S1:S3'! followed by a reference
  defp parse_quoted_sheet_ref(input) do
    case Regex.run(~r/^'((?:[^']|'')+)'!/, input) do
      [sheet_prefix, sheet_inner] ->
        sheet = String.replace(sheet_inner, "''", "'")
        rest = String.slice(input, String.length(sheet_prefix)..-1//1)
        parse_ref_after_sheet(sheet, rest, sheet_prefix)

      nil ->
        :error
    end
  end

  # Parses Sheet1!A1 or Sheet1:Sheet3!A1 or bare A1
  defp parse_reference_or_sheet_prefixed(input) do
    case Regex.run(~r/^([A-Za-z_][A-Za-z0-9_.]*(?::[A-Za-z_][A-Za-z0-9_.]*)?)!/, input) do
      [sheet_prefix, sheet_name] ->
        rest = String.slice(input, String.length(sheet_prefix)..-1//1)
        parse_ref_after_sheet(sheet_name, rest, sheet_prefix)

      nil ->
        parse_bare_reference(input, nil, "")
    end
  end

  defp parse_ref_after_sheet(sheet, rest, sheet_prefix) do
    case parse_bare_reference(rest, sheet, sheet_prefix) do
      {:ok, _, _} = ok ->
        ok

      :error ->
        :error
    end
  end

  defp parse_bare_reference(input, sheet, sheet_prefix) do
    cond do
      match = Regex.run(~r/^(\$?[A-Za-z]+\$?[0-9]+):(\$?[A-Za-z]+\$?[0-9]+)/, input) ->
        [whole, a_text, b_text] = match

        with {:ok, a} <- Reference.parse(a_text),
             {:ok, b} <- Reference.parse(b_text) do
          rest = String.slice(input, String.length(whole)..-1//1)

          {:ok,
           %Token{
             kind: :range_ref,
             sheet: sheet,
             start_ref: a,
             end_ref: b,
             text: sheet_prefix <> whole
           }, rest}
        else
          _ -> :error
        end

      match = Regex.run(~r/^(\$?[A-Za-z]+\$?[0-9]+)/, input) ->
        [_, ref_text] = match

        case Reference.parse(ref_text) do
          {:ok, ref} ->
            rest = String.slice(input, String.length(ref_text)..-1//1)

            {:ok, %Token{kind: :cell_ref, sheet: sheet, ref: ref, text: sheet_prefix <> ref_text},
             rest}

          :error ->
            :error
        end

      match = Regex.run(~r/^(\$?)([0-9]+):(\$?)([0-9]+)/, input) ->
        [whole, sa, sr, ea, er] = match
        rest = String.slice(input, String.length(whole)..-1//1)

        {:ok,
         %Token{
           kind: :row_range,
           sheet: sheet,
           start_row: String.to_integer(sr),
           start_abs?: sa == "$",
           end_row: String.to_integer(er),
           end_abs?: ea == "$",
           text: sheet_prefix <> whole
         }, rest}

      match = Regex.run(~r/^(\$?)([A-Za-z]+):(\$?)([A-Za-z]+)/, input) ->
        [whole, sa, sl, ea, el] = match
        rest = String.slice(input, String.length(whole)..-1//1)

        {:ok,
         %Token{
           kind: :col_range,
           sheet: sheet,
           start_col: col_num(sl),
           start_abs?: sa == "$",
           end_col: col_num(el),
           end_abs?: ea == "$",
           text: sheet_prefix <> whole
         }, rest}

      true ->
        :error
    end
  end

  defp col_num(letters) do
    letters
    |> String.upcase()
    |> :erlang.binary_to_list()
    |> Enum.reduce(0, fn ch, acc -> acc * 26 + (ch - ?A + 1) end)
  end

  defp coalesce_literals(tokens), do: coalesce_literals(tokens, [])

  defp coalesce_literals([], acc), do: Enum.reverse(acc)

  defp coalesce_literals(
         [%Token{kind: :literal, text: a}, %Token{kind: :literal, text: b} | rest],
         acc
       ) do
    coalesce_literals([%Token{kind: :literal, text: a <> b} | rest], acc)
  end

  defp coalesce_literals([t | rest], acc), do: coalesce_literals(rest, [t | acc])
end
