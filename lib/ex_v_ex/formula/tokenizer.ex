defmodule ExVEx.Formula.Tokenizer do
  @moduledoc """
  Tokenises an Excel formula string into a flat list of
  `ExVEx.Formula.Token` records.

  The output is not an AST: cell, range, and structured references are
  emitted as structured tokens and everything else passes through as
  literal text. That is enough to rewrite references on row/column
  insert without understanding operator precedence or function
  semantics.

  Supported reference forms:

    * Plain cells:            `A1`, `$A$1`, `A$1`, `$A1`
    * Ranges:                 `A1:B5`, `$A$1:$B$5`
    * Row ranges:             `1:5`, `$1:5`
    * Column ranges:          `A:C`
    * Sheet prefix:           `Sheet1!A1`, `'Data Sheet'!A1`
    * 3D sheet spans:         `Sheet1:Sheet3!A1`
    * Structured references:  `Table1[Amount]`, `[@Price]`,
                              `Sales[[#Headers],[Q4]]`

  An identifier that merely resembles a reference — a function name
  such as `LOG10(`, a defined name such as `Table1`, or a dotted name
  such as `My.A1` — is emitted as a literal, never as a cell reference.
  String literals (`"text"`) and quoted sheet names (`'Sheet Name'`)
  are preserved verbatim.
  """

  alias ExVEx.Formula.{Reference, Token}

  defguardp identifier_start?(ch)
            when ch in ?A..?Z or ch in ?a..?z or ch == ?$ or ch in ?0..?9 or ch == ?_

  defguardp identifier_char?(ch)
            when ch in ?A..?Z or ch in ?a..?z or ch in ?0..?9 or ch == ?_ or ch == ?. or
                   ch == ?$

  @spec tokenize(String.t()) :: [Token.t()]
  def tokenize(formula) when is_binary(formula) do
    formula |> tokenize(<<>>, []) |> coalesce_literals()
  end

  defp tokenize(<<>>, <<>>, acc), do: Enum.reverse(acc)

  defp tokenize(<<>>, literal, acc) do
    Enum.reverse([%Token{kind: :literal, text: literal} | acc])
  end

  defp tokenize(<<?", rest::binary>>, literal, acc) do
    {string, rest} = consume_string(rest, <<"\"">>)
    acc = flush_literal(literal, acc)
    tokenize(rest, <<>>, [%Token{kind: :literal, text: string} | acc])
  end

  defp tokenize(<<?', _::binary>> = input, literal, acc) do
    case parse_quoted_sheet_ref(input) do
      {:ok, token, rest} ->
        acc = flush_literal(literal, acc)
        tokenize(rest, <<>>, [token | acc])

      :error ->
        <<ch::utf8, rest::binary>> = input
        tokenize(rest, literal <> <<ch::utf8>>, acc)
    end
  end

  defp tokenize(<<?[, _::binary>> = input, literal, acc) do
    case consume_bracket(input) do
      {:ok, body, rest} ->
        acc = flush_literal(literal, acc)
        tokenize(rest, <<>>, [structured_ref(nil, body) | acc])

      :error ->
        <<_, rest::binary>> = input
        tokenize(rest, literal <> "[", acc)
    end
  end

  defp tokenize(<<ch::utf8, _::binary>> = input, literal, acc) when identifier_start?(ch) do
    case parse_reference_or_sheet_prefixed(input) do
      {:ok, token, rest} ->
        accept_or_fall_back(token, rest, input, literal, acc)

      :error ->
        tokenize_identifier(input, literal, acc)
    end
  end

  defp tokenize(<<ch::utf8, rest::binary>>, literal, acc) do
    tokenize(rest, literal <> <<ch::utf8>>, acc)
  end

  defp accept_or_fall_back(_token, <<ch::utf8, _::binary>>, input, literal, acc)
       when identifier_char?(ch) or ch == ?( or ch == ?[ do
    tokenize_identifier(input, literal, acc)
  end

  defp accept_or_fall_back(token, rest, _input, literal, acc) do
    acc = flush_literal(literal, acc)
    tokenize(rest, <<>>, [token | acc])
  end

  defp tokenize_identifier(input, literal, acc) do
    {name, rest} = consume_identifier(input, <<>>)

    case consume_bracket(rest) do
      {:ok, body, rest_after_bracket} ->
        acc = flush_literal(literal, acc)
        tokenize(rest_after_bracket, <<>>, [structured_ref(name, body) | acc])

      :error ->
        tokenize(rest, literal <> name, acc)
    end
  end

  defp consume_identifier(<<ch::utf8, rest::binary>>, buf) when identifier_char?(ch) do
    consume_identifier(rest, buf <> <<ch::utf8>>)
  end

  defp consume_identifier(rest, buf), do: {buf, rest}

  defp structured_ref(table, body) do
    %Token{
      kind: :structured_ref,
      table: table,
      body: body,
      text: (table || "") <> "[" <> body <> "]"
    }
  end

  defp consume_bracket(<<?[, rest::binary>>), do: consume_bracket(rest, 1, <<>>)
  defp consume_bracket(_), do: :error

  defp consume_bracket(<<>>, _depth, _buf), do: :error

  defp consume_bracket(<<?', ch::utf8, rest::binary>>, depth, buf) do
    consume_bracket(rest, depth, buf <> <<?', ch::utf8>>)
  end

  defp consume_bracket(<<?], rest::binary>>, 1, buf), do: {:ok, buf, rest}

  defp consume_bracket(<<?], rest::binary>>, depth, buf) do
    consume_bracket(rest, depth - 1, buf <> "]")
  end

  defp consume_bracket(<<?[, rest::binary>>, depth, buf) do
    consume_bracket(rest, depth + 1, buf <> "[")
  end

  defp consume_bracket(<<ch::utf8, rest::binary>>, depth, buf) do
    consume_bracket(rest, depth, buf <> <<ch::utf8>>)
  end

  defp flush_literal(<<>>, acc), do: acc
  defp flush_literal(text, acc), do: [%Token{kind: :literal, text: text} | acc]

  defp consume_string(<<?", ?", rest::binary>>, buf), do: consume_string(rest, buf <> <<"\"\"">>)
  defp consume_string(<<?", rest::binary>>, buf), do: {buf <> <<"\"">>, rest}

  defp consume_string(<<ch::utf8, rest::binary>>, buf),
    do: consume_string(rest, buf <> <<ch::utf8>>)

  defp consume_string(<<>>, buf), do: {buf, <<>>}

  defp parse_quoted_sheet_ref(input) do
    case Regex.run(~r/^'((?:[^']|'')+)'!/, input) do
      [sheet_prefix, sheet_inner] ->
        sheet = String.replace(sheet_inner, "''", "'")
        rest = String.slice(input, String.length(sheet_prefix)..-1//1)
        parse_bare_reference(rest, sheet, sheet_prefix)

      nil ->
        :error
    end
  end

  defp parse_reference_or_sheet_prefixed(input) do
    case Regex.run(~r/^([A-Za-z_][A-Za-z0-9_.]*(?::[A-Za-z_][A-Za-z0-9_.]*)?)!/, input) do
      [sheet_prefix, sheet_name] ->
        rest = String.slice(input, String.length(sheet_prefix)..-1//1)
        parse_bare_reference(rest, sheet_name, sheet_prefix)

      nil ->
        parse_bare_reference(input, nil, "")
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
