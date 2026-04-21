defmodule ExVEx.Formula.TokenizerTest do
  use ExUnit.Case, async: true

  alias ExVEx.Formula.Reference
  alias ExVEx.Formula.Serializer
  alias ExVEx.Formula.Token
  alias ExVEx.Formula.Tokenizer

  describe "tokenize/1" do
    test "plain cell reference" do
      assert [%Token{kind: :cell_ref, ref: %Reference{row: 1, col: 1}, sheet: nil}] =
               Tokenizer.tokenize("A1")
    end

    test "absolute reference keeps $ markers" do
      assert [
               %Token{
                 kind: :cell_ref,
                 ref: %Reference{row: 1, col: 1, row_abs?: true, col_abs?: true}
               }
             ] = Tokenizer.tokenize("$A$1")
    end

    test "mixed abs/rel" do
      assert [
               %Token{
                 kind: :cell_ref,
                 ref: %Reference{row_abs?: false, col_abs?: true}
               }
             ] = Tokenizer.tokenize("$A1")
    end

    test "range reference" do
      tokens = Tokenizer.tokenize("A1:B10")
      assert [%Token{kind: :range_ref, start_ref: a, end_ref: b}] = tokens
      assert a.row == 1 and a.col == 1
      assert b.row == 10 and b.col == 2
    end

    test "row range" do
      assert [%Token{kind: :row_range, start_row: 1, end_row: 5}] =
               Tokenizer.tokenize("1:5")
    end

    test "column range" do
      assert [%Token{kind: :col_range, start_col: 1, end_col: 3}] =
               Tokenizer.tokenize("A:C")
    end

    test "sheet-qualified reference" do
      assert [%Token{kind: :cell_ref, sheet: "Sheet1"}] = Tokenizer.tokenize("Sheet1!A1")
    end

    test "quoted sheet name with spaces" do
      assert [%Token{kind: :cell_ref, sheet: "Data Sheet"}] =
               Tokenizer.tokenize("'Data Sheet'!A1")
    end

    test "3D sheet span" do
      assert [%Token{kind: :range_ref, sheet: "Sheet1:Sheet3"}] =
               Tokenizer.tokenize("Sheet1:Sheet3!A1:B10")
    end

    test "=A1+B1 (simple formula with operator)" do
      tokens = Tokenizer.tokenize("=A1+B1")

      assert [
               %Token{kind: :literal, text: "="},
               %Token{kind: :cell_ref, ref: %Reference{row: 1, col: 1}},
               %Token{kind: :literal, text: "+"},
               %Token{kind: :cell_ref, ref: %Reference{row: 1, col: 2}}
             ] = tokens
    end

    test "SUM call with range" do
      tokens = Tokenizer.tokenize("=SUM(A1:A10)")
      assert Enum.any?(tokens, &match?(%Token{kind: :range_ref}, &1))
      assert Enum.any?(tokens, &(&1.kind == :literal and String.contains?(&1.text, "SUM")))
    end

    test "nested function with string literal" do
      formula = ~S|=IF(A1="hello",SUM(B:B),0)|
      tokens = Tokenizer.tokenize(formula)
      assert Enum.any?(tokens, &match?(%Token{kind: :cell_ref}, &1))
      assert Enum.any?(tokens, &match?(%Token{kind: :col_range}, &1))

      rebuilt = Serializer.to_string(tokens)
      assert rebuilt =~ ~S|"hello"|
    end

    test "cross-workbook reference passes through without crashing" do
      # We don't model cross-workbook refs; the tokenizer's contract is
      # just that the formula survives tokenize + serialize round-trip.
      formula = "=[Book1.xlsx]Sheet1!A1"
      tokens = Tokenizer.tokenize(formula)
      assert Serializer.to_string(tokens) == formula
    end

    test "tokenizing empty string gives an empty token list" do
      assert [] = Tokenizer.tokenize("")
    end
  end

  describe "round-trip via serializer" do
    formulas = [
      "=A1",
      "=$A$1",
      "=A1+B1",
      "=SUM(A1:A10)",
      "=Sheet1!A1",
      "='Data Sheet'!A1",
      ~S|=IF(A1="hello",SUM(B:B),0)|,
      "=SUM(A1:A10)+AVERAGE(B1:B10)",
      "=SUM(Sheet1:Sheet3!A1:B10)",
      "=1:5",
      "=A:C"
    ]

    for formula <- formulas do
      test "round-trips #{inspect(formula)}" do
        tokens = Tokenizer.tokenize(unquote(formula))
        assert Serializer.to_string(tokens) == unquote(formula)
      end
    end
  end
end
