defmodule ExVEx.Formula.StructuredReferenceTokenizerTest do
  use ExUnit.Case, async: true

  alias ExVEx.Formula.{Serializer, Shift, Token, Tokenizer}
  alias ExVEx.Mutation.Shift, as: MutShift

  describe "identifiers that look like references" do
    test "a function name with trailing digits is a literal" do
      assert [
               %Token{kind: :literal, text: "=LOG10("},
               %Token{kind: :cell_ref},
               %Token{kind: :literal, text: ")"}
             ] =
               Tokenizer.tokenize("=LOG10(A1)")
    end

    test "a defined name with an embedded cell-looking suffix is a literal" do
      assert [%Token{kind: :literal, text: "=Table1"}] = Tokenizer.tokenize("=Table1")
    end

    test "a name followed by a dot-qualified segment is a literal" do
      assert [%Token{kind: :literal, text: "=My.A1"}] = Tokenizer.tokenize("=My.A1")
    end
  end

  describe "structured references" do
    test "table-qualified single column" do
      assert [
               %Token{kind: :literal, text: "=SUM("},
               %Token{
                 kind: :structured_ref,
                 table: "Table1",
                 body: "Amount",
                 text: "Table1[Amount]"
               },
               %Token{kind: :literal, text: ")"}
             ] = Tokenizer.tokenize("=SUM(Table1[Amount])")
    end

    test "implicit this-row reference has no table" do
      assert [
               %Token{kind: :literal, text: "="},
               %Token{kind: :structured_ref, table: nil, body: "@Price"},
               %Token{kind: :literal, text: "*"},
               %Token{kind: :structured_ref, table: nil, body: "@Qty"}
             ] = Tokenizer.tokenize("=[@Price]*[@Qty]")
    end

    test "nested brackets stay inside one token" do
      assert [
               %Token{kind: :literal, text: "="},
               %Token{kind: :structured_ref, table: "Sales", body: "[#Headers],[Q4]"}
             ] = Tokenizer.tokenize("=Sales[[#Headers],[Q4]]")
    end

    test "escaped closing bracket inside a column name does not end the token" do
      assert [
               %Token{kind: :literal, text: "="},
               %Token{kind: :structured_ref, table: "T", body: "Col']1"}
             ] = Tokenizer.tokenize("=T[Col']1]")
    end
  end

  describe "shift leaves structured references untouched" do
    for formula <- [
          "=SUM(Table1[Amount])",
          "=SUM(Table1[Q4])",
          "=[@Price]*[@Qty]",
          "=Sales[[#Headers],[A1]]",
          "=LOG10(B2)+Table1[Total]"
        ] do
      test "#{formula} on row insert at 1" do
        formula = unquote(formula)
        mut = MutShift.insert(:row, 1, 3, nil)
        tokens = Tokenizer.tokenize(formula)
        shifted = tokens |> Shift.apply(mut) |> Serializer.to_string()

        expected =
          case formula do
            "=LOG10(B2)+Table1[Total]" -> "=LOG10(B5)+Table1[Total]"
            other -> other
          end

        assert shifted == expected
      end
    end
  end

  describe "round-trip" do
    for formula <- [
          "=SUM(Table1[Amount])",
          "=[@Price]*[@Qty]",
          "=Sales[[#Headers],[Q4]]",
          "=T[Col']1]",
          "=SUBTOTAL(109,Sales[Amount])",
          "=LOG10(A1)",
          "=Table1[[#This Row],[Amount]]"
        ] do
      test "round-trips #{inspect(formula)}" do
        assert Serializer.to_string(Tokenizer.tokenize(unquote(formula))) == unquote(formula)
      end
    end
  end
end
