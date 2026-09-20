defmodule ExVEx.Formula.StructuredReferenceTest do
  use ExUnit.Case, async: true

  alias ExVEx.Formula.{Serializer, StructuredReference, Tokenizer}

  defp rename_column(formula, table, old, new, opts \\ []) do
    formula
    |> Tokenizer.tokenize()
    |> StructuredReference.rename_column(table, old, new, opts)
    |> Serializer.to_string()
  end

  defp rename_table(formula, old, new) do
    formula
    |> Tokenizer.tokenize()
    |> StructuredReference.rename_table(old, new)
    |> Serializer.to_string()
  end

  describe "rename_column/5" do
    test "single bare column" do
      assert rename_column("=SUM(Sales[Amount])", "Sales", "Amount", "Total") ==
               "=SUM(Sales[Total])"
    end

    test "column inside a multi-item list" do
      assert rename_column("=Sales[[#Headers],[Amount]]", "Sales", "Amount", "Total") ==
               "=Sales[[#Headers],[Total]]"
    end

    test "column range renames both ends independently" do
      assert rename_column("=SUM(Sales[[Qty]:[Amount]])", "Sales", "Amount", "Total") ==
               "=SUM(Sales[[Qty]:[Total]])"
    end

    test "this-row shorthand" do
      assert rename_column("=Sales[@Amount]*2", "Sales", "Amount", "Total") ==
               "=Sales[@Total]*2"
    end

    test "this-row with bracketed column" do
      assert rename_column("=Sales[@[Unit Price]]", "Sales", "Unit Price", "Price") ==
               "=Sales[@[Price]]"
    end

    test "implicit reference is renamed only when the formula sits inside the table" do
      assert rename_column("=[@Amount]*[@Qty]", "Sales", "Amount", "Total", inside: true) ==
               "=[@Total]*[@Qty]"

      assert rename_column("=[@Amount]*[@Qty]", "Sales", "Amount", "Total", inside: false) ==
               "=[@Amount]*[@Qty]"
    end

    test "other tables are untouched" do
      assert rename_column("=Sales[Amount]+Costs[Amount]", "Sales", "Amount", "Total") ==
               "=Sales[Total]+Costs[Amount]"
    end

    test "matching is case-insensitive" do
      assert rename_column("=Sales[amount]", "Sales", "Amount", "Total") == "=Sales[Total]"
      assert rename_column("=sales[Amount]", "Sales", "Amount", "Total") == "=sales[Total]"
    end

    test "special items are never treated as columns" do
      assert rename_column("=Sales[[#Totals],[Amount]]", "Sales", "#Totals", "X") ==
               "=Sales[[#Totals],[Amount]]"
    end

    test "new name with special characters is escaped and bracketed" do
      assert rename_column("=Sales[Amount]", "Sales", "Amount", "Total [USD]") ==
               "=Sales[[Total '[USD']]]"
    end

    test "escaped old name is matched after unescaping" do
      assert rename_column("=Sales[Total '[USD']]", "Sales", "Total [USD]", "Amount") ==
               "=Sales[Amount]"
    end

    test "non-structured tokens pass through" do
      assert rename_column("=A1+Sales[Amount]", "Sales", "Amount", "Total") ==
               "=A1+Sales[Total]"
    end
  end

  describe "rename_table/3" do
    test "renames qualified references" do
      assert rename_table("=SUM(Sales[Amount])+Costs[Amount]", "Sales", "Revenue") ==
               "=SUM(Revenue[Amount])+Costs[Amount]"
    end

    test "leaves implicit references alone" do
      assert rename_table("=[@Amount]", "Sales", "Revenue") == "=[@Amount]"
    end
  end

  describe "column_names/1" do
    test "lists every column a token refers to" do
      [_, token] = Tokenizer.tokenize("=Sales[[#Headers],[Amount]:[Qty]]")
      assert StructuredReference.column_names(token) == ["Amount", "Qty"]
    end
  end
end
