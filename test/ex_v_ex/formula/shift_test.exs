defmodule ExVEx.Formula.ShiftTest do
  use ExUnit.Case, async: true

  alias ExVEx.Formula.Serializer
  alias ExVEx.Formula.Shift, as: FormulaShift
  alias ExVEx.Formula.Tokenizer
  alias ExVEx.Mutation.Shift

  defp rewrite(formula, shift, current_sheet \\ nil) do
    formula
    |> Tokenizer.tokenize()
    |> FormulaShift.apply(shift, current_sheet)
    |> Serializer.to_string()
  end

  describe "insert rows" do
    test "cell reference at the insert point shifts down by 1" do
      shift = Shift.insert(:row, 3, 1, nil)
      assert "=A4" == rewrite("=A3", shift)
    end

    test "cell reference far below the insert shifts by count" do
      shift = Shift.insert(:row, 3, 2, nil)
      assert "=A12" == rewrite("=A10", shift)
    end

    test "cell reference above the insert is unchanged" do
      shift = Shift.insert(:row, 5, 1, nil)
      assert "=A3" == rewrite("=A3", shift)
    end

    test "absolute row reference still shifts on structural insert" do
      shift = Shift.insert(:row, 3, 1, nil)
      assert "=$A$6" == rewrite("=$A$5", shift)
    end

    test "range expands down" do
      shift = Shift.insert(:row, 3, 2, nil)
      assert "=SUM(A1:A12)" == rewrite("=SUM(A1:A10)", shift)
    end

    test "row range expands" do
      shift = Shift.insert(:row, 3, 5, nil)
      assert "=SUM(1:10)" == rewrite("=SUM(1:5)", shift)
    end

    test "column range unaffected by row insert" do
      shift = Shift.insert(:row, 3, 2, nil)
      assert "=SUM(A:A)" == rewrite("=SUM(A:A)", shift)
    end

    test "cross-sheet reference shifts only when the shift targets that sheet" do
      shift = Shift.insert(:row, 3, 1, "Data")

      # Same sheet: shifts
      assert "=Data!A4" == rewrite("=Data!A3", shift)

      # Different sheet: unchanged
      assert "=Other!A3" == rewrite("=Other!A3", shift)
    end

    test "sheet-less reference shifts when shift is workbook-wide" do
      shift = Shift.insert(:row, 3, 1, nil)
      assert "=A4" == rewrite("=A3", shift)
    end

    test "sheet-less reference shifts when current sheet matches the shift" do
      shift = Shift.insert(:row, 3, 1, "Sheet1")
      assert "=A4" == rewrite("=A3", shift, "Sheet1")
    end

    test "sheet-less reference NOT shifted when current sheet differs" do
      shift = Shift.insert(:row, 3, 1, "Sheet2")
      assert "=A3" == rewrite("=A3", shift, "Sheet1")
    end
  end

  describe "insert columns" do
    test "cell reference at the insert column shifts" do
      shift = Shift.insert(:col, 3, 1, nil)
      assert "=D1" == rewrite("=C1", shift)
    end

    test "cell reference left of the insert is unchanged" do
      shift = Shift.insert(:col, 5, 1, nil)
      assert "=B1" == rewrite("=B1", shift)
    end

    test "column range expands" do
      shift = Shift.insert(:col, 3, 2, nil)
      assert "=SUM(A:E)" == rewrite("=SUM(A:C)", shift)
    end
  end

  describe "delete rows" do
    test "references inside the deletion span become #REF!" do
      shift = Shift.delete(:row, 3, 2, nil)
      assert "=#REF!" == rewrite("=A4", shift)
    end

    test "references after the deletion shift up" do
      shift = Shift.delete(:row, 3, 2, nil)
      assert "=A4" == rewrite("=A6", shift)
    end

    test "references before the deletion unchanged" do
      shift = Shift.delete(:row, 3, 2, nil)
      assert "=A2" == rewrite("=A2", shift)
    end
  end

  describe "mixed / edge cases" do
    test "string literals inside a formula are not mistaken for refs" do
      shift = Shift.insert(:row, 3, 1, nil)
      formula = ~S|=IF(A1="hello",SUM(B2:B10),0)|
      # A1 is above the insert so it stays; B2:B10 shifts to B2:B11.
      assert ~S|=IF(A1="hello",SUM(B2:B11),0)| == rewrite(formula, shift)
    end

    test "cross-workbook refs not broken on shifts to our workbook" do
      shift = Shift.insert(:row, 3, 1, nil)
      formula = "=[Book1.xlsx]Sheet1!A1"
      # Non-trivial — our tokenizer might parse Sheet1!A1. A1 < at(3), unchanged.
      assert "=[Book1.xlsx]Sheet1!A1" == rewrite(formula, shift)
    end
  end
end
