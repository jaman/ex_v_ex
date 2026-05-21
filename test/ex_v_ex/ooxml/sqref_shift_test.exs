defmodule ExVEx.OOXML.SqrefShiftTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.SqrefShift

  test "shifts a single range on row insert" do
    shift = Shift.insert(:row, 3, 2, nil)
    assert "A5:B12" == SqrefShift.shift("A3:B10", shift)
  end

  test "shifts space-separated ranges independently" do
    shift = Shift.insert(:row, 5, 1, nil)
    assert "A1:A4 A6:A11 B8" == SqrefShift.shift("A1:A4 A5:A10 B7", shift)
  end

  test "drops tokens fully inside a deletion span" do
    shift = Shift.delete(:row, 3, 3, nil)
    assert "A7" == SqrefShift.shift("A4 A10", shift)
  end

  test "shifts single cell tokens" do
    shift = Shift.insert(:col, 2, 1, nil)
    assert "C1 D1" == SqrefShift.shift("B1 C1", shift)
  end
end
