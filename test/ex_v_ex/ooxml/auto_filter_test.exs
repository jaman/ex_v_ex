defmodule ExVEx.OOXML.AutoFilterTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.AutoFilter

  test "shifts the ref attribute on row insert" do
    node = {"autoFilter", [{"ref", "A1:C10"}], []}
    shift = Shift.insert(:row, 5, 2, nil)
    {"autoFilter", [{"ref", "A1:C12"}], []} = AutoFilter.shift_node(node, shift)
  end

  test "leaves the node untouched for shifts below the range" do
    node = {"autoFilter", [{"ref", "A1:C10"}], []}
    shift = Shift.insert(:row, 50, 1, nil)
    assert node == AutoFilter.shift_node(node, shift)
  end
end
