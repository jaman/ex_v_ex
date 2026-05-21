defmodule ExVEx.OOXML.ConditionalFormattingTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.ConditionalFormatting

  test "shifts sqref and formula children on row insert" do
    node =
      {"conditionalFormatting", [{"sqref", "A5:A20"}],
       [
         {"cfRule", [{"type", "expression"}, {"priority", "1"}],
          [{"formula", [], ["A5>SUM($B$5:$B$10)"]}]}
       ]}

    shift = Shift.insert(:row, 3, 2, "Sheet1")

    {"conditionalFormatting", attrs, [{"cfRule", _, [{"formula", _, [text]}]}]} =
      ConditionalFormatting.shift_node(node, shift, "Sheet1")

    assert {_, "A7:A22"} = List.keyfind(attrs, "sqref", 0)
    assert text == "A7>SUM($B$7:$B$12)"
  end

  test "leaves unrelated nodes unchanged when shift is above the range" do
    node =
      {"conditionalFormatting", [{"sqref", "A1:A2"}],
       [{"cfRule", [], [{"formula", [], ["A1>0"]}]}]}

    shift = Shift.insert(:row, 10, 1, "Sheet1")
    assert node == ConditionalFormatting.shift_node(node, shift, "Sheet1")
  end
end
