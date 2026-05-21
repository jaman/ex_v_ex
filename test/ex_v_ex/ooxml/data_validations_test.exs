defmodule ExVEx.OOXML.DataValidationsTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.DataValidations

  test "shifts sqref and formula1/formula2 children on row insert" do
    node =
      {"dataValidations", [{"count", "1"}],
       [
         {"dataValidation", [{"type", "list"}, {"sqref", "A5:A20"}],
          [
            {"formula1", [], ["$B$5:$B$10"]},
            {"formula2", [], ["A5+1"]}
          ]}
       ]}

    shift = Shift.insert(:row, 3, 2, "Sheet1")

    {"dataValidations", _,
     [
       {"dataValidation", attrs,
        [{"formula1", _, [f1]}, {"formula2", _, [f2]}]}
     ]} = DataValidations.shift_node(node, shift, "Sheet1")

    assert {_, "A7:A22"} = List.keyfind(attrs, "sqref", 0)
    assert f1 == "$B$7:$B$12"
    assert f2 == "A7+1"
  end
end
