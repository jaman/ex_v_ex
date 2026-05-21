defmodule ExVEx.OOXML.CommentsTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.Comments

  @xml ~s(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>A</author></authors>
  <commentList>
    <comment ref="B2" authorId="0"><text><t>one</t></text></comment>
    <comment ref="B5" authorId="0"><text><t>two</t></text></comment>
  </commentList>
</comments>)

  test "shifts each comment ref on row insert" do
    shift = Shift.insert(:row, 3, 1, nil)
    {:ok, out} = Comments.shift_xml(@xml, shift)
    assert out =~ ~s(ref="B2")
    assert out =~ ~s(ref="B6")
  end

  test "drops comments inside a deletion span" do
    shift = Shift.delete(:row, 4, 2, nil)
    {:ok, out} = Comments.shift_xml(@xml, shift)
    assert out =~ ~s(ref="B2")
    refute out =~ ~s(ref="B5")
    refute out =~ ~s(ref="B4")
  end
end
