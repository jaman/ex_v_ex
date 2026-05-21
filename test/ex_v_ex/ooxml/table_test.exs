defmodule ExVEx.OOXML.TableTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.Table

  @xml ~s(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C10" totalsRowShown="0">
  <autoFilter ref="A1:C10"/>
  <tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns>
</table>)

  test "shifts table ref and autoFilter ref on row insert" do
    shift = Shift.insert(:row, 5, 2, nil)
    {:ok, out} = Table.shift_xml(@xml, shift)
    assert out =~ ~s(ref="A1:C12")
    assert out =~ ~s(<autoFilter ref="A1:C12")
  end
end
