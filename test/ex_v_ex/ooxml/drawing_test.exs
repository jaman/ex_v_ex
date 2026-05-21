defmodule ExVEx.OOXML.DrawingTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.Drawing

  @xml ~s(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  </xdr:twoCellAnchor>
</xdr:wsDr>)

  test "shifts row anchors (zero-based) on row insert" do
    shift = Shift.insert(:row, 3, 2, nil)
    {:ok, out} = Drawing.shift_xml(@xml, shift)
    assert out =~ ~s(<xdr:row>6</xdr:row>)
    assert out =~ ~s(<xdr:row>11</xdr:row>)
  end

  test "leaves col values alone on row shift" do
    shift = Shift.insert(:row, 3, 2, nil)
    {:ok, out} = Drawing.shift_xml(@xml, shift)
    assert out =~ ~s(<xdr:col>1</xdr:col>)
    assert out =~ ~s(<xdr:col>3</xdr:col>)
  end
end
