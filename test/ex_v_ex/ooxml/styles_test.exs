defmodule ExVEx.OOXML.StylesTest do
  use ExUnit.Case, async: true

  alias ExVEx.OOXML.Styles
  alias ExVEx.OOXML.Styles.{CellFormat, NumFmt}

  @stylesheet """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="2">
      <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
      <numFmt numFmtId="165" formatCode="0.000%"/>
    </numFmts>
    <cellXfs count="4">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
      <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      <xf numFmtId="165" fontId="1" fillId="0" borderId="0" applyNumberFormat="1"/>
    </cellXfs>
  </styleSheet>
  """

  describe "parse/1" do
    test "extracts custom number formats" do
      assert {:ok, %Styles{num_fmts: num_fmts}} = Styles.parse(@stylesheet)

      assert %NumFmt{id: 164, format_code: "yyyy-mm-dd"} in num_fmts
      assert %NumFmt{id: 165, format_code: "0.000%"} in num_fmts
    end

    test "extracts cell formats (xf) with numFmtId and font references" do
      assert {:ok, %Styles{cell_formats: cell_formats}} = Styles.parse(@stylesheet)

      assert length(cell_formats) == 4
      assert Enum.at(cell_formats, 0).num_fmt_id == 0
      assert Enum.at(cell_formats, 1).num_fmt_id == 14
      assert Enum.at(cell_formats, 2).num_fmt_id == 164
      assert Enum.at(cell_formats, 3) == %CellFormat{num_fmt_id: 165, font_id: 1}
    end
  end

  describe "date_format?/2" do
    test "recognizes built-in date format IDs" do
      {:ok, styles} = Styles.parse(@stylesheet)

      for id <- [14, 15, 16, 17, 22, 45, 46, 47] do
        assert Styles.date_format?(styles, %CellFormat{num_fmt_id: id}),
               "numFmtId #{id} should be recognized as a date format"
      end
    end

    test "recognizes custom formats that contain date tokens" do
      {:ok, styles} = Styles.parse(@stylesheet)

      assert Styles.date_format?(styles, %CellFormat{num_fmt_id: 164})
    end

    test "rejects non-date formats" do
      {:ok, styles} = Styles.parse(@stylesheet)

      refute Styles.date_format?(styles, %CellFormat{num_fmt_id: 0})
      refute Styles.date_format?(styles, %CellFormat{num_fmt_id: 165})
    end
  end

  describe "cell_format/2" do
    test "returns the xf at the given index" do
      {:ok, styles} = Styles.parse(@stylesheet)

      assert {:ok, %CellFormat{num_fmt_id: 14}} = Styles.cell_format(styles, 1)
      assert {:ok, %CellFormat{num_fmt_id: 164}} = Styles.cell_format(styles, 2)
    end

    test "returns :error for out-of-range indexes" do
      {:ok, styles} = Styles.parse(@stylesheet)
      assert :error = Styles.cell_format(styles, 99)
    end
  end
end
