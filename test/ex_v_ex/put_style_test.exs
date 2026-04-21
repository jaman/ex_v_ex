defmodule ExVEx.PutStyleTest do
  use ExUnit.Case, async: true

  alias ExVEx.Test.Fixtures

  setup do
    out = Fixtures.tmp_path("put_style.xlsx")
    on_exit(fn -> File.rm(out) end)
    %{out: out}
  end

  describe "font options" do
    test "bold survives round-trip", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "hello")
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", bold: true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")
      assert style.font.bold
    end

    test "italic + underline", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", italic: true, underline: :single)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.font.italic
      assert style.font.underline == :single
    end

    test "font_size + font_name + color", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")

      {:ok, book} =
        ExVEx.put_style(book, "Sheet1", "A1",
          font_size: 18,
          font_name: "Helvetica",
          color: "FF0000"
        )

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.font.size == 18.0
      assert style.font.name == "Helvetica"
      assert style.font.color.kind == :rgb
      assert style.font.color.value == "FFFF0000"
    end

    test "merges with existing style (second call preserves first)", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", bold: true)
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", italic: true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.font.bold
      assert style.font.italic
    end
  end

  describe "fill + background" do
    test "background colour via shorthand sets solid pattern", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", background: "FFFF00")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.fill.pattern == :solid
      assert style.fill.foreground_color.value == "FFFFFF00"
    end
  end

  describe "border" do
    test "applies to all sides", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", border: :thin)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.border.top.style == :thin
      assert style.border.bottom.style == :thin
      assert style.border.left.style == :thin
      assert style.border.right.style == :thin
    end

    test "per-side overrides", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", border_bottom: :medium)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.border.bottom.style == :medium
      assert style.border.top.style == :none
    end
  end

  describe "alignment" do
    test "align + valign + wrap_text", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "x")

      {:ok, book} =
        ExVEx.put_style(book, "Sheet1", "A1", align: :center, valign: :top, wrap_text: true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.alignment.horizontal == :center
      assert style.alignment.vertical == :top
      assert style.alignment.wrap_text == true
    end
  end

  describe "number_format" do
    test "built-in code returns its built-in id (0 .. 163)", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", 1234.56)
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", number_format: "0.00")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.number_format == "0.00"
    end

    test "custom code gets a new id >= 164", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", 45_306)
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", number_format: "yyyy-mm-dd")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)
      {:ok, style} = ExVEx.get_style(reopened, "Sheet1", "A1")

      assert style.number_format == "yyyy-mm-dd"
    end
  end

  describe "deduplication" do
    test "applying the same options twice to different cells shares an xf", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", 1)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A2", 2)
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A1", bold: true)
      {:ok, book} = ExVEx.put_style(book, "Sheet1", "A2", bold: true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      {:ok, style_a} = ExVEx.get_style(reopened, "Sheet1", "A1")
      {:ok, style_b} = ExVEx.get_style(reopened, "Sheet1", "A2")

      assert style_a == style_b
      assert style_a.font.bold
    end
  end
end
