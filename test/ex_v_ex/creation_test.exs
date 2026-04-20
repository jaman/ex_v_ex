defmodule ExVEx.CreationTest do
  use ExUnit.Case, async: true

  alias ExVEx.Test.Fixtures

  describe "new/0" do
    test "returns a blank single-sheet workbook" do
      assert {:ok, book} = ExVEx.new()
      assert ExVEx.sheet_names(book) == ["Sheet1"]
      assert {:ok, nil} = ExVEx.get_cell(book, "Sheet1", "A1")
    end

    test "new + save + reopen round-trips" do
      out = Fixtures.tmp_path("from_scratch.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "Hello")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", 42)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.sheet_names(reopened) == ["Sheet1"]
      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "Hello"}
      assert ExVEx.get_cell(reopened, "Sheet1", "B1") == {:ok, 42}
    end
  end

  describe "add_sheet/2" do
    test "appends a new empty sheet" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "Inventory")

      assert ExVEx.sheet_names(book) == ["Sheet1", "Inventory"]
      assert {:ok, nil} = ExVEx.get_cell(book, "Inventory", "A1")
    end

    test "multiple sheets can be added" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "A")
      {:ok, book} = ExVEx.add_sheet(book, "B")
      {:ok, book} = ExVEx.add_sheet(book, "C")

      assert ExVEx.sheet_names(book) == ["Sheet1", "A", "B", "C"]
    end

    test "rejects a duplicate name" do
      {:ok, book} = ExVEx.new()
      assert {:error, :duplicate_sheet_name} = ExVEx.add_sheet(book, "Sheet1")
    end

    test "added sheets accept cell writes that survive save+reopen" do
      out = Fixtures.tmp_path("added_sheets.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "Data")
      {:ok, book} = ExVEx.add_sheet(book, "Summary")
      {:ok, book} = ExVEx.put_cell(book, "Data", "A1", "row1")
      {:ok, book} = ExVEx.put_cell(book, "Summary", "B2", 3.14)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.sheet_names(reopened) == ["Sheet1", "Data", "Summary"]
      assert ExVEx.get_cell(reopened, "Data", "A1") == {:ok, "row1"}
      assert ExVEx.get_cell(reopened, "Summary", "B2") == {:ok, 3.14}
    end
  end

  describe "rename_sheet/3" do
    test "changes the sheet name while preserving its contents" do
      out = Fixtures.tmp_path("renamed.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "kept")
      {:ok, book} = ExVEx.rename_sheet(book, "Sheet1", "Main")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.sheet_names(reopened) == ["Main"]
      assert ExVEx.get_cell(reopened, "Main", "A1") == {:ok, "kept"}
      assert {:error, :unknown_sheet} = ExVEx.get_cell(reopened, "Sheet1", "A1")
    end

    test "rejects renaming an unknown sheet" do
      {:ok, book} = ExVEx.new()
      assert {:error, :unknown_sheet} = ExVEx.rename_sheet(book, "Ghost", "Main")
    end

    test "rejects a rename that collides with another sheet" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "Target")

      assert {:error, :duplicate_sheet_name} =
               ExVEx.rename_sheet(book, "Sheet1", "Target")
    end

    test "renaming to the same name is a no-op" do
      {:ok, book} = ExVEx.new()
      assert {:ok, ^book} = ExVEx.rename_sheet(book, "Sheet1", "Sheet1")
    end
  end

  describe "remove_sheet/2" do
    test "drops the sheet from the workbook" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "Temp")
      {:ok, book} = ExVEx.add_sheet(book, "Keep")
      {:ok, book} = ExVEx.remove_sheet(book, "Temp")

      assert ExVEx.sheet_names(book) == ["Sheet1", "Keep"]
    end

    test "removes the worksheet part and its manifest entries" do
      out = Fixtures.tmp_path("removed.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "Scratch")
      {:ok, book} = ExVEx.put_cell(book, "Scratch", "A1", "will vanish")

      # Grab the path of the scratch sheet before removal
      {:ok, scratch_path} = ExVEx.sheet_path(book, "Scratch")
      {:ok, book} = ExVEx.remove_sheet(book, "Scratch")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      refute Map.has_key?(reopened.parts, scratch_path)

      refute Enum.any?(
               reopened.content_types.overrides,
               &(&1.part_name == "/" <> scratch_path)
             )

      refute Enum.any?(
               reopened.workbook_rels.entries,
               &String.ends_with?(&1.target, Path.basename(scratch_path))
             )

      assert ExVEx.sheet_names(reopened) == ["Sheet1"]
    end

    test "rejects removing an unknown sheet" do
      {:ok, book} = ExVEx.new()
      assert {:error, :unknown_sheet} = ExVEx.remove_sheet(book, "Ghost")
    end

    test "rejects removing the last remaining sheet" do
      {:ok, book} = ExVEx.new()
      assert {:error, :last_sheet} = ExVEx.remove_sheet(book, "Sheet1")
    end
  end

  describe "full creation workflow" do
    test "build a multi-sheet template from scratch and round-trip it" do
      out = Fixtures.tmp_path("full_template.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.rename_sheet(book, "Sheet1", "Summary")
      {:ok, book} = ExVEx.add_sheet(book, "Data A")
      {:ok, book} = ExVEx.add_sheet(book, "Data B")
      {:ok, book} = ExVEx.add_sheet(book, "Formulas")
      {:ok, book} = ExVEx.put_cell(book, "Data A", "A1", "alpha")
      {:ok, book} = ExVEx.put_cell(book, "Data B", "A1", 100)

      {:ok, book} =
        ExVEx.put_cell(book, "Formulas", "A1", {:formula, "=SUM('Data B'!A1:A10)"})

      {:ok, book} = ExVEx.merge_cells(book, "Summary", "A1:C1", preserve_values: true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.sheet_names(reopened) == ["Summary", "Data A", "Data B", "Formulas"]
      assert ExVEx.get_cell(reopened, "Data A", "A1") == {:ok, "alpha"}
      assert ExVEx.get_cell(reopened, "Data B", "A1") == {:ok, 100}
      assert {:ok, "=SUM('Data B'!A1:A10)"} = ExVEx.get_formula(reopened, "Formulas", "A1")
      assert {:ok, ["A1:C1"]} = ExVEx.merged_ranges(reopened, "Summary")
    end
  end
end
