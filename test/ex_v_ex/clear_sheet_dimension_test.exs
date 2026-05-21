defmodule ExVEx.ClearSheetDimensionTest do
  use ExUnit.Case, async: true

  alias ExVEx.Test.Fixtures
  alias ExVEx.Workbook

  describe "clear_sheet_dimension/2" do
    test "removes the <dimension> element while preserving other pre-sheetData siblings" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      {:ok, path} = ExVEx.sheet_path(book, "Sheet1")

      {:ok, before, _} = Workbook.fetch_sheet_tree(book, path)
      assert Enum.any?(before.pre_sheet_data, &match?({"dimension", _, _}, &1))

      assert {:ok, book} = ExVEx.clear_sheet_dimension(book, "Sheet1")

      {:ok, ed, _} = Workbook.fetch_sheet_tree(book, path)
      refute Enum.any?(ed.pre_sheet_data, &match?({"dimension", _, _}, &1))
      assert Enum.any?(ed.pre_sheet_data, &match?({"sheetViews", _, _}, &1))
      assert Enum.any?(ed.pre_sheet_data, &match?({"sheetFormatPr", _, _}, &1))
    end

    test "cells remain intact after clearing" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      {:ok, book} = ExVEx.clear_sheet_dimension(book, "Sheet1")

      assert ExVEx.get_cell(book, "Sheet1", "A1") == {:ok, "A1"}
    end

    test "round-trips through save/open" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      {:ok, book} = ExVEx.clear_sheet_dimension(book, "Sheet1")

      out = Path.join(System.tmp_dir!(), "exvex_clear_dim_#{System.unique_integer([:positive])}.xlsx")
      on_exit(fn -> File.rm_rf!(out) end)
      :ok = ExVEx.save(book, out)

      {:ok, reopened} = ExVEx.open(out)
      {:ok, path} = ExVEx.sheet_path(reopened, "Sheet1")
      {:ok, ed, _} = Workbook.fetch_sheet_tree(reopened, path)
      refute Enum.any?(ed.pre_sheet_data, &match?({"dimension", _, _}, &1))
    end

    test "is a no-op when the sheet has no <dimension>" do
      {:ok, book} = ExVEx.new()
      assert {:ok, book} = ExVEx.clear_sheet_dimension(book, "Sheet1")

      {:ok, path} = ExVEx.sheet_path(book, "Sheet1")
      {:ok, ed, _} = Workbook.fetch_sheet_tree(book, path)
      refute Enum.any?(ed.pre_sheet_data, &match?({"dimension", _, _}, &1))
    end

    test "returns {:error, :unknown_sheet} for a sheet that doesn't exist" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      assert ExVEx.clear_sheet_dimension(book, "Nope") == {:error, :unknown_sheet}
    end
  end
end
