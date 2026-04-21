defmodule ExVEx.InsertRowTest do
  use ExUnit.Case, async: true

  alias ExVEx.Test.Fixtures

  setup do
    out = Fixtures.tmp_path("insert_row.xlsx")
    on_exit(fn -> File.rm(out) end)
    %{out: out}
  end

  describe "insert_row/4 — basic cell shift" do
    test "cells at or below the insert point move down", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "header")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A3", "row3")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A5", "row5")

      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 3)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "header"}
      assert ExVEx.get_cell(reopened, "Sheet1", "A3") == {:ok, nil}
      assert ExVEx.get_cell(reopened, "Sheet1", "A4") == {:ok, "row3"}
      assert ExVEx.get_cell(reopened, "Sheet1", "A6") == {:ok, "row5"}
    end

    test "inserting multiple rows at once", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A5", "five")
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 3, 3)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A8") == {:ok, "five"}
    end
  end

  describe "insert_row/4 — formula rewriting" do
    test "formulas referencing shifted cells update", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A5", 10)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", {:formula, "=A5+1"})

      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 3)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_formula(reopened, "Sheet1", "B1") == {:ok, "=A6+1"}
    end

    test "range formulas expand", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", {:formula, "=SUM(A1:A10)"})
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 5, 2)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_formula(reopened, "Sheet1", "B1") == {:ok, "=SUM(A1:A12)"}
    end

    test "formulas referencing cells above the insert point are unchanged", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", {:formula, "=A1+A2"})
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 5)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_formula(reopened, "Sheet1", "B1") == {:ok, "=A1+A2"}
    end
  end

  describe "insert_row/4 — merged cells" do
    test "merged range below the insert shifts down", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A5:B7")
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 3, 2)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, ["A7:B9"]} = ExVEx.merged_ranges(reopened, "Sheet1")
    end

    test "merged range above the insert is unchanged", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 5)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, ["A1:B2"]} = ExVEx.merged_ranges(reopened, "Sheet1")
    end
  end

  describe "insert_row/4 — defined names" do
    test "defined names referencing shifted cells update", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "Total", "Sheet1!$A$5")
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 3)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert [%{name: "Total", reference: "Sheet1!$A$6"}] = ExVEx.defined_names(reopened)
    end
  end

  describe "delete_row/4" do
    test "removes rows in the deletion span and shifts below up", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "keep")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A3", "drop")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A5", "survive")
      {:ok, book} = ExVEx.delete_row(book, "Sheet1", 3)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "keep"}
      assert ExVEx.get_cell(reopened, "Sheet1", "A3") == {:ok, nil}
      assert ExVEx.get_cell(reopened, "Sheet1", "A4") == {:ok, "survive"}
    end

    test "formulas referencing deleted rows become #REF!", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A5", 10)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", {:formula, "=A5+1"})
      {:ok, book} = ExVEx.delete_row(book, "Sheet1", 5)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, formula} = ExVEx.get_formula(reopened, "Sheet1", "B1")
      assert formula =~ "#REF!"
    end
  end

  describe "insert_column/4" do
    test "cells right of the insert column shift right", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "a")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "C1", "c")
      {:ok, book} = ExVEx.insert_column(book, "Sheet1", 2)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "a"}
      assert ExVEx.get_cell(reopened, "Sheet1", "D1") == {:ok, "c"}
    end
  end

  describe "errors" do
    test "rejects an unknown sheet" do
      {:ok, book} = ExVEx.new()
      assert {:error, :unknown_sheet} = ExVEx.insert_row(book, "Ghost", 1)
    end
  end
end
