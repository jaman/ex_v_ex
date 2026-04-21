defmodule ExVEx.DefinedNamesTest do
  use ExUnit.Case, async: true

  alias ExVEx.Test.Fixtures

  describe "define_name/4 global" do
    test "adds a workbook-wide name that survives save + reopen" do
      out = Fixtures.tmp_path("def_name_global.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "TaxRate", "Sheet1!$B$2")
      :ok = ExVEx.save(book, out)

      {:ok, reopened} = ExVEx.open(out)

      assert [
               %{name: "TaxRate", reference: "Sheet1!$B$2", scope: :global, hidden: false}
             ] = ExVEx.defined_names(reopened)
    end

    test "accepts dynamic formulas (OFFSET / INDIRECT)" do
      out = Fixtures.tmp_path("dynamic_name.xlsx")
      on_exit(fn -> File.rm(out) end)

      ref = "OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)"

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "DataColumn", ref)
      :ok = ExVEx.save(book, out)

      {:ok, reopened} = ExVEx.open(out)
      assert [%{name: "DataColumn", reference: ^ref}] = ExVEx.defined_names(reopened)
    end

    test "replacing a name with the same scope overwrites rather than duplicating" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "X", "Sheet1!$A$1")
      {:ok, book} = ExVEx.define_name(book, "X", "Sheet1!$B$2")

      assert [%{name: "X", reference: "Sheet1!$B$2"}] = ExVEx.defined_names(book)
    end
  end

  describe "define_name/4 sheet-local scope" do
    test "binds the name to a specific sheet" do
      out = Fixtures.tmp_path("local_name.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.add_sheet(book, "Data")

      {:ok, book} =
        ExVEx.define_name(book, "LocalSum", "Sheet1!$A$1:$A$10", scope: {:sheet, "Sheet1"})

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert [%{name: "LocalSum", scope: {:sheet, "Sheet1"}}] =
               ExVEx.defined_names(reopened)
    end

    test "rejects scope pointing at a non-existent sheet" do
      {:ok, book} = ExVEx.new()

      assert {:error, :unknown_sheet} =
               ExVEx.define_name(book, "X", "$A$1", scope: {:sheet, "Ghost"})
    end

    test "same name can coexist in global + sheet-local scope" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "Total", "Sheet1!$A$1")

      {:ok, book} =
        ExVEx.define_name(book, "Total", "Sheet1!$B$1", scope: {:sheet, "Sheet1"})

      names = ExVEx.defined_names(book)
      assert length(names) == 2
      assert Enum.any?(names, &(&1.scope == :global))
      assert Enum.any?(names, &(&1.scope == {:sheet, "Sheet1"}))
    end
  end

  describe "define_name/4 hidden" do
    test "stores a hidden name that round-trips" do
      out = Fixtures.tmp_path("hidden_name.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "_xlnm.Print_Area", "Sheet1!$A$1:$C$5", hidden: true)
      :ok = ExVEx.save(book, out)

      {:ok, reopened} = ExVEx.open(out)
      assert [%{name: "_xlnm.Print_Area", hidden: true}] = ExVEx.defined_names(reopened)
    end
  end

  describe "remove_defined_name/3" do
    test "drops the named range" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "A", "Sheet1!$A$1")
      {:ok, book} = ExVEx.define_name(book, "B", "Sheet1!$B$1")
      {:ok, book} = ExVEx.remove_defined_name(book, "A")

      names = Enum.map(ExVEx.defined_names(book), & &1.name)
      assert names == ["B"]
    end

    test "returns :error for an unknown name" do
      {:ok, book} = ExVEx.new()
      assert {:error, :unknown_name} = ExVEx.remove_defined_name(book, "Ghost")
    end

    test "scope-aware removal — global and sheet-local are independent" do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.define_name(book, "X", "Sheet1!$A$1")
      {:ok, book} = ExVEx.define_name(book, "X", "Sheet1!$B$1", scope: {:sheet, "Sheet1"})

      {:ok, book} = ExVEx.remove_defined_name(book, "X")

      # global removed, local remains
      assert [%{scope: {:sheet, "Sheet1"}}] = ExVEx.defined_names(book)
    end
  end

  describe "defined_names/1 on a workbook opened from an xlsm fixture" do
    test "parses names that were already present" do
      {:ok, book} = ExVEx.open(Fixtures.path("with_macros.xlsm"))
      names = ExVEx.defined_names(book)
      assert is_list(names)
    end
  end
end
