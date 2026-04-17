defmodule ExVEx.Packaging.ZipTest do
  use ExUnit.Case, async: true

  alias ExVEx.Packaging.Zip
  alias ExVEx.Test.Fixtures

  describe "read/1" do
    test "returns every entry with its path and bytes for a real xlsx" do
      {:ok, entries} = Zip.read(Fixtures.path("empty.xlsx"))

      paths = Enum.map(entries, & &1.path)
      assert "[Content_Types].xml" in paths
      assert "xl/workbook.xml" in paths
      assert "xl/worksheets/sheet1.xml" in paths

      content_types =
        Enum.find(entries, &(&1.path == "[Content_Types].xml"))

      assert content_types.data =~ "<Types"
      assert byte_size(content_types.data) > 0
    end

    test "returns an error tuple when the file does not exist" do
      assert {:error, _} = Zip.read(Fixtures.tmp_path("does_not_exist.xlsx"))
    end
  end

  describe "read/1 + write/2 round-trip" do
    test "writing the entries back and reading them again yields the same bytes" do
      {:ok, entries} = Zip.read(Fixtures.path("empty.xlsx"))
      out = Fixtures.tmp_path("empty_roundtrip.xlsx")

      :ok = Zip.write(out, entries)
      {:ok, round_tripped} = Zip.read(out)

      assert entries_by_path(entries) == entries_by_path(round_tripped)
    after
      File.rm(Fixtures.tmp_path("empty_roundtrip.xlsx"))
    end

    test "preserves xlsm macro binary byte-for-byte" do
      {:ok, entries} = Zip.read(Fixtures.path("with_macros.xlsm"))

      vba =
        Enum.find(entries, fn entry ->
          String.ends_with?(entry.path, "vbaProject.bin")
        end)

      assert vba, "fixture should contain a vbaProject.bin macro blob"
      assert byte_size(vba.data) > 0

      out = Fixtures.tmp_path("macros_roundtrip.xlsm")
      :ok = Zip.write(out, entries)
      {:ok, round_tripped} = Zip.read(out)

      round_tripped_vba =
        Enum.find(round_tripped, &(&1.path == vba.path))

      assert round_tripped_vba.data == vba.data
    after
      File.rm(Fixtures.tmp_path("macros_roundtrip.xlsm"))
    end
  end

  defp entries_by_path(entries) do
    Map.new(entries, fn entry -> {entry.path, entry.data} end)
  end
end
