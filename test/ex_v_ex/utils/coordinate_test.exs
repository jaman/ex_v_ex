defmodule ExVEx.Utils.CoordinateTest do
  use ExUnit.Case, async: true

  alias ExVEx.Utils.Coordinate

  describe "parse/1" do
    test "single letter columns" do
      assert {:ok, {1, 1}} = Coordinate.parse("A1")
      assert {:ok, {7, 2}} = Coordinate.parse("B7")
      assert {:ok, {42, 26}} = Coordinate.parse("Z42")
    end

    test "double letter columns" do
      assert {:ok, {1, 27}} = Coordinate.parse("AA1")
      assert {:ok, {5, 28}} = Coordinate.parse("AB5")
      assert {:ok, {1, 702}} = Coordinate.parse("ZZ1")
    end

    test "triple letter columns (Excel caps at XFD/1048576)" do
      assert {:ok, {1, 703}} = Coordinate.parse("AAA1")
      assert {:ok, {1_048_576, 16_384}} = Coordinate.parse("XFD1048576")
    end

    test "case insensitive" do
      assert {:ok, {1, 1}} = Coordinate.parse("a1")
      assert {:ok, {1, 27}} = Coordinate.parse("aa1")
    end

    test "rejects malformed input" do
      assert :error = Coordinate.parse("")
      assert :error = Coordinate.parse("1A")
      assert :error = Coordinate.parse("A")
      assert :error = Coordinate.parse("1")
      assert :error = Coordinate.parse("A0")
      assert :error = Coordinate.parse("A 1")
      assert :error = Coordinate.parse("@1")
    end
  end

  describe "to_string/1" do
    test "inverts parse/1" do
      assert Coordinate.to_string({1, 1}) == "A1"
      assert Coordinate.to_string({42, 26}) == "Z42"
      assert Coordinate.to_string({1, 27}) == "AA1"
      assert Coordinate.to_string({1, 702}) == "ZZ1"
      assert Coordinate.to_string({1, 703}) == "AAA1"
      assert Coordinate.to_string({1_048_576, 16_384}) == "XFD1048576"
    end
  end

  describe "column_label/1" do
    test "1 -> A, 26 -> Z, 27 -> AA" do
      assert Coordinate.column_label(1) == "A"
      assert Coordinate.column_label(26) == "Z"
      assert Coordinate.column_label(27) == "AA"
      assert Coordinate.column_label(52) == "AZ"
      assert Coordinate.column_label(702) == "ZZ"
      assert Coordinate.column_label(703) == "AAA"
    end
  end

  describe "column_number/1" do
    test "inverts column_label/1" do
      for n <- [1, 26, 27, 52, 702, 703, 16_384] do
        assert Coordinate.column_number(Coordinate.column_label(n)) == n
      end
    end
  end
end
