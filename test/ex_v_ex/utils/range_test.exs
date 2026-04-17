defmodule ExVEx.Utils.RangeTest do
  use ExUnit.Case, async: true

  alias ExVEx.Utils.Range, as: R

  describe "parse/1" do
    test "parses a proper A1:B2 range" do
      assert {:ok, %R{top_left: {1, 1}, bottom_right: {2, 2}}} = R.parse("A1:B2")
    end

    test "parses a multi-column, multi-row range" do
      assert {:ok, %R{top_left: {3, 2}, bottom_right: {10, 5}}} = R.parse("B3:E10")
    end

    test "parses a 1x1 range as both endpoints equal" do
      assert {:ok, %R{top_left: {1, 1}, bottom_right: {1, 1}}} = R.parse("A1:A1")
    end

    test "rejects malformed input" do
      assert :error = R.parse("")
      assert :error = R.parse("A1")
      assert :error = R.parse("A1:oops")
      assert :error = R.parse("1:2")
    end

    test "normalises swapped corners" do
      assert {:ok, %R{top_left: {1, 1}, bottom_right: {2, 2}}} = R.parse("B2:A1")
    end
  end

  describe "to_string/1" do
    test "emits the canonical A1:B2 form" do
      {:ok, range} = R.parse("A1:B2")
      assert R.to_string(range) == "A1:B2"
    end
  end

  describe "overlaps?/2" do
    test "two identical ranges overlap" do
      {:ok, a} = R.parse("A1:B2")
      {:ok, b} = R.parse("A1:B2")
      assert R.overlaps?(a, b)
    end

    test "partial overlap is detected" do
      {:ok, a} = R.parse("A1:C3")
      {:ok, b} = R.parse("B2:D4")
      assert R.overlaps?(a, b)
    end

    test "edge-adjacent ranges don't overlap" do
      {:ok, a} = R.parse("A1:B2")
      {:ok, b} = R.parse("C1:D2")
      refute R.overlaps?(a, b)
    end

    test "ranges on different rows don't overlap" do
      {:ok, a} = R.parse("A1:Z1")
      {:ok, b} = R.parse("A5:Z5")
      refute R.overlaps?(a, b)
    end
  end

  describe "cells/1" do
    test "enumerates every coordinate in row-major order" do
      {:ok, range} = R.parse("A1:B2")
      assert R.cells(range) == [{1, 1}, {1, 2}, {2, 1}, {2, 2}]
    end
  end

  describe "anchor/1" do
    test "returns the top-left coordinate" do
      {:ok, range} = R.parse("B3:E10")
      assert R.anchor(range) == {3, 2}
    end
  end
end
