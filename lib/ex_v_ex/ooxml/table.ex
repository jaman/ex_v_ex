defmodule ExVEx.OOXML.Table do
  @moduledoc """
  Parse, shift, and serialize an `xl/tables/table*.xml` part.

  A table has a `ref="A1:C10"` attribute; it may embed `<autoFilter>` and
  `<sortState>` children that carry their own `ref`. Structural row/column
  shifts on the owning sheet rewrite each of these.
  """

  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.SqrefShift

  @spec shift_xml(binary(), MutShift.t()) :: {:ok, binary()} | {:error, term()}
  def shift_xml(xml, %MutShift{} = shift) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"table", attrs, children}} ->
        new_attrs = Enum.map(attrs, &shift_ref_attr(&1, shift))
        new_children = Enum.map(children, &shift_child(&1, shift))
        root = build_element({"table", new_attrs, new_children})
        {:ok, Saxy.encode!(root, version: "1.0", encoding: "UTF-8", standalone: true)}

      {:ok, _other} ->
        {:error, :not_a_table_file}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp shift_child({tag, attrs, children}, shift) when tag in ["autoFilter", "sortState"] do
    new_attrs = Enum.map(attrs, &shift_ref_attr(&1, shift))
    {tag, new_attrs, children}
  end

  defp shift_child(other, _shift), do: other

  defp shift_ref_attr({"ref", value}, shift), do: {"ref", SqrefShift.shift(value, shift)}
  defp shift_ref_attr(other, _shift), do: other

  defp build_element({tag, attrs, children}) do
    Saxy.XML.element(tag, attrs, Enum.map(children, &build_child/1))
  end

  defp build_child({_, _, _} = node), do: build_element(node)
  defp build_child(text) when is_binary(text), do: Saxy.XML.characters(text)
  defp build_child(other), do: other
end
