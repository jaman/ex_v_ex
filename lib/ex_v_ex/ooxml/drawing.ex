defmodule ExVEx.OOXML.Drawing do
  @moduledoc """
  Parse, shift, and serialize an `xl/drawings/drawing*.xml` part.

  Drawings are anchored to cells via `<xdr:from>` and `<xdr:to>` (or the
  namespace-local forms) elements that contain `<col>`, `<row>`, `<colOff>`,
  `<rowOff>` children. A structural row/column shift on the owning sheet
  rewrites the row/col values inside each anchor.

  The OOXML schema uses zero-based indices for drawing anchors while the
  shift primitive expects 1-based indices, so this module translates on
  the way in and out.
  """

  alias ExVEx.Mutation.Shift, as: MutShift

  @spec shift_xml(binary(), MutShift.t()) :: {:ok, binary()} | {:error, term()}
  def shift_xml(xml, %MutShift{} = shift) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, tree} ->
        new_tree = walk(tree, shift)
        root = build_element(new_tree)
        {:ok, Saxy.encode!(root, version: "1.0", encoding: "UTF-8", standalone: true)}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp walk({tag, attrs, children}, shift) do
    if local_name(tag) in ["from", "to"] do
      {tag, attrs, Enum.map(children, &shift_anchor_child(&1, shift))}
    else
      {tag, attrs, Enum.map(children, &walk_child(&1, shift))}
    end
  end

  defp walk_child({_, _, _} = node, shift), do: walk(node, shift)
  defp walk_child(other, _shift), do: other

  defp shift_anchor_child({tag, attrs, children}, shift) do
    case {local_name(tag), shift.axis} do
      {"row", :row} -> {tag, attrs, shift_index_text(children, shift)}
      {"col", :col} -> {tag, attrs, shift_index_text(children, shift)}
      _ -> {tag, attrs, children}
    end
  end

  defp shift_anchor_child(other, _shift), do: other

  defp shift_index_text(children, shift) do
    Enum.map(children, fn
      text when is_binary(text) -> rewrite_number(text, shift)
      other -> other
    end)
  end

  defp rewrite_number(text, shift) do
    case Integer.parse(String.trim(text)) do
      {zero_based, _} ->
        one_based = zero_based + 1

        case MutShift.apply_index(shift, one_based) do
          {:ok, new_one_based} -> Integer.to_string(new_one_based - 1)
          :unchanged -> text
          :deleted -> Integer.to_string(max(shift.at - 1, 1) - 1)
        end

      :error ->
        text
    end
  end

  defp local_name(tag) do
    case :binary.split(tag, ":") do
      [_prefix, local] -> local
      [only] -> only
    end
  end

  defp build_element({tag, attrs, children}) do
    Saxy.XML.element(tag, attrs, Enum.map(children, &build_child/1))
  end

  defp build_child({_, _, _} = node), do: build_element(node)
  defp build_child(text) when is_binary(text), do: Saxy.XML.characters(text)
  defp build_child(other), do: other
end
