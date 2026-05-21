defmodule ExVEx.OOXML.Comments do
  @moduledoc """
  Parse, shift, and serialize an `xl/comments*.xml` part.

  Comments are addressed by `ref="A1"` on each `<comment>` element. A
  structural row/column shift on the owning sheet rewrites those refs;
  comments inside a deletion span are dropped.
  """

  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.Utils.Coordinate

  @spec shift_xml(binary(), MutShift.t()) :: {:ok, binary()} | {:error, term()}
  def shift_xml(xml, %MutShift{} = shift) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"comments", attrs, children}} ->
        new_children = Enum.map(children, &shift_section(&1, shift))
        root = build_element({"comments", attrs, new_children})
        {:ok, Saxy.encode!(root, version: "1.0", encoding: "UTF-8", standalone: true)}

      {:ok, _other} ->
        {:error, :not_a_comments_file}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp shift_section({"commentList", attrs, comments}, shift) do
    new_comments = Enum.flat_map(comments, &shift_comment(&1, shift))
    {"commentList", attrs, new_comments}
  end

  defp shift_section(other, _shift), do: other

  defp shift_comment({"comment", attrs, children}, shift) do
    case List.keyfind(attrs, "ref", 0) do
      {_, ref} -> shift_comment_ref({"comment", attrs, children}, ref, shift)
      nil -> [{"comment", attrs, children}]
    end
  end

  defp shift_comment(other, _shift), do: [other]

  defp shift_comment_ref({"comment", attrs, children}, ref, shift) do
    case Coordinate.parse(ref) do
      {:ok, {row, col}} -> replace_ref({row, col}, shift, attrs, children)
      :error -> [{"comment", attrs, children}]
    end
  end

  defp replace_ref({row, col}, %MutShift{axis: :row} = shift, attrs, children) do
    case MutShift.apply_index(shift, row) do
      {:ok, new_row} -> [rebuild_comment(attrs, children, {new_row, col})]
      :unchanged -> [{"comment", attrs, children}]
      :deleted -> []
    end
  end

  defp replace_ref({row, col}, %MutShift{axis: :col} = shift, attrs, children) do
    case MutShift.apply_index(shift, col) do
      {:ok, new_col} -> [rebuild_comment(attrs, children, {row, new_col})]
      :unchanged -> [{"comment", attrs, children}]
      :deleted -> []
    end
  end

  defp rebuild_comment(attrs, children, new_coord) do
    new_ref = Coordinate.to_string(new_coord)
    new_attrs = List.keyreplace(attrs, "ref", 0, {"ref", new_ref})
    {"comment", new_attrs, children}
  end

  defp build_element({tag, attrs, children}) do
    Saxy.XML.element(tag, attrs, Enum.map(children, &build_child/1))
  end

  defp build_child({_, _, _} = node), do: build_element(node)
  defp build_child(text) when is_binary(text), do: Saxy.XML.characters(text)
  defp build_child(other), do: other
end
