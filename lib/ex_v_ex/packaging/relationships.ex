defmodule ExVEx.Packaging.Relationships do
  @moduledoc """
  An OPC `.rels` file — a graph of `<Relationship>` records linking parts
  of the package together.

  Every `.rels` file lives under a `_rels/` directory at its source scope:

  * `_rels/.rels` — package-level relationships (root).
  * `xl/_rels/workbook.xml.rels` — relationships for `xl/workbook.xml`.
  * `xl/worksheets/_rels/sheet1.xml.rels` — relationships for `sheet1.xml`.

  Relationship targets are resolved relative to the *source* part's
  directory, never the `.rels` file's directory. `resolve/2` handles the
  path arithmetic.
  """

  alias ExVEx.Packaging.Relationships.Relationship

  @namespace "http://schemas.openxmlformats.org/package/2006/relationships"

  @type t :: %__MODULE__{entries: [Relationship.t()]}
  defstruct entries: []

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"Relationships", _attrs, children}} ->
        {:ok, %__MODULE__{entries: Enum.flat_map(children, &relationship_from_child/1)}}

      {:ok, _other} ->
        {:error, :not_a_relationships_file}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @spec serialize(t()) :: binary()
  def serialize(%__MODULE__{entries: entries}) do
    children = Enum.map(entries, &relationship_element/1)
    root = Saxy.XML.element("Relationships", [{"xmlns", @namespace}], children)
    Saxy.encode!(root, version: "1.0", encoding: "UTF-8", standalone: true)
  end

  defp relationship_element(%Relationship{} = rel) do
    attrs =
      [{"Id", rel.id}, {"Type", rel.type}, {"Target", rel.target}] ++
        target_mode_attr(rel.target_mode)

    Saxy.XML.element("Relationship", attrs, [])
  end

  defp target_mode_attr(:external), do: [{"TargetMode", "External"}]
  defp target_mode_attr(_), do: []

  @spec get(t(), String.t()) :: {:ok, Relationship.t()} | :error
  def get(%__MODULE__{entries: entries}, id) do
    case Enum.find(entries, &(&1.id == id)) do
      %Relationship{} = rel -> {:ok, rel}
      nil -> :error
    end
  end

  @doc """
  Returns the absolute package path that `relationship.target` points to,
  given the path of the source `.rels` file that declared the relationship.
  """
  @spec resolve(Relationship.t(), String.t()) :: String.t()
  def resolve(%Relationship{target: target}, _rels_path) when binary_part(target, 0, 1) == "/" do
    String.trim_leading(target, "/")
  end

  def resolve(%Relationship{target: target}, rels_path) do
    source_dir = source_directory(rels_path)
    Path.join(source_dir, target) |> Path.expand("/") |> String.trim_leading("/")
  end

  defp source_directory(rels_path) do
    rels_path
    |> Path.dirname()
    |> String.replace_suffix("/_rels", "")
    |> String.replace_suffix("_rels", "")
  end

  defp relationship_from_child({"Relationship", attrs, _}) do
    [
      %Relationship{
        id: fetch_attr!(attrs, "Id"),
        type: fetch_attr!(attrs, "Type"),
        target: fetch_attr!(attrs, "Target"),
        target_mode: target_mode(attrs)
      }
    ]
  end

  defp relationship_from_child(_), do: []

  defp target_mode(attrs) do
    case List.keyfind(attrs, "TargetMode", 0) do
      {_, "External"} -> :external
      _ -> :internal
    end
  end

  defp fetch_attr!(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {^name, value} -> value
      nil -> raise ArgumentError, "missing #{name} attribute on <Relationship>"
    end
  end
end
