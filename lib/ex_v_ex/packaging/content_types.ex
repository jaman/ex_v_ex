defmodule ExVEx.Packaging.ContentTypes do
  @moduledoc """
  The `[Content_Types].xml` manifest that lives at the root of every OPC
  package.

  It maps file extensions (`Default`) and individual part paths (`Override`) to
  MIME-like content type strings. ExVEx preserves every entry verbatim so
  unknown content types (custom XML parts, plugin data, future schema
  extensions) survive a read/edit/save cycle.
  """

  alias ExVEx.Packaging.ContentTypes.{Default, Override}

  @namespace "http://schemas.openxmlformats.org/package/2006/content-types"

  @type t :: %__MODULE__{
          defaults: [Default.t()],
          overrides: [Override.t()]
        }

  defstruct defaults: [], overrides: []

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"Types", _attrs, children}} ->
        {:ok, build_manifest(children)}

      {:ok, _other} ->
        {:error, :not_a_content_types_manifest}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @spec serialize(t()) :: binary()
  def serialize(%__MODULE__{defaults: defaults, overrides: overrides}) do
    default_elems = Enum.map(defaults, &default_to_element/1)
    override_elems = Enum.map(overrides, &override_to_element/1)

    root =
      Saxy.XML.element("Types", [{"xmlns", @namespace}], default_elems ++ override_elems)

    Saxy.encode!(root, version: "1.0", encoding: "UTF-8", standalone: true)
  end

  defp build_manifest(children) do
    Enum.reduce(children, %__MODULE__{}, fn
      {"Default", attrs, _}, acc ->
        %{acc | defaults: acc.defaults ++ [default_from_attrs(attrs)]}

      {"Override", attrs, _}, acc ->
        %{acc | overrides: acc.overrides ++ [override_from_attrs(attrs)]}

      _ignored, acc ->
        acc
    end)
  end

  defp default_from_attrs(attrs) do
    %Default{
      extension: fetch_attr!(attrs, "Extension"),
      content_type: fetch_attr!(attrs, "ContentType")
    }
  end

  defp override_from_attrs(attrs) do
    %Override{
      part_name: fetch_attr!(attrs, "PartName"),
      content_type: fetch_attr!(attrs, "ContentType")
    }
  end

  defp default_to_element(%Default{extension: ext, content_type: ct}) do
    Saxy.XML.element("Default", [{"Extension", ext}, {"ContentType", ct}], [])
  end

  defp override_to_element(%Override{part_name: pn, content_type: ct}) do
    Saxy.XML.element("Override", [{"PartName", pn}, {"ContentType", ct}], [])
  end

  defp fetch_attr!(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {^name, value} -> value
      nil -> raise ArgumentError, "missing #{name} attribute in [Content_Types].xml"
    end
  end
end
