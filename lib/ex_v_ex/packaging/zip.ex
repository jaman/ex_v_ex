defmodule ExVEx.Packaging.Zip do
  @moduledoc """
  Archive IO for the Open Packaging Convention container used by `.xlsx` /
  `.xlsm` / `.xltx` files.

  An xlsx file is a ZIP archive containing XML parts. This module is the
  lowest layer: it reads the archive into a flat list of `%Entry{}` records
  and writes a list of entries back to disk. It does not know about the OOXML
  schema — higher-level modules interpret the parts.
  """

  alias ExVEx.Packaging.Zip.Entry

  @type path :: Path.t()

  @spec read(path()) :: {:ok, [Entry.t()]} | {:error, term()}
  def read(path) do
    case :zip.unzip(to_charlist(path), [:memory]) do
      {:ok, raw_entries} ->
        entries =
          raw_entries
          |> Enum.map(&to_entry/1)
          |> Enum.reject(&directory_stub?/1)

        {:ok, entries}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @spec write(path(), [Entry.t()]) :: :ok | {:error, term()}
  def write(path, entries) do
    file_list = Enum.map(entries, &to_zip_tuple/1)

    case :zip.create(to_charlist(path), file_list) do
      {:ok, _} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  defp to_entry({name, data}) when is_binary(data) do
    %Entry{path: IO.iodata_to_binary([name]), data: data}
  end

  defp to_zip_tuple(%Entry{path: path, data: data}) do
    {to_charlist(path), data}
  end

  defp directory_stub?(%Entry{path: path, data: data}) do
    data == "" and String.ends_with?(path, "/")
  end
end
