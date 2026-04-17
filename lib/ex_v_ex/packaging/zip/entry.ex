defmodule ExVEx.Packaging.Zip.Entry do
  @moduledoc """
  A single member of an xlsx ZIP archive — a path inside the package and the
  raw bytes at that path. Directory entries (zero-length, name ending in `/`)
  are dropped by `:zip.unzip/2` and are not represented here.
  """

  @enforce_keys [:path, :data]
  defstruct [:path, :data]

  @type t :: %__MODULE__{path: String.t(), data: binary()}
end
