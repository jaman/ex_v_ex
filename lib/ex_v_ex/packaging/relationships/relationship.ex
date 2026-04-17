defmodule ExVEx.Packaging.Relationships.Relationship do
  @moduledoc """
  A single `<Relationship>` record: an opaque `id`, a type URI, and a
  `target` path that is resolved relative to the directory containing
  the `.rels` file that declared it.
  """

  @enforce_keys [:id, :type, :target]
  defstruct [:id, :type, :target, target_mode: :internal]

  @type target_mode :: :internal | :external
  @type t :: %__MODULE__{
          id: String.t(),
          type: String.t(),
          target: String.t(),
          target_mode: target_mode()
        }
end
