defmodule ExVEx.Formula.Token do
  @moduledoc """
  A single token produced by `ExVEx.Formula.Tokenizer`. Every formula is
  a flat list of these tokens, preserving enough structure for
  shift-on-insert/delete to walk the stream and rewrite cell references.

  Token variants (the `:kind` field):

    * `:literal` — opaque text that should be emitted verbatim
      (operators, function names, punctuation, string literals).
    * `:cell_ref` — a concrete cell reference, optionally scoped to a
      sheet. Carries a `%Reference{}` plus the original sheet prefix.
    * `:range_ref` — a start / end pair for `A1:B5` references.
    * `:row_range` — full-row reference like `1:5`.
    * `:col_range` — full-column reference like `A:C`.
    * `:sheet_range` — 3D sheet-span prefix (`Sheet1:Sheet3!`).
  """

  alias ExVEx.Formula.Reference

  @type kind :: :literal | :cell_ref | :range_ref | :row_range | :col_range | :sheet_range

  @type t ::
          %__MODULE__{kind: :literal, text: String.t()}
          | %__MODULE__{
              kind: :cell_ref,
              sheet: String.t() | nil,
              ref: Reference.t(),
              text: String.t()
            }
          | %__MODULE__{
              kind: :range_ref,
              sheet: String.t() | nil,
              start_ref: Reference.t(),
              end_ref: Reference.t(),
              text: String.t()
            }
          | %__MODULE__{
              kind: :row_range,
              sheet: String.t() | nil,
              start_row: pos_integer(),
              start_abs?: boolean(),
              end_row: pos_integer(),
              end_abs?: boolean(),
              text: String.t()
            }
          | %__MODULE__{
              kind: :col_range,
              sheet: String.t() | nil,
              start_col: pos_integer(),
              start_abs?: boolean(),
              end_col: pos_integer(),
              end_abs?: boolean(),
              text: String.t()
            }

  defstruct [
    :kind,
    :text,
    :sheet,
    :ref,
    :start_ref,
    :end_ref,
    :start_row,
    :end_row,
    :start_col,
    :end_col,
    :start_abs?,
    :end_abs?
  ]
end
