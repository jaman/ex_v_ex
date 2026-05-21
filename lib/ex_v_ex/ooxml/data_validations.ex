defmodule ExVEx.OOXML.DataValidations do
  @moduledoc """
  Cascades a `%ExVEx.Mutation.Shift{}` through a `<dataValidations>` node:
  the `sqref` attribute and any `<formula1>` / `<formula2>` text children
  of each child `<dataValidation>` element.
  """

  alias ExVEx.Formula.Serializer
  alias ExVEx.Formula.Shift, as: FormulaShift
  alias ExVEx.Formula.Tokenizer
  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.SqrefShift

  @type node_tuple :: {String.t(), list(), list()}

  @spec shift_node(node_tuple(), MutShift.t(), String.t()) :: node_tuple()
  def shift_node({"dataValidations", attrs, children}, %MutShift{} = shift, sheet_name) do
    new_children = Enum.map(children, &shift_child(&1, shift, sheet_name))
    {"dataValidations", attrs, new_children}
  end

  defp shift_child({"dataValidation", attrs, children}, shift, sheet_name) do
    new_attrs = Enum.map(attrs, &shift_sqref(&1, shift))
    new_children = Enum.map(children, &shift_formula_child(&1, shift, sheet_name))
    {"dataValidation", new_attrs, new_children}
  end

  defp shift_child(other, _shift, _sheet_name), do: other

  defp shift_sqref({"sqref", value}, shift), do: {"sqref", SqrefShift.shift(value, shift)}
  defp shift_sqref(other, _shift), do: other

  defp shift_formula_child({tag, attrs, children}, shift, sheet_name)
       when tag in ["formula1", "formula2"] do
    new_children = Enum.map(children, &shift_text(&1, shift, sheet_name))
    {tag, attrs, new_children}
  end

  defp shift_formula_child(other, _shift, _sheet_name), do: other

  defp shift_text(text, shift, sheet_name) when is_binary(text) do
    ("=" <> text)
    |> Tokenizer.tokenize()
    |> FormulaShift.apply(shift, sheet_name)
    |> Serializer.to_string()
    |> String.trim_leading("=")
  end

  defp shift_text(other, _shift, _sheet_name), do: other
end
