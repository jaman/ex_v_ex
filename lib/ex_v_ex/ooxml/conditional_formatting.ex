defmodule ExVEx.OOXML.ConditionalFormatting do
  @moduledoc """
  Cascades a `%ExVEx.Mutation.Shift{}` through a `<conditionalFormatting>`
  SimpleForm node: its `sqref` attribute (space-separated ranges) and any
  `<formula>` text inside child `<cfRule>` elements.
  """

  alias ExVEx.Formula.Serializer
  alias ExVEx.Formula.Shift, as: FormulaShift
  alias ExVEx.Formula.Tokenizer
  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.SqrefShift

  @type node_tuple :: {String.t(), list(), list()}

  @spec shift_node(node_tuple(), MutShift.t(), String.t()) :: node_tuple()
  def shift_node({"conditionalFormatting", attrs, children}, %MutShift{} = shift, sheet_name) do
    new_attrs = Enum.map(attrs, &shift_sqref_attr(&1, shift))
    new_children = Enum.map(children, &shift_child(&1, shift, sheet_name))
    {"conditionalFormatting", new_attrs, new_children}
  end

  defp shift_sqref_attr({"sqref", value}, shift), do: {"sqref", SqrefShift.shift(value, shift)}
  defp shift_sqref_attr(other, _shift), do: other

  defp shift_child({"cfRule", attrs, children}, shift, sheet_name) do
    new_children = Enum.map(children, &shift_rule_child(&1, shift, sheet_name))
    {"cfRule", attrs, new_children}
  end

  defp shift_child(other, _shift, _sheet_name), do: other

  defp shift_rule_child({"formula", attrs, children}, shift, sheet_name) do
    new_children = Enum.map(children, &shift_formula_text(&1, shift, sheet_name))
    {"formula", attrs, new_children}
  end

  defp shift_rule_child(other, _shift, _sheet_name), do: other

  defp shift_formula_text(text, shift, sheet_name) when is_binary(text) do
    ("=" <> text)
    |> Tokenizer.tokenize()
    |> FormulaShift.apply(shift, sheet_name)
    |> Serializer.to_string()
    |> String.trim_leading("=")
  end

  defp shift_formula_text(other, _shift, _sheet_name), do: other
end
