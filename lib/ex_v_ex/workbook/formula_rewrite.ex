defmodule ExVEx.Workbook.FormulaRewrite do
  @moduledoc """
  Applies one token-level rewrite to every formula in a workbook: cell
  formulas, conditional-formatting and data-validation formulas, defined
  names, and the calculated-column and totals-row formulas of table
  parts.

  The rewrite function receives the token list and a context map with
  `:sheet` (the sheet name, or `nil` for defined names and table parts)
  and `:coord` (the cell coordinate, or `nil`). Only formulas whose
  token list changes are written back.
  """

  alias ExVEx.Formula.{Serializer, Token, Tokenizer}
  alias ExVEx.OOXML.Table.Column
  alias ExVEx.OOXML.Workbook, as: WorkbookXml
  alias ExVEx.OOXML.Worksheet.Editable
  alias ExVEx.Workbook
  alias ExVEx.Workbook.TableParts

  @type context :: %{sheet: String.t() | nil, coord: {pos_integer(), pos_integer()} | nil}
  @type rewrite :: ([Token.t()], context() -> [Token.t()])

  @spec apply(Workbook.t(), rewrite()) :: Workbook.t()
  def apply(%Workbook{} = book, fun) when is_function(fun, 2) do
    book
    |> rewrite_sheets(fun)
    |> rewrite_defined_names(fun)
    |> rewrite_table_parts(fun)
  end

  defp rewrite_sheets(book, fun) do
    Enum.reduce(book.workbook.sheets, book, fn sheet_ref, acc ->
      with {:ok, path} <- Workbook.sheet_path(acc, sheet_ref.name),
           {:ok, editable, acc} <- Workbook.fetch_sheet_tree(acc, path) do
        rewrite_sheet(acc, path, editable, sheet_ref.name, fun)
      else
        _ -> acc
      end
    end)
  end

  defp rewrite_sheet(book, path, editable, sheet_name, fun) do
    sheet_fun = fn tokens, coord -> fun.(tokens, %{sheet: sheet_name, coord: coord}) end

    case Editable.rewrite_formulas(editable, sheet_fun) do
      {_unchanged, false} -> book
      {new_editable, true} -> Workbook.put_sheet_tree(book, path, new_editable)
    end
  end

  defp rewrite_defined_names(book, fun) do
    context = %{sheet: nil, coord: nil}

    new_names =
      Enum.map(book.workbook.defined_names, fn name ->
        new_ref = rewrite_text(name.reference, fun, context)
        if new_ref == name.reference, do: name, else: %{name | reference: new_ref}
      end)

    if new_names == book.workbook.defined_names do
      book
    else
      workbook = %{book.workbook | defined_names: new_names}
      xml = WorkbookXml.serialize_into(workbook, Map.fetch!(book.parts, book.workbook_path))

      %{
        book
        | workbook: workbook,
          parts: Map.put(book.parts, book.workbook_path, xml),
          calc_dirty: true
      }
    end
  end

  defp rewrite_table_parts(book, fun) do
    context = %{sheet: nil, coord: nil}

    Enum.reduce(TableParts.list(book), book, fn entry, acc ->
      new_columns = Enum.map(entry.table.columns, &rewrite_column(&1, fun, context))

      if new_columns == entry.table.columns do
        acc
      else
        TableParts.put(acc, entry, %{entry.table | columns: new_columns})
      end
    end)
  end

  defp rewrite_column(%Column{children: children} = column, fun, context) do
    new_children =
      Enum.map(children, fn
        {tag, attrs, inner} when tag in ["calculatedColumnFormula", "totalsRowFormula"] ->
          text = inner |> Enum.filter(&is_binary/1) |> Enum.join("")
          new_text = rewrite_text(text, fun, context)
          if new_text == text, do: {tag, attrs, inner}, else: {tag, attrs, [new_text]}

        other ->
          other
      end)

    %{column | children: new_children}
  end

  defp rewrite_text("", _fun, _context), do: ""

  defp rewrite_text(text, fun, context) do
    text |> Tokenizer.tokenize() |> fun.(context) |> Serializer.to_string()
  end
end
