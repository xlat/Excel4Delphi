
# Excel4Delphi
Read, Write excel 2002/2003 XML (SpreadsheetML / XML Spreadsheet) library.

fork from https://github.com/Avemey/zexmlss

## Exchamples

```pas
// Creating new workbook
var workBook: TZWorkBook;
...
workBook := TZWorkBook.Create();
try
  workBook.Sheets.Add('My sheet');
  workBook.Sheets[0].SetDimensions(10, 20);
  workBook.Sheets[0].CellRef['A', 0].AsString := 'Hello';
  workBook.Sheets[0].RangeRef['A', 0, 'B', 2].Merge();
  workBook.SaveToFile('file.xlsx');
finally
  workBook.Free();
end
```

```pas
// Editing exists workbook
var workBook: TZWorkBook;
...
workBook := TZWorkBook.Create();
try
  workBook.LoadFromFile('file.xlsx');
  workBook.Sheets[0].CellRef['A', 0].AsString := 'Hello';
  workBook.Sheets[0].CellRef['A', 0].FontStyle := [fsBold];
  workBook.SaveToFile('file.xlsx');
finally
  workBook.Free();
end
```

## TODO
* TZMergeCells = class -> TZMergedCells = class(TObjectList<TZMergeArea>)
* TZSheets = class -> TZSheets = class(TObjectList<TZSheet>)
* TZStyles = class -> TZStyles = class(TObjectList<TZStyle>)