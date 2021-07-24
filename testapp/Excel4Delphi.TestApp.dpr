program Excel4Delphi.TestApp;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Excel4Delphi
  ;

const xlsx_dir: string = '..\..\..\testapp\';

var
  workBook: TZWorkBook;
  xlsx_file: string;
begin
  xlsx_file := xlsx_dir + 'test.xlsx';
  //xlsx_file := 'C:\Users\gawri\Desktop\Book_2.xlsx';

  try
    workBook := TZWorkBook.Create();
    try
      workBook.Load(xlsx_file);

      // todo: any changes

      workBook.SaveToFile(xlsx_dir + 'test2.xlsx');
    finally
      workBook.Free();
    end;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
