//****************************************************************
// zexmlssutils  (Z Excel XML SpreadSheet Utils)
// Различные дополнительные утилитки для ZEXMLSS
// Накалякано в Мозыре в 2009 году
// Автор:  Неборак Руслан Владимирович (Ruslan V. Neborak)
// e-mail: avemey(мяу)tut(точка)by
// URL:    http://avemey.com
// Ver:    0.0.11
// Лицензия: zlib
// Last update: 2016.09.10
//----------------------------------------------------------------
// This software is provided "as-is", without any express or implied warranty.
// In no event will the authors be held liable for any damages arising from the
// use of this software.
//****************************************************************
unit zexmlssutils;

interface

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

uses
  {$IFNDEF FPC}
  Windows,
  {$ENDIF}
  SysUtils, UITypes, Types, Classes, Grids, Math,
  zexmlss, zsspxml, zesavecommon, graphics
  ;

function SaveXmlssToHtml(sheet: TZSheet; CodePageName: string = 'UTF-8'): string;

// needed for uniform save routines: zeSave*
// missed in pre-Unicode Dlephi and in FPC
function SplitString(const buffer: string; const delimeter: char): TStringDynArray;
{$IfDef DELPHI_UNICODE} overload; {$EndIf}

implementation
uses
  zenumberformats  //ConvertFormatNativeToXlsx / ConvertFormatXlsxToNative
{$IfDef DELPHI_UNICODE}
  ,
  StrUtils,  // stock SplitString(string, string) implementation
  AnsiStrings  // AnsiString targeted overloaded versions of Pos, Trim, etc
{$EndIf}
  ;

{$IFDEF DELPHI_UNICODE}
  {$DEFINE USE_STRUTILS_SPLIT_STRING}
{$ENDIF}

{$IFDEF VER200} // RAD Studio 2009
  {$UNDEF USE_STRUTILS_SPLIT_STRING} //There are no StrUtils.SplitString in D2009!!!!
{$ENDIF VER200}

function SplitString(const buffer: string; const delimeter: char): TStringDynArray;
{$IfDef USE_STRUTILS_SPLIT_STRING}
begin
   Result := StrUtils.SplitString(buffer, delimeter); // implicit typecast
end;
{$Else}
var
  i, from, till: integer;
  L: array of integer;
  _count, _maxcount: integer;

  procedure _add(num: integer);
  begin
    if (_count + 1 >= _maxcount) then
    begin
      inc(_maxcount, 10);
      setlength(L, _maxcount)
    end;
    L[_count] := num;
    inc(_count);
  end;

begin
  _maxcount := 20;
  _count := 0;
  SetLength(L, _maxcount);
  try
    for i := 1 to length(buffer) do
      if delimeter = buffer[i] then
        _add(i);
    _add(length(buffer) + 1);

    SetLength(Result, _count);

    from := 1;
    for i := 0 to _count - 1 do
    begin
      till := L[i];
      Result[i] := Copy(buffer, from, till - from);
      from := till + 1;
    end;
  finally
    setlength(L, 0);
  end;
end;

{$EndIf}

//Сохраняет страницу TZEXMLSS в поток в формате HTML
//Input:
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страницы
//      Title: string - Заголовок
//      Stream: TStream - поток
//      TextConverter: TAnsiToCodePageConverter - конвертер
//      CodePageName: string - имя кодировки
//Output:
//      0 - сохранение удалось
function SaveXmlssToHtml(sheet: TZSheet; CodePageName: string = 'UTF-8'): string;
var
  _xml: TZsspXMLWriterH;
  i, j, t, l,r: integer;
  NumTopLeft, NumArea: integer;
  s, value,numformat: string;
  Att: TZAttributesH;
  max_width: Real;
  strArray: TArray<string>;
  Stream: TStringStream;

function HTMLStyleTable(name: string; const Style: TZStyle): string;
var
  s: string;
  i, l: integer;

begin
  result := #13#10 + ' .' + name + '{'#13#10;
  for i := 0 to 3 do
  begin
    s := 'border-';
    l := 0;
    case i of
      0: s := s + 'left:';
      1: s := s + 'top:';
      2: s := s + 'right:';
      3: s := s + 'bottom:';
    end;
    s := s + '#' + ColorToHTMLHex(Style.Border[ TZBordersPos(i)].Color);
    if Style.Border[TZBordersPos(i)].Weight <> 0 then
      s := s + ' ' + IntToStr(Style.Border[TZBordersPos(i)].Weight) + 'px'
    else
        inc(l);
    case Style.Border[TZBordersPos(i)].LineStyle of
      ZEContinuous  : s := s + ' ' + 'solid';
      ZEHair        : s := s + ' ' + 'solid';
      ZEDot         : s := s + ' ' + 'dotted';
      ZEDashDotDot  : s := s + ' ' + 'dotted';
      ZEDash        : s := s + ' ' + 'dashed';
      ZEDashDot     : s := s + ' ' + 'dashed';
      ZESlantDashDot: s := s + ' ' + 'dashed';
      ZEDouble      : s := s + ' ' + 'double';
      else
        inc(l);
    end;
    s := s + ';';
    if l <> 2 then
      result := result + s+#13#10;
  end;
  result := result + 'background:#' +  ColorToHTMLHex(Style.BGColor) + ';}';
end;

function HTMLStyleFont(name: string; const Style: TZStyle): string;
begin
  result := #13#10 + ' .' + name + '{'#13#10;
  result := result + 'color:#' + ColorToHTMLHex(Style.Font.Color) + ';';
  result := result + 'font-size:' + inttostr(Style.Font.Size) + 'px;';
  result := result + 'font-family:' + Style.Font.Name + ';}';
end;

begin
  result := '';
  Stream := TStringStream.Create('', TEncoding.UTF8);
  try
    _xml := TZsspXMLWriterH.Create();
    with _xml do
    begin
      TabLength := 1;
      BeginSaveToStream(Stream);
      //
      Attributes.Clear();
      WriteRaw('<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">', true, false);
      WriteTagNode('HTML',true, true, false);
      WriteTagNode('HEAD',true, true, false);
      WriteTag('TITLE', sheet.Title, true, false, false);

      //styles
      s := 'body {';
      s := s + 'background:#' + ColorToHTMLHex(sheet.WorkBook.Styles.DefaultStyle.BGColor) + ';';
      s := s + 'color:#' + ColorToHTMLHex(sheet.WorkBook.Styles.DefaultStyle.Font.Color) + ';';
      s := s + 'font-size:' + inttostr(sheet.WorkBook.Styles.DefaultStyle.Font.Size) + 'px;';
      s := s + 'font-family:' + sheet.WorkBook.Styles.DefaultStyle.Font.Name + ';}';

      s := s + HTMLStyleTable('T19', sheet.WorkBook.Styles.DefaultStyle);
      s := s + HTMLStyleFont('F19', sheet.WorkBook.Styles.DefaultStyle);

      for i := 0 to sheet.WorkBook.Styles.Count - 1 do
      begin
        s := s + HTMLStyleTable('T' + IntToStr(i + 20), sheet.WorkBook.Styles[i]);
        s := s +  HTMLStyleFont('F' + IntToStr(i + 20), sheet.WorkBook.Styles[i]);
      end;

      WriteTag('STYLE', s, true, true, false);
      Attributes.Add('HTTP-EQUIV', 'CONTENT-TYPE');

      s := '';
      if trim(CodePageName) > '' then
        s := '; CHARSET='+ CodePageName;
      Attributes.Add('CONTENT', 'TEXT/HTML'+s);
      WriteTag('META','', true, false, false);
      WriteEndTagNode(); // HEAD

      max_width := 0.0;
      for i := 0 to sheet.ColCount-1 do begin
        max_width:=max_width+sheet.ColWidths[i];
      end;

      //BODY
      Attributes.Clear();
      WriteTagNode('BODY',true, true, false);

      //Table
      Attributes.Clear();
      Attributes.Add('cellSpacing', '0');
      Attributes.Add('border','0');
      Attributes.Add('width', FloatToStr(max_width).Replace(',','.'));
      WriteTagNode('TABLE',true, true, false);

      Att := TZAttributesH.Create();
      Att.Clear();
      for i := 0 to sheet.RowCount - 1 do
      begin
        Attributes.Clear();
        Attributes.Add('height', floattostr(sheet.RowHeights[i]).Replace(',','.'));
        WriteTagNode('TR',true, true, true);
        Attributes.Clear();
        for j := 0 to sheet.ColCount - 1 do begin
          NumTopLeft := sheet.MergeCells.InLeftTopCorner(j, i);
          NumArea := sheet.MergeCells.InMergeRange(j, i);
          // если ячейка входит в объединённые области и не является
          // верхней левой ячейков в этой области - пропускаем её
          if not ((NumArea >= 0) and (NumTopLeft = -1)) then begin
            Attributes.Clear();
            if NumTopLeft >= 0 then begin
              {tut}
              t := sheet.MergeCells.Items[NumTopLeft].Right -
                   sheet.MergeCells.Items[NumTopLeft].Left;
              if t > 0 then
                Attributes.Add('colspan', InttOstr(t+1));
              t := sheet.MergeCells.Items[NumTopLeft].Bottom -
                   sheet.MergeCells.Items[NumTopLeft].Top;
              if t > 0 then
                Attributes.Add('rowspan', InttOstr(t+1));
            end;
            t := sheet.Cell[j,i].CellStyle;
            if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHCenter then
              Attributes.Add('align', 'center') else
            if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHRight then
              Attributes.Add('align', 'right') else
            if sheet.WorkBook.Styles[t].Alignment.Horizontal = ZHJustify then
              Attributes.Add('align', 'justify');
            numformat := sheet.WorkBook.Styles[t].NumberFormat;
            Attributes.Add('class', 'T'+IntToStr(t + 20));
            Attributes.Add('width', inttostr(sheet.Columns[j].WidthPix)+'px');

            WriteTagNode('TD', true, false, false);
            Attributes.Clear();
            Att.Clear();
            Att.Add('class', 'F' + IntToStr(t + 20));
            if fsbold in sheet.WorkBook.Styles[t].Font.Style then
              WriteTagNode('B', false, false, false);
            if fsItalic in sheet.WorkBook.Styles[t].Font.Style then
              WriteTagNode('I', false, false, false);
            if fsUnderline in sheet.WorkBook.Styles[t].Font.Style then
              WriteTagNode('U', false, false, false);
            if fsStrikeOut in sheet.WorkBook.Styles[t].Font.Style then
              WriteTagNode('S', false, false, false);

            l := Length(sheet.Cell[j, i].Href);
            if l > 0 then
            begin
              Attributes.Add('href', sheet.Cell[j, i].Href);
              //target?
              WriteTagNode('A', false, false, false);
              Attributes.Clear();
            end;

            value := sheet.Cell[j, i].Data;

            //value := value.Replace(#13#10, '<br>');
            case sheet.Cell[j, i].CellType of
              TZCellType.ZENumber: begin
                r := numformat.IndexOf('.');
                if r > -1 then begin
                    value := FloatToStrF(
                        sheet.Cell[j, i].AsDouble,
                        ffNumber,
                        12,
                        Min(4, Max(0, numformat.Substring(r).Length-1)));
                end else begin
                    value := FloatToStr(sheet.Cell[j, i].AsDouble);
                end;
              end;
              TZCellType.ZEDateTime : begin
                // todo: make datetimeformat from cell NumberFormat
                value := FormatDateTime('dd.mm.yyyy', sheet.Cell[j, i].AsDateTime);
              end;
            end;
            strArray := value.Split([#13,#10], TStringSplitOptions.ExcludeEmpty);
            for r := 0 to Length(strArray)-1 do begin
                if r > 0 then WriteTag('BR','');
                WriteTag('FONT', strArray[r], Att, false, false, true);
            end;

            if l > 0 then
              WriteEndTagNode(); // A

            if fsbold in sheet.WorkBook.Styles[t].Font.Style then
              WriteEndTagNode(); // B
            if fsItalic in sheet.WorkBook.Styles[t].Font.Style then
              WriteEndTagNode(); // I
            if fsUnderline in sheet.WorkBook.Styles[t].Font.Style then
              WriteEndTagNode(); // U
            if fsStrikeOut in sheet.WorkBook.Styles[t].Font.Style then
              WriteEndTagNode(); // S
            WriteEndTagNode(); // TD
          end;

        end;
        WriteEndTagNode(); // TR
      end;

      WriteEndTagNode(); // BODY
      WriteEndTagNode(); // HTML
      EndSaveTo();
      Result := Stream.DataString
    end;
    FreeAndNil(Att);
  finally
    FreeAndNil(_xml);
    Stream.Free();
  end;
end;

end.
