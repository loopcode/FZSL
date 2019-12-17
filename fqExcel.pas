unit fqExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,
  ComObj;

function todoExcel(AFileName: string; AMainTitle: string = ''; ASubTitle: string = ''; AHeadCount: Integer = 1;
                   AReportFooter: string = ''; ALeftFooter: string = ''; ARightFooter: string = '&9第 &P 页，共 &N 页'): Boolean;

implementation

function IntToLetter(ACol: Integer): string;
begin
  Result := '';
  if ((ACol - 1) div 26 > 0) then
   Result := Chr($40 + (ACol - 1) div 26);
  Result := Result + Chr($41 + (ACol - 1) mod 26);
end;

function todoExcel(AFileName: string; AMainTitle: string; ASubTitle: string; AHeadCount: Integer;
                   AReportFooter: string; ALeftFooter: string; ARightFooter: string): Boolean;
const
  xlTop    = -4160;
  xlDown   = -4121;
  xlLeft   = -4131;
  xlRight  = -4152;
  xlCenter = -4108;
var
  Excel, WorkBook, WorkSheet: Variant;
  MaxRow, MaxColumn: LongInt;
begin
  if AMainTitle = '' then
    AMainTitle := ChangeFileExt(ExtractFileName(AFileName), '');

  // 创建 Excel
  try
    Excel := CreateOLEObject('Excel.Application');
  except
    Application.MessageBox('无法正确联接您计算机中的 Microsoft Excel，请确认是否正确安装！', '创建 OLE 错误', MB_ICONQUESTION	);
    Result := False;
    Exit;
  end;

  // 打开文件
  Result := True;
  WorkBook := Excel.WorkBooks.Open(AFileName);
  WorkSheet := WorkBook.WorkSheets.Item[1];

  // 最大行、最大列
  MaxRow := WorkSheet.UsedRange.Rows.Count;
  MaxColumn := WorkSheet.UsedRange.Columns.Count;

  // 添加标题行
  WorkSheet.Rows['1:1'].Select;
  Excel.Selection.Insert(xlDown);
  Inc(MaxRow);
  Inc(AHeadCount);
  if (ASubTitle <> '') then
  begin
    Excel.Selection.Insert(xlDown);
    Inc(MaxRow);
    Inc(AHeadCount);
  end;

  // 生成主标题
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Merge;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].HorizontalAlignment := xlCenter;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].VerticalAlignment := xlCenter;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Font.Name := '黑体';
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Font.Size := 24;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Value := AMainTitle;
  WorkSheet.Rows['1:1'].RowHeight := 60;

  // 生成副标题
  if (ASubTitle <> '') then
  begin
    WorkSheet.Range['A2'].HorizontalAlignment := xlLeft;
    WorkSheet.Range['A2'].VerticalAlignment := xlCenter;
    WorkSheet.Range['A2'].Font.Name := '宋体';
    WorkSheet.Range['A2'].Font.Size := 10;
    WorkSheet.Range['A2'].Value := ASubTitle;
    WorkSheet.Rows['2:2'].RowHeight := 20;
  end;

  // 生成报表尾
  if (AReportFooter <> '') then
  begin
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].HorizontalAlignment := xlLeft;
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].VerticalAlignment := xlCenter;
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].Font.Name := '宋体';
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].Font.Size := 10;
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].Value := AReportFooter;
    WorkSheet.Rows[IntToStr(MaxRow + 1) + ':' + IntToStr(MaxRow + 1)].RowHeight := 20;
  end;

  // 固定表头
  WorkSheet.PageSetup.PrintTitleRows := '$1:$' + IntToStr(AHeadCount);

  // 设置页脚
  if (ALeftFooter <> '') then
    WorkSheet.PageSetup.LeftFooter := ALeftFooter;
  if (ARightFooter <> '') then
    WorkSheet.PageSetup.RightFooter := ARightFooter;


  // 保存文件
  try
    WorkSheet.Range['A1:A1'].Select;
    WorkBook.Save;
    WorkBook.Close();
  except
    Application.MessageBox('不能正确保存 Excel 文件，可能是该文件已被其他程序打开，或系统错误。', '警告', MB_ICONQUESTION);
    Result := False;
  end;

  // 释放 Excel
  try
    WorkBook := Null;
    Excel := Null;
  except
    Result := False;
  end;
end;

end.
