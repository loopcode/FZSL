unit fqExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,
  ComObj;

function todoExcel(AFileName: string; AMainTitle: string = ''; ASubTitle: string = ''; AHeadCount: Integer = 1;
                   AReportFooter: string = ''; ALeftFooter: string = ''; ARightFooter: string = '&9�� &P ҳ���� &N ҳ'): Boolean;

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

  // ���� Excel
  try
    Excel := CreateOLEObject('Excel.Application');
  except
    Application.MessageBox('�޷���ȷ������������е� Microsoft Excel����ȷ���Ƿ���ȷ��װ��', '���� OLE ����', MB_ICONQUESTION	);
    Result := False;
    Exit;
  end;

  // ���ļ�
  Result := True;
  WorkBook := Excel.WorkBooks.Open(AFileName);
  WorkSheet := WorkBook.WorkSheets.Item[1];

  // ����С������
  MaxRow := WorkSheet.UsedRange.Rows.Count;
  MaxColumn := WorkSheet.UsedRange.Columns.Count;

  // ��ӱ�����
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

  // ����������
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Merge;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].HorizontalAlignment := xlCenter;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].VerticalAlignment := xlCenter;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Font.Name := '����';
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Font.Size := 24;
  WorkSheet.Range['A1:' + IntToLetter(MaxColumn) + '1'].Value := AMainTitle;
  WorkSheet.Rows['1:1'].RowHeight := 60;

  // ���ɸ�����
  if (ASubTitle <> '') then
  begin
    WorkSheet.Range['A2'].HorizontalAlignment := xlLeft;
    WorkSheet.Range['A2'].VerticalAlignment := xlCenter;
    WorkSheet.Range['A2'].Font.Name := '����';
    WorkSheet.Range['A2'].Font.Size := 10;
    WorkSheet.Range['A2'].Value := ASubTitle;
    WorkSheet.Rows['2:2'].RowHeight := 20;
  end;

  // ���ɱ���β
  if (AReportFooter <> '') then
  begin
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].HorizontalAlignment := xlLeft;
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].VerticalAlignment := xlCenter;
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].Font.Name := '����';
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].Font.Size := 10;
    WorkSheet.Range['A' + IntToStr(MaxRow + 1)].Value := AReportFooter;
    WorkSheet.Rows[IntToStr(MaxRow + 1) + ':' + IntToStr(MaxRow + 1)].RowHeight := 20;
  end;

  // �̶���ͷ
  WorkSheet.PageSetup.PrintTitleRows := '$1:$' + IntToStr(AHeadCount);

  // ����ҳ��
  if (ALeftFooter <> '') then
    WorkSheet.PageSetup.LeftFooter := ALeftFooter;
  if (ARightFooter <> '') then
    WorkSheet.PageSetup.RightFooter := ARightFooter;


  // �����ļ�
  try
    WorkSheet.Range['A1:A1'].Select;
    WorkBook.Save;
    WorkBook.Close();
  except
    Application.MessageBox('������ȷ���� Excel �ļ��������Ǹ��ļ��ѱ���������򿪣���ϵͳ����', '����', MB_ICONQUESTION);
    Result := False;
  end;

  // �ͷ� Excel
  try
    WorkBook := Null;
    Excel := Null;
  except
    Result := False;
  end;
end;

end.
