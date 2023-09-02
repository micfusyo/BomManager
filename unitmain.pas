unit unitMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Grids, ComCtrls, unitSelf, LCLType, StrUtils;

type

  { TFormMain }

  TFormMain = class(TForm)
    ButtonCheck2: TButton;
    ButtonCheck1: TButton;
    ButtonOrder: TButton;
    ButtonRenum1: TButton;
    ButtonRenum2: TButton;
    ButtonCopy1: TButton;
    ButtonCopy2: TButton;
    ButtonInsert2: TButton;
    ButtonRemove2: TButton;
    ButtonInsert1: TButton;
    ButtonRemove1: TButton;
    ButtonUpdate: TButton;
    ButtonFormat1: TButton;
    ButtonLoad1: TButton;
    ButtonSave1: TButton;
    ButtonLoad2: TButton;
    ButtonSave2: TButton;
    ButtonExit: TButton;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Edit1: TEdit;
    Edit2: TEdit;
    EditFind: TEdit;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    ProgressBar: TProgressBar;
    SaveDialog: TSaveDialog;
    Splitter1: TSplitter;
    StringGrid1: TStringGrid;
    StringGrid2: TStringGrid;
    procedure ButtonCheck1Click(Sender: TObject);
    procedure ButtonCheck2Click(Sender: TObject);
    procedure ButtonCopy1Click(Sender: TObject);
    procedure ButtonCopy2Click(Sender: TObject);
    procedure ButtonExitClick(Sender: TObject);
    procedure ButtonFormat1Click(Sender: TObject);
    procedure ButtonInsert1Click(Sender: TObject);
    procedure ButtonInsert2Click(Sender: TObject);
    procedure ButtonLoad1Click(Sender: TObject);
    procedure ButtonLoad2Click(Sender: TObject);
    procedure ButtonOrderClick(Sender: TObject);
    procedure ButtonRemove1Click(Sender: TObject);
    procedure ButtonRemove2Click(Sender: TObject);
    procedure ButtonRenum1Click(Sender: TObject);
    procedure ButtonRenum2Click(Sender: TObject);
    procedure ButtonSave1Click(Sender: TObject);
    procedure ButtonSave2Click(Sender: TObject);
    procedure ButtonUpdateClick(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure Edit1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EditFindKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure StringGrid2Click(Sender: TObject);
  private

  public

  end;

var
  FormMain: TFormMain;

implementation

{$R *.lfm}

{ TFormMain }

procedure TFormMain.ButtonExitClick(Sender: TObject);
begin
  Close;
end;

procedure TFormMain.ButtonCopy1Click(Sender: TObject);
var
  Str: String;
  Sel, Row1, Row2: Integer;
begin
  Row1:= StringGrid1.Selection.Top;
  Row2:= StringGrid2.Selection.Top;
  Str:= 'StringGrid1 selected row =' + StringGrid1.Selection.Top.ToString + constCrLf;
  Str:= Str + 'StringGrid2 selected row =' + StringGrid2.Selection.Top.ToString + constCrLf;
  Str:= Str + 'Confirm Copy row from StringGrid1 to StringGrid2?';
  Sel:= MessageDlg(Str, mtConfirmation, mbOKCancel, 0);
  if Sel = mrOK then
  begin
    StringGrid2.Cells[StkColSyoPart, Row2]:= StringGrid1.Cells[BomColSyoPart, Row1];
    StringGrid2.Cells[StkColName, Row2]:= StringGrid1.Cells[BomColName, Row1];
    StringGrid2.Cells[StkColManufacturerPart, Row2]:= StringGrid1.Cells[BomColManufacturerPart, Row1];
    StringGrid2.Cells[StkColManufacturer, Row2]:= StringGrid1.Cells[BomColManufacturer, Row1];
    StringGrid2.Cells[StkColSupplier, Row2]:= StringGrid1.Cells[BomColSupplier, Row1];
    StringGrid2.Cells[StkColSupplierPart, Row2]:= StringGrid1.Cells[BomColSupplierPart, Row1];
    StringGrid2.Cells[StkColSupplierLink, Row2]:= StringGrid1.Cells[BomColSupplierLink, Row1];
    StringGrid2.Cells[StkColPrice, Row2]:= StringGrid1.Cells[BomColPrice, Row1];
    StringGrid2.Cells[StkColPins, Row2]:= StringGrid1.Cells[BomColPins, Row1];
    StringGrid2.Cells[StkColRemark, Row2]:= StringGrid1.Cells[BomColRemark, Row1];
    StringGrid2.Cells[StkColReplacement, Row2]:= StringGrid1.Cells[BomColReplacement, Row1];
    StringGrid2.AutoSizeColumns;
  end;
end;

procedure TFormMain.ButtonCheck1Click(Sender: TObject);
var
  idx, Num1, Num2: Integer;
  Str1, Str2: String;
begin
  ProgressBar.Position:= 0;
  ProgressBar.Min:= 0;
  ProgressBar.Max:= StringGrid2.RowCount-1;
  with StringGrid2 do
  begin
    for idx:= 2 to RowCount-2 do
    begin
      ProgressBar.Position:= idx-1;
      // Stage 1
      Str1:= LeftStr(Cells[StkColSyoPart, idx], 3);
      Str2:= LeftStr(Cells[StkColSyoPart, idx+1], 3);
      try
        Num1:= Str1.ToInteger;
        Num2:= Str2.ToInteger;
      except
        ShowMessage('Stage1->Not a number ERROR: Row = ' + IntToStr(idx));
        ProgressBar.Position:= 0;
        TopRow:= idx-1;
        Exit;
      end;
      if (Num1 > Num2) then
      begin
        ShowMessage('Stage1->Bigger then last number ERROR: Row = ' + IntToStr(idx));
        ProgressBar.Position:= 0;
        TopRow:= idx-1;
        Exit;
      end;
      // Stage 2
      Str1:= LeftStr(Cells[StkColSyoPart, idx], 13);
      Str2:= LeftStr(Cells[StkColSyoPart, idx+1], 13);
      if (Str1 = Str2) then
      begin
        Str1:= RightStr(Cells[StkColSyoPart, idx], 5);
        Str2:= RightStr(Cells[StkColSyoPart, idx+1], 5);
        try
          Num1:= Str1.ToInteger;
          Num2:= Str2.ToInteger;
        except
          ShowMessage('Stage2->Not a number ERROR: Row = ' + IntToStr(idx));
          ProgressBar.Position:= 0;
          TopRow:= idx-1;
          Exit;
        end;
        if (Num1 > Num2) then
        begin
          ShowMessage('Stage2->Bigger then last number ERROR: Row = ' + IntToStr(idx));
          ProgressBar.Position:= 0;
          TopRow:= idx-1;
          Exit;
        end;
      end
      // Stage 3
      else
      begin
        Str1:= MidStr(Cells[StkColSyoPart, idx], 5, 2) + MidStr(Cells[StkColSyoPart, idx], 8, 2) + MidStr(Cells[StkColSyoPart, idx], 11, 3);
        Str2:= MidStr(Cells[StkColSyoPart, idx], 5, 2) + MidStr(Cells[StkColSyoPart, idx], 8, 2) + MidStr(Cells[StkColSyoPart, idx], 11, 3);
        try
          Num1:= Str1.ToInteger;
          Num2:= Str2.ToInteger;
        except
          ShowMessage('Stage3->Not a number ERROR: Row = ' + IntToStr(idx));
          ProgressBar.Position:= 0;
          TopRow:= idx-1;
          Exit;
        end;
        if (Num1 > Num2) then
        begin
          ShowMessage('Stage3->Bigger then last number ERROR: Row = ' + IntToStr(idx));
          ProgressBar.Position:= 0;
          TopRow:= idx-1;
          Exit;
        end;
      end;
    end;
    ShowMessage('OK');
    ProgressBar.Position:= 0;
  end;
end;

procedure TFormMain.ButtonCheck2Click(Sender: TObject);
var
  Idx1, Idx2: Integer;
  Str1, Str2: String;
begin
  with StringGrid2 do
  begin
    for Idx1:= 2 to RowCount-2 do
    begin
      Str1:= Trim(Cells[StkColSupplierPart, Idx1]);
      if (Str1 = '') then
        Continue;
      for Idx2:= Idx1+1 to RowCount-1 do
      begin
        Str2:= Trim(Cells[StkColSupplierPart, Idx2]);
        if (Str2 = '') then
          Continue;
        if (Str1 = Str2) then
        begin
          ShowMessage('Same supplier part number ERROR: Row1= ' + IntToStr(Idx1) + ' Row2= ' + IntToStr(Idx2));
          TopRow:= Idx1;
          Exit;
        end;
      end;
    end;
  end;
  ShowMessage('OK');
end;

procedure TFormMain.ButtonCopy2Click(Sender: TObject);
var
  Str: String;
  Sel, Row1, Row2: Integer;
begin
  Row1:= StringGrid1.Selection.Top;
  Row2:= StringGrid2.Selection.Top;
  Str:= 'StringGrid1 selected row =' + StringGrid1.Selection.Top.ToString + constCrLf;
  Str:= Str + 'StringGrid2 selected row =' + StringGrid2.Selection.Top.ToString + constCrLf;
  Str:= Str + 'Confirm Copy row from StringGrid2 to StringGrid1?';
  Sel:= MessageDlg(Str, mtConfirmation, mbOKCancel, 0);
  if Sel = mrOK then
  begin
    StringGrid1.Cells[BomColSyoPart, Row1]:= StringGrid2.Cells[StkColSyoPart, Row2];
    StringGrid1.Cells[BomColName, Row1]:= StringGrid2.Cells[StkColName, Row2];
    StringGrid1.Cells[BomColManufacturerPart, Row1]:= StringGrid2.Cells[StkColManufacturerPart, Row2];
    StringGrid1.Cells[BomColManufacturer, Row1]:= StringGrid2.Cells[StkColManufacturer, Row2];
    StringGrid1.Cells[BomColSupplier, Row1]:= StringGrid2.Cells[StkColSupplier, Row2];
    StringGrid1.Cells[BomColSupplierPart, Row1]:= StringGrid2.Cells[StkColSupplierPart, Row2];
    StringGrid1.Cells[BomColSupplierLink, Row1]:= StringGrid2.Cells[StkColSupplierLink, Row2];
    StringGrid1.Cells[BomColPrice, Row1]:= StringGrid2.Cells[StkColPrice, Row2];
    StringGrid1.Cells[BomColPins, Row1]:= StringGrid2.Cells[StkColPins, Row2];
    StringGrid1.Cells[BomColRemark, Row1]:= StringGrid2.Cells[StkColRemark, Row2];
    StringGrid1.Cells[BomColReplacement, Row1]:= StringGrid2.Cells[StkColReplacement, Row2];
    StringGrid1.AutoSizeColumns;
  end;
end;

procedure TFormMain.ButtonFormat1Click(Sender: TObject);
var
  IdxCol: integer;
  IdxRow: integer;
  Qty: integer;
  Sub: double;
  Str: string;
  ListStr: TStringList;
begin
  ListStr:= TStringList.Create;
  with StringGrid1 do
  begin
    // Add Col of PinTotal, Remark, Replacement
    InsertColRow(true, 11);
    StringGrid1.MoveColRow(true, 11, 10);
    InsertColRow(true, 11);
    InsertColRow(true, 11);
    // Add Col of SubTotal
    InsertColRow(true, 10);
    // Add Col of SupplierLink
    InsertColRow(true, 9);
    // Add Col of ID, SYOPart
    InsertColRow(true, 1);
    InsertColRow(true, 1);
    // Add Row of Title, Pcb
    InsertColRow(false, 1);
    InsertColRow(false, 1);
    // Set Title and Pcb
    for IdxCol:= 0 to ColCount-1 do
    begin
      Cells[IdxCol, 0]:= TilArray[IdxCol];
      Cells[IdxCol, 1]:= TilArray[IdxCol];
      Cells[IdxCol, 2]:= PcbArray[IdxCol];
    end;
    for IdxRow:= 3 to RowCount-1 do
    begin
      // ID
      Cells[BomColID, IdxRow]:= IntToStr(IdxRow-1);
      // Designator
      if not (Cells[BomColDesignator, IdxRow] = '') then
      begin
        ListStr.Clear;
        ListStr.CommaText:= Cells[BomColDesignator, IdxRow];
        if ListStr.Count > 2 then
        begin
          Str:= ListStr.Strings[0] + '-' + ListStr.Strings[ListStr.Count-1];
          Cells[BomColDesignator, IdxRow]:= Str;
        end;
      end;
      // Quantity
      if (Cells[BomColQuantity, IdxRow] = '') then
      begin
        Qty:= 0;
        Cells[BomColQuantity, IdxRow]:= '0';
      end
      else
      begin
        Qty:= Cells[BomColQuantity, IdxRow].ToInteger;
      end;
      // Cells[BomColQuantity, IdxRow]:= '=F$2*' + IntToStr(Qty);
      // Supplier Link
      Str:= Cells[BomColSupplierPart, IdxRow];
      if not (Str = '') then
      begin
        if (Copy(Str, 1, 1) = 'C') then
        begin
          // Str:= SetStringToHyperLink(Str, Cells[BomColSupplier, IdxRow]);
          Str:= Cells[BomColSupplier, IdxRow];
          Cells[BomColSupplierLink, IdxRow]:= Str;
        end;
      end;
      // Price
      if (Cells[BomColPrice, IdxRow] = '') then
        Cells[BomColPrice, IdxRow]:= '0.001';
      // SubTotal
      // Cells[BomColSubTotal, IdxRow]:= '=F'+ IntToStr(IdxRow) + '*L' + IntToStr(IdxRow);
      // Cells[BomColSubTotal, IdxRow]:= IntToStr(Qty * Cells[BomColPrice, IdxRow].ToExtended);
      Sub:= Cells[BomColPrice, IdxRow].ToDouble;
      Sub:= Sub * Qty;
      Cells[BomColSubTotal, IdxRow]:= FloatToStr(Sub);
      // PinTotal
      // Cells[BomColPinTotal, IdxRow]:= '=F'+ IntToStr(IdxRow) + '*N' + IntToStr(IdxRow);
      if (Cells[BomColPins, IdxRow] = '') then
      begin
        Cells[BomColPins, IdxRow]:= '0';
        Cells[BomColPinTotal, IdxRow]:= '0';
      end
      else
      begin
        Cells[BomColPinTotal, IdxRow]:= IntToStr(Qty * Cells[BomColPins, IdxRow].ToInteger);
      end;
    end;
    RowCount:= RowCount + 1;
    StringGridReNumber(StringGrid1);
  end;
  ListStr.Free;
end;

procedure TFormMain.ButtonInsert1Click(Sender: TObject);
var
  Str: String;
  Sel, Row: Integer;
begin
  Row:= StringGrid1.Selection.Top;
  Str:= 'StringGrid1 selected row =' + StringGrid1.Selection.Top.ToString + constCrLf;
  Str:= Str + 'Confirm insertion of row into stringgrid1?';
  Sel:= MessageDlg(Str, mtConfirmation, mbOKCancel, 0);
  if Sel = mrOK then
    StringGrid1.InsertColRow(false, Row);
end;

procedure TFormMain.ButtonInsert2Click(Sender: TObject);
var
  Str: String;
  Sel, Row: Integer;
begin
  Row:= StringGrid2.Selection.Top;
  Str:= 'StringGrid2 selected row =' + StringGrid2.Selection.Top.ToString + constCrLf;
  Str:= Str + 'Confirm insertion of row into stringgrid2?';
  Sel:= MessageDlg(Str, mtConfirmation, mbOKCancel, 0);
  if Sel = mrOK then
    StringGrid2.InsertColRow(false, Row);
end;

procedure TFormMain.ButtonLoad1Click(Sender: TObject);
var
  str: string;
begin
  StringGrid1.Clear;
  StringGrid1.ColCount:= 5;
  StringGrid1.RowCount:= 5;
  OpenDialog.FileName:= '';
  if OpenDialog.Execute then
  begin
    str:= ExtractFileExt(OpenDialog.FileName);
    if str = '.csv' then
    begin
      StringGrid1.LoadFromCSVFile(OpenDialog.FileName, constTab, true, 0, true);
      StringGrid1.AutoSizeColumns;
    end
    else if str = '.xlsx' then
    begin
      StringGridLoadFromXLSFile(StringGrid1, OpenDialog.FileName, ProgressBar);
      StringGridReNumber(StringGrid1);
    end;
  end;
end;

procedure TFormMain.ButtonLoad2Click(Sender: TObject);
var
  str: string;
begin
  StringGrid2.Clear;
  StringGrid2.ColCount:= 5;
  StringGrid2.RowCount:= 5;
  OpenDialog.FileName:= '';
  if OpenDialog.Execute then
  begin
    str:= ExtractFileExt(OpenDialog.FileName);
    if str = '.csv' then
    begin
      StringGrid2.LoadFromCSVFile(OpenDialog.FileName, constTab, true, 0, true);
      StringGrid2.AutoSizeColumns;
    end
    else if str = '.xlsx' then
    begin
      StringGridLoadFromXLSFile(StringGrid2, OpenDialog.FileName, ProgressBar);
      StringGridReNumber(StringGrid2);
    end;
  end;
end;

procedure TFormMain.ButtonOrderClick(Sender: TObject);
var
  Str: string;
begin
  SaveDialog.FileName:= '請申購單-' + FormatDateTime('yymmdd', now()) + '.xlsx';
  if SaveDialog.Execute then
  begin
    Str:= ExtractFileExt(SaveDialog.FileName);
    if Str = '' then
    begin
      Str:='.xlsx';
      SaveDialog.FileName:= SaveDialog.FileName + Str;
    end;
    if str = '.csv' then
      StringGrid1.SaveToCSVFile(SaveDialog.FileName)
    else if Str = '.xlsx' then
    begin
      Str:= ExtractFilePath(SaveDialog.FileName) + DefaultOrderXltx;
      if not FileExists(Str) then
      begin
        Str:= GetCurrentDir + '\' + DefaultOrderXltx;
        if not FileExists(Str) then
        begin
          Str:= ExtractFilePath(Application.ExeName) + DefaultOrderXltx;
          if not FileExists(Str) then
          begin
            ShowMessage('Template File: ' + DefaultOrderXltx + ' Not found!');
            exit;
          end;
        end;
      end;
      StringGridSaveToXlsFileBom(StringGrid1, SaveDialog.FileName, Str, ProgressBar);
      StringGrid1.AutoSizeColumns;
    end;
  end;
end;

procedure TFormMain.ButtonRemove1Click(Sender: TObject);
var
  Str: String;
  Sel, Row: Integer;
begin
  Row:= StringGrid1.Selection.Top;
  Str:= 'StringGrid1 selected row =' + StringGrid1.Selection.Top.ToString + constCrLf;
  Str:= Str + 'Confirm deletion of row from stringgrid1?';
  Sel:= MessageDlg(Str, mtConfirmation, mbOKCancel, 0);
  if Sel = mrOK then
    StringGrid1.DeleteColRow(false, Row);
end;

procedure TFormMain.ButtonRemove2Click(Sender: TObject);
var
  Str: String;
  Sel, Row: Integer;
begin
  Row:= StringGrid2.Selection.Top;
  Str:= 'StringGrid2 selected row =' + StringGrid2.Selection.Top.ToString + constCrLf;
  Str:= Str + 'Confirm deletion of row from stringgrid2?';
  Sel:= MessageDlg(Str, mtConfirmation, mbOKCancel, 0);
  if Sel = mrOK then
    StringGrid2.DeleteColRow(false, Row);
end;

procedure TFormMain.ButtonRenum1Click(Sender: TObject);
var
  Row: Integer;
begin
  for Row:= 2 to StringGrid1.RowCount-1 do
  begin
    StringGrid1.Cells[0, Row]:= IntToStr(Row);
    StringGrid1.Cells[BomColID, Row]:= IntToStr(Row-1);
  end;
  StringGrid1.AutoSizeColumns;
end;

procedure TFormMain.ButtonRenum2Click(Sender: TObject);
var
  Row: Integer;
begin
  for Row:= 1 to StringGrid2.RowCount-1 do
  begin
    StringGrid2.Cells[0, Row]:= IntToStr(Row);
  end;
  StringGrid2.AutoSizeColumns;
end;

procedure TFormMain.ButtonSave1Click(Sender: TObject);
var
  Str: string;
begin
  SaveDialog.FileName:= '材料清單-' + FormatDateTime('yymmdd', now()) + '.xlsx';
  if SaveDialog.Execute then
  begin
    Str:= ExtractFileExt(SaveDialog.FileName);
    if Str = '' then
    begin
      Str:='.xlsx';
      SaveDialog.FileName:= SaveDialog.FileName + Str;
    end;
    if Str = '.csv' then
      StringGrid1.SaveToCSVFile(SaveDialog.FileName)
    else if Str = '.xlsx' then
    begin
      Str:= ExtractFilePath(SaveDialog.FileName) + DefaultBomXltx;
      if not FileExists(Str) then
      begin
        Str:= GetCurrentDir + '\' + DefaultBomXltx;
        if not FileExists(Str) then
        begin
          Str:= ExtractFilePath(Application.ExeName) + DefaultBomXltx;
          if not FileExists(Str) then
          begin
            ShowMessage('Template File: ' + DefaultBomXltx + ' Not found!');
            exit;
          end;
        end;
      end;
      StringGridSaveToXlsFileBom(StringGrid1, SaveDialog.FileName, Str, ProgressBar);
      StringGrid1.AutoSizeColumns;
    end;
  end;
end;

procedure TFormMain.ButtonSave2Click(Sender: TObject);
var
  Str: string;
begin
  SaveDialog.FileName:= '材料總表-' + FormatDateTime('yymmdd', now()) + '.xlsx';
  if SaveDialog.Execute then
  begin
    Str:= ExtractFileExt(SaveDialog.FileName);
    if Str = '' then
    begin
      Str:='.xlsx';
      SaveDialog.FileName:= SaveDialog.FileName + Str;
    end;
    if str = '.csv' then
      StringGrid1.SaveToCSVFile(SaveDialog.FileName)
    else if Str = '.xlsx' then
    begin
      Str:= ExtractFilePath(SaveDialog.FileName) + DefaultStockXltx;
      if not FileExists(Str) then
      begin
        Str:= GetCurrentDir + '\' + DefaultStockXltx;
        if not FileExists(Str) then
        begin
          Str:= ExtractFilePath(Application.ExeName) + DefaultStockXltx;
          if not FileExists(Str) then
          begin
            ShowMessage('Template File: ' + DefaultStockXltx + ' Not found!');
            exit;
          end;
        end;
      end;
      StringGridSaveToXlsFileStk(StringGrid2, SaveDialog.FileName, Str, ProgressBar);
      StringGrid2.AutoSizeColumns;
    end;
  end;
end;

procedure TFormMain.ButtonUpdateClick(Sender: TObject);
var
  IdxBom, IdxStk, IdxNum, Idx: Integer;
  Qty: Integer;
  Sub: double;
  Str: String;
  ListQty, ListStr, ListTmp: TStringList;
begin
  ListQty:= TStringList.Create;
  ListStr:= TStringList.Create;
  ListTmp:= TStringList.Create;
  // Search excel of stock by supplier's part number
  for IdxBom:= 2 to StringGrid1.RowCount-1 do
  begin
    // Skip if supplier's part number have nothing
    Str:= StringGrid1.Cells[BomColSupplierPart, IdxBom];
    if Str = '' then
      continue;
    IdxStk:= StringGridSearchString(StringGrid2, Str);
    if IdxStk > 0 then
    begin
      // Copy SYO part to BOM
      StringGrid1.Cells[BomColSyoPart, IdxBom]:= StringGrid2.Cells[StkColSyoPart, IdxStk];
      // Copy name to BOM
      StringGrid1.Cells[BomColName, IdxBom]:= StringGrid2.Cells[StkColName, IdxStk];
      // Copy manufacturer part to BOM
      StringGrid1.Cells[BomColManufacturerPart, IdxBom]:= StringGrid2.Cells[StkColManufacturerPart, IdxStk];
      // Manufacturer
      StringGrid2.Cells[StkColManufacturer, IdxStk]:= StringGrid1.Cells[BomColManufacturer, IdxBom];
      // Supplier
      StringGrid2.Cells[StkColSupplier, IdxStk]:= StringGrid1.Cells[BomColSupplier, IdxBom];
      // Price
      StringGrid2.Cells[StkColPrice, IdxStk]:= StringGrid1.Cells[BomColPrice, IdxBom];
      // Pins
      StringGrid2.Cells[StkColPins, IdxStk]:= StringGrid1.Cells[BomColPins, IdxBom];
      // Remark
      StringGrid2.Cells[StkColRemark, IdxStk]:= StringGrid1.Cells[BomColRemark, IdxBom];
      // Replacement
      StringGrid2.Cells[StkColReplacement, IdxStk]:= StringGrid1.Cells[BomColReplacement, IdxBom];
      // Dependances
      if not (StringGrid2.Cells[StkColDependances, IdxStk] = '') then
      begin
        ListQty.Add(StringGrid1.Cells[BomColQuantity, IdxBom]);
        ListStr.Add(StringGrid2.Cells[StkColDependances, IdxStk]);
      end;
    end;
  end;
  // Precess Dependances
  if ListStr.Count > 0 then
  begin
    Str:= 'Dependances = ' + IntToStr(ListStr.Count) + ', List = ';
    for IdxBom:= 0 to ListQty.Count-1 do
      Str:= Str + ListQty[IdxBom] + ' ';
    Str:= Str + constCrLf;
    for IdxBom:= 0 to ListStr.Count-1 do
      Str:= Str + ListStr[IdxBom] + constCrLf;
    ShowMessage(Str);
    // Insert all dependances
    IdxBom:= StringGrid1.RowCount;
    for IdxNum:= 0 to ListStr.Count-1 do
    begin
      ListTmp.Clear;
      ListTmp.CommaText:= ListStr.Strings[IdxNum];
      for Idx:= 0 to ListTmp.Count-1 do
      begin
        IdxStk:= StringGridSearchString(StringGrid2, ListTmp.Strings[Idx]);
        if IdxStk > 0 then
        begin
          // Insert new row
          StringGrid1.InsertColRow(false, StringGrid1.RowCount);
          StringGrid1.Cells[0, IdxBom]:= IntToStr(IdxBom);
          // ID
          StringGrid1.Cells[BomColID, IdxBom]:= IntToStr(IdxBom-1);
          // Copy SYO part to BOM
          StringGrid1.Cells[BomColSyoPart, IdxBom]:= StringGrid2.Cells[StkColSyoPart, IdxStk];
          // Copy name to BOM
          StringGrid1.Cells[BomColName, IdxBom]:= StringGrid2.Cells[StkColName, IdxStk];
          // Set quantity
          Qty:= ListQty[IdxNum].ToInteger;
          StringGrid1.Cells[BomColQuantity, IdxBom]:= IntToStr(Qty);
          // Copy manufacturer part to BOM
          StringGrid1.Cells[BomColManufacturerPart, IdxBom]:= StringGrid2.Cells[StkColManufacturerPart, IdxStk];
          // Copy manufacturer to BOM
          StringGrid1.Cells[BomColManufacturer, IdxBom]:= StringGrid2.Cells[StkColManufacturer, IdxStk];
          // Copy supplier to BOM
          StringGrid1.Cells[BomColSupplier, IdxBom]:= StringGrid2.Cells[StkColSupplier, IdxStk];
          // Copy supplier part to BOM
          StringGrid1.Cells[BomColSupplierPart, IdxBom]:= StringGrid2.Cells[StkColSupplierPart, IdxStk];
          // Copy supplier link to BOM
          StringGrid1.Cells[BomColSupplierLink, IdxBom]:= StringGrid2.Cells[StkColSupplierLink, IdxBom];
          // Copy price to BOM
          if StringGrid2.Cells[StkColPrice, IdxStk] = '' then
            StringGrid2.Cells[StkColPrice, IdxStk]:= '0.001';
          StringGrid1.Cells[BomColPrice, IdxBom]:= StringGrid2.Cells[StkColPrice, IdxStk];
          // SubTotal
          // StringGrid1.Cells[BomColSubTotal, IdxBOM]:= '=F'+ IntToStr(IdxBom) + '*L' + IntToStr(IdxBom);
          Sub:= StringGrid1.Cells[BomColPrice, IdxBom].ToDouble;
          StringGrid1.Cells[BomColSubTotal, IdxBOM]:= FloatToStr(Sub * Qty);
          // Copy pins to BOM
          if StringGrid2.Cells[StkColPins, IdxStk] = '' then
            StringGrid2.Cells[StkColPins, IdxStk]:= '0';
          StringGrid1.Cells[BomColPins, IdxBom]:= StringGrid2.Cells[StkColPins, IdxStk];
          // PinTotal
          // StringGrid1.Cells[BomColPinTotal, IdxBom]:= '=F'+ IntToStr(IdxBom) + '*N' + IntToStr(IdxBom);
          StringGrid1.Cells[BomColPinTotal, IdxBom]:= IntToStr(Qty * StringGrid1.Cells[BomColPins, IdxBom].ToInteger);
          // Copy remark to BOM
          StringGrid1.Cells[BomColRemark, IdxBom]:= StringGrid2.Cells[StkColRemark, IdxStk];
          // Copy replacement to BOM
          StringGrid1.Cells[BomColReplacement, IdxBom]:= StringGrid2.Cells[StkColReplacement, IdxStk];
          IdxBom:= IdxBom + 1;
        end;
      end;
    end;
  end;
  StringGrid1.AutoSizeColumns;
  ListTmp.Free;
  ListQty.Free;
  ListStr.Free;
end;

procedure TFormMain.ComboBox1Change(Sender: TObject);
var
  Str: string;
begin
  ComboBox2.Items.Clear;
  ComboBox2.Items.CommaText:= StrArray300[ComboBox1.ItemIndex];
  Str:= Format('113-%.2d', [ComboBox1.ItemIndex]);
  EditFind.Text:= Str;
  StringGridFindString(StringGrid2, Str);
end;

procedure TFormMain.ComboBox2Change(Sender: TObject);
var
  Str: string;
begin
  Str:= Format('113-%.2d', [ComboBox1.ItemIndex]);
  Str:= Str + Format('-%.2d', [ComboBox2.ItemIndex]);
  EditFind.Text:= Str;
  StringGridFindString(StringGrid2, Str);
end;

procedure TFormMain.Edit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    StringGrid1.Cells[StringGrid1.Selection.Left, StringGrid1.Selection.Top]:= Edit1.Text;
    Edit1.SetFocus;
  end;
  if (Key = VK_F2) and (ssAlt in Shift) then
    ShowMessage('Alt F2 was pressed');
end;

procedure TFormMain.Edit2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    StringGrid2.Cells[StringGrid2.Selection.Left, StringGrid2.Selection.Top]:= Edit2.Text;
    Edit2.SetFocus;
  end;
  if (Key in [VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN]) or (Key = VK_SPACE) and (ssCtrl in Shift) then
    ShowMessage('key was pressed');
end;

procedure TFormMain.EditFindKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    StringGridFindString(StringGrid2, EditFind.Text);
  end;
  if (Key in [VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN]) or (Key = VK_SPACE) and (ssCtrl in Shift) then
    ShowMessage('key was pressed');
end;

procedure TFormMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction:= TCloseAction.caFree;
  IniFileWrite(FormMain);
end;

procedure TFormMain.FormCreate(Sender: TObject);
var
  Idx: Integer;
begin
  SetHintFontStyle();
  IniFileRead(FormMain);
  ComboBox1.Items.Clear;
  ComboBox2.Items.Clear;
  for Idx:= 0 to Length(StrArray113)-1 do
    ComboBox1.Items.Add(StrArray113[Idx]);
  {
  for Idx:= 0 to Length(StrArray000)-1 do
    ComboBox1.Items.Add(StrArray000[Idx]);
  ComboBox1.ItemIndex:= 2;
  for Idx:= 0 to Length(StrArray113)-1 do
    ComboBox2.Items.Add(StrArray113[Idx]);
  ComboBox2.ItemIndex:= 20;
  for Idx:= 0 to Length(StrArray300)-1 do
    ComboBox3.Items.Add(StrArray300[Idx]);
  ComboBox3.ItemIndex:= 2;
  }
end;

procedure TFormMain.FormResize(Sender: TObject);
begin
  Panel1.Height:= (Panel1.Height + Panel2.Height) div 2;
  ButtonExit.Left:= Panel3.Width - 120;
  ProgressBar.Left:= Panel3.Width - 256;
end;

procedure TFormMain.StringGrid1Click(Sender: TObject);
begin
  Edit1.Text:= StringGrid1.Cells[StringGrid1.Selection.Left, StringGrid1.Selection.Top];
end;

procedure TFormMain.StringGrid2Click(Sender: TObject);
begin
  Edit2.Text:= StringGrid2.Cells[StringGrid2.Selection.Left, StringGrid2.Selection.Top];
end;

end.

