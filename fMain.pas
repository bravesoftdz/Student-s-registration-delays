unit fMain;

// ���������� ������������ ����� ��� ������ � ����������� SQLite
{$I DI.inc}
{$I DISQLite3.inc}

interface

uses
  DISystemCompat, Classes, Controls, Forms, StdCtrls, ExtCtrls, Grids, Variants,
  DISQLite3Database, Menus, ComCtrls, DateUtils, comobj, OleServer, ExcelXP;

type
  TfrmMain = class(TForm)
    cbGroups: TComboBox;
    cbStudy: TComboBox;
    MainMenu1: TMainMenu;
    mFile: TMenuItem;
    mEdit: TMenuItem;
    mAddGroup: TMenuItem;
    mAddStudy: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    Label1: TLabel;
    Label2: TLabel;
    mEditStudy: TMenuItem;
    mDelStudy: TMenuItem;
    mEditGroup: TMenuItem;
    mDelGroup: TMenuItem;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    dtpAddDelay: TDateTimePicker;
    edtHoursU: TEdit;
    edtHoursN: TEdit;
    btnAddDelay: TButton;
    dtpFrom: TDateTimePicker;
    dtpTo: TDateTimePicker;
    btnSearchDelays: TButton;
    Label6: TLabel;
    Label7: TLabel;
    TabSheet3: TTabSheet;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    lblPrevU: TLabel;
    lblPrevN: TLabel;
    lblStatN: TLabel;
    lblStatU: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Panel1: TPanel;
    TabSheet4: TTabSheet;
    mExit: TMenuItem;
    btnSelectGroup: TButton;
    btnSelectStudy: TButton;
    StatusBar1: TStatusBar;
    Panel2: TPanel;
    cbOnGroup: TCheckBox;
    cbOnStudy: TCheckBox;
    dtpSearchFrom: TDateTimePicker;
    Label11: TLabel;
    Label12: TLabel;
    dtpSearchTo: TDateTimePicker;
    btnFSearch: TButton;
    btnPrint: TButton;
    StringGrid: TStringGrid;
    Label16: TLabel;
    procedure btnFSearchClick(Sender: TObject);
    procedure btnSelectStudyClick(Sender: TObject);
    procedure btnSelectGroupClick(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure StringGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StringGridDblClick(Sender: TObject);
    procedure TabSheet3Show(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure cbStudyKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure cbGroupsKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormDestroy(Sender: TObject);
    procedure mExitClick(Sender: TObject);
    procedure mDelGroupClick(Sender: TObject);
    procedure mEditGroupClick(Sender: TObject);
    procedure cbStudyChange(Sender: TObject);
    procedure btnSearchDelaysClick(Sender: TObject);
    procedure btnAddDelayClick(Sender: TObject);
    procedure mDelStudyClick(Sender: TObject);
    procedure mEditStudyClick(Sender: TObject);
    procedure mAddStudyClick(Sender: TObject);
    procedure mAddGroupClick(Sender: TObject);
    procedure cbGroupsChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DelaysSelect;
    procedure DelaysDel;
    procedure LoadGroups;
    procedure CalculateStats;
    procedure btnPrintClick(Sender: TObject);
  private
    FDatabase: TDISQLite3Database;
    function CheckCanEdit(const ACol, ARow: Integer): Boolean;
    procedure SetGridRow(
      const ARowIdx: Integer;
      const AID: Integer = -1;
      const AName: UnicodeString= '';
      const AGruppa: UnicodeString = '';
      const ADOD: UnicodeString = '';
      const AHoursU: UnicodeString = '';
      const AHoursN: UnicodeString = '');
  end;

var
  frmMain: TfrmMain;

implementation

uses
  SysUtils, Dialogs, DISQLite3Api, fAbout;

{$R *.dfm}

const
  ColumnNames: array[0..4] of UnicodeString = ('���', '������', '����', '�� ������������', '�� ��������������');
  ColumnFields: array[0..4] of UnicodeString = ('fio', 'gruppa', 'dod', 'hours_u', 'hours_n');

// ������� ���������� StringGrid
procedure TfrmMain.SetGridRow(
  const ARowIdx: Integer;
  const AID: Integer = -1;
  const AName: UnicodeString = '';
  const AGruppa: UnicodeString = '';
  const ADOD: UnicodeString = '';
  const AHoursU: UnicodeString = '';
  const AHoursN: UnicodeString = '');
var
  Row: TStrings;
begin
  Row := StringGrid.Rows[ARowIdx];
  Row[0] := AName; Row.Objects[0] := TObject(AID);
  Row[1] := AGruppa; Row.Objects[1] := TObject(AID);
  Row[2] := ADOD; Row.Objects[2] := TObject(AID);
  Row[3] := AHoursU; Row.Objects[3] := TObject(AID);
  Row[4] := AHoursN; Row.Objects[4] := TObject(AID);
end;

procedure TfrmMain.StringGridDblClick(Sender: TObject);
const
  { ������ ��� ���������� �������� � ���������� ������� � ��������� ������ }
  UpdateSql = 'UPDATE delays SET "%s"=? WHERE ID=?;';
var
  c, r: Integer;
  ID: Integer;
  s: string;
  SQL: UnicodeString;
  Stmt: TDISQLite3Statement;
begin
  c := StringGrid.Col; r := StringGrid.Row;
  // ��������� ����������� �������������� ������ ������
  if not CheckCanEdit(c, r) then
    Exit;
  if c < 3 then
    Exit;
  s := StringGrid.Cells[c, r];
  // ��� ������� �� � ������� ������� ����������� ���������� ������ � ����
  if InputQuery('Update', Format('����� �������� %s:', [ColumnNames[c]]), s) then
    begin
      { �������������� ������, ����������� ������� ������ ������� }
      SQL := Format(UpdateSql, [ColumnFields[c]]);
      Stmt := FDatabase.Prepare16(SQL);
      try
      // ������������� ��������� ��� �������
        Stmt.Bind_Str16(1, s);
        // �������� ID ���������� ������ ������������� � ����� �������
        ID := Integer(StringGrid.Objects[c, r]);
        Stmt.Bind_Int(2, ID);
        // ��������� ���� ������ � ��������� ������ � ����
        Stmt.Step;
        // ������� �������� � StringGrid
        StringGrid.Cells[c, r] := s;
      finally
        Stmt.Free;
      end;
    end;
end;

procedure TfrmMain.StringGridKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
{ ������� �� ������� ������� �� ����������
  46 - ��� ������� Del
  116 - ��� ������� F5 }
  case Key of
    46: DelaysDel; // ������� ���������� ������
    116: DelaysSelect; // ��������� ������ ��������� (���������� ��� ������)
  end;
end;

procedure TfrmMain.TabSheet3Show(Sender: TObject);
begin
// ��� �������� �� ������ ������� (����������) ����������� ������� 
  CalculateStats;
end;

procedure TfrmMain.LoadGroups;
const
  SelectSQL = 'SELECT id, gruppa FROM groups ORDER BY gruppa COLLATE NoCase;';
var
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
// �������� ������ ����� �� ���� � ���������� ��� ���������� ComboBox
  StringList := cbGroups.Items;
  StringList.BeginUpdate;
  try
    cbGroups.Clear;
    Stmt := FDatabase.Prepare16(SelectSQL);
    try
      while Stmt.Step = SQLITE_ROW do
        begin
          StringList.AddObject(
            Stmt.Column_Str16(1), // ���
            TObject(Stmt.Column_Int(0))); // ID, �������� ��� TObject
        end;
    finally
      Stmt.Free;
    end;
  finally
    StringList.EndUpdate;
  end;
end;

procedure TfrmMain.CalculateStats;
const
  SelectSQL = 'SELECT id, fio, gruppa, hours_u, hours_n, dod FROM delays WHERE (fio=? and gruppa=?) AND dod BETWEEN ? and ?;';
var
  Stmt, Stmt2: TDISQLite3Statement;
  cHoursU, cHoursN,  cPrevHoursN, cPrevHoursU: Integer;
  PrevMonth, CurrMonth: TDate;
begin
// ������ ���������� ��������
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    cHoursU := 0; // ������� ������������ ��������
    cHoursN := 0; // ������� �������������� ��������
    CurrMonth := Date; // �������� ������� ����

    // ��������� ��������� ��� �������
    Stmt.Bind_Str16(1, cbStudy.Text); // ��������� �������
    Stmt.Bind_Str16(2, cbGroups.Text); // ��������� ������
    // �� ������ ������� ���� ���������� ������ � ��������� ���� ������
    // � ����������� ���������� ���� � ��������� ������
    Stmt.Bind_Double(3, DateToJulianDate(StartOfTheMonth(CurrMonth)));
    Stmt.Bind_Double(4, DateToJulianDate(EndOfTheMonth(CurrMonth)));
    while Stmt.Step = SQLITE_ROW do
      begin
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(3)); //��������� �������
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(4));
      end;
    lblStatU.Caption := IntToStr(cHoursU);
    lblStatN.Caption := IntToStr(cHoursN);
  finally
    Stmt.Free;
  end;
  // ���� ����������� �� �� ���������, �� ��� ����������� ������
  Stmt2 := FDatabase.Prepare16(SelectSQL);
  try
    cPrevHoursU := 0; // ������� ������������ �������� �� ������� �����
    cPrevHoursN := 0; // ������� �������������� �������� �� ������� �����
    PrevMonth := IncMonth(Date, -1);

    Stmt2.Bind_Str16(1, cbStudy.Text);
    Stmt2.Bind_Str16(2, cbGroups.Text);
    Stmt2.Bind_Double(3, DateToJulianDate(StartOfTheMonth(PrevMonth)));
    Stmt2.Bind_Double(4, DateToJulianDate(EndOfTheMonth(PrevMonth)));
    while Stmt2.Step = SQLITE_ROW do
      begin
        cPrevHoursU := cPrevHoursU + StrToInt(Stmt2.Column_Str16(3));
        cPrevHoursN := cPrevHoursN + StrToInt(Stmt2.Column_Str16(4));
      end;
    lblPrevU.Caption := IntToStr(cPrevHoursU);
    lblPrevN.Caption := IntToStr(cPrevHoursN);
  finally
    Stmt2.Free;
  end;
end;

// ��������� �������� ���������� ������
procedure TfrmMain.mDelGroupClick(Sender: TObject);
const
  DeleteSQL = 'DELETE FROM groups WHERE ID=?;';
var
  i, Idx: Integer;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  { ��������� StringList. }
  StringList := cbGroups.Items;

  i := cbGroups.ItemIndex;
  if i < 0 then
    begin
      ShowMessage('�������� ������� �������.');
      Exit;
    end;

  if MessageDlg('�������?', mtConfirmation, mbOKCancel, 0) = mrOk then
    begin
      Idx := Integer(StringList.Objects[i]);
      { �������������� ������ �� ��������. }
      Stmt := FDatabase.Prepare16(DeleteSQL);
      try
        Stmt.Bind_Int(1, Idx);
        { ���� ���������� ������� � �������� ������ �� ���� }
        Stmt.Step;
        { ������� ������� �� StringList. }
        StringList.Delete(i);
      finally
        Stmt.Free;
      end;
    end;
  // ����������� ������
  LoadGroups;
end;

procedure TfrmMain.mExitClick(Sender: TObject);
begin
  // ����� ���� �����. ��������� ����������
  Application.Terminate;
end;

procedure TfrmMain.N6Click(Sender: TObject);
begin
  // ���������� ��������� ���� � �������
  frmAbout.ShowModal;
end;

// ��������� ���������� ������
procedure TfrmMain.mAddGroupClick(Sender: TObject);
const
  InsertSQL = 'INSERT INTO groups (gruppa) VALUES (?);';
var
  NewGroupName: String;
  Stmt: TDISQLite3Statement;
begin
  NewGroupName := ''; // ������� ����������
  // ������� ������ ������� ��� ����� ������
  if InputQuery('��������', '������� ����� ������:', NewGroupName) then
    begin
      { �������������� ������ }
      Stmt := FDatabase.Prepare16(InsertSQL);
      try
        // ��������� ���������
        Stmt.Bind_Str16(1, NewGroupName);
        { ������ ���� ������ � ��������� ������ � ���� }
        Stmt.Step;
      finally
        Stmt.Free;
      end;
    end;
    // ����������� ������
    LoadGroups;
end;

// ��������� ���������� ��������
procedure TfrmMain.mAddStudyClick(Sender: TObject);
const
  InsertSQL = 'INSERT INTO study (fio, gruppa) VALUES (?, ?);';
var
  Stmt: TDISQLite3Statement;
  s: String;
begin
  s := ''; // ������� ����������
  // ������� ������ �� ���� ������ ��� ����������
  if InputQuery('��������', '������� ��� ��������:', s) then
    begin
      { �������������� ������ }
      Stmt := FDatabase.Prepare16(InsertSQL);
      try
        // ��������� ���������
        Stmt.Bind_Str16(1, s);
        Stmt.Bind_Str16(2, cbGroups.Text);
        { ������ ���� ������ � ��������� ������ }
        Stmt.Step;
      finally
        Stmt.Free;
      end;
    end;
    cbGroupsChange(cbGroups);
end;

// ��������� �������������� ��������
procedure TfrmMain.mEditStudyClick(Sender: TObject);
const
  UpdateSql = 'UPDATE study SET fio=? WHERE id=?';
var
  i, Idx: Integer;
  s: string;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  // ��������� StringList ���������� �� ComboBox �� ������� ���������
  StringList := cbStudy.Items;

  i := cbStudy.ItemIndex; // ������ �������� � ������
  s := StringList[i]; // ����������� ���������� �������� �������

  // ������ ������ ������� �� �������������� ������
  if InputQuery('��������', '������� ����� ���:', s) then
    begin
      // ����������� ���������� ������ ����������
      Idx := Integer(StringList.Objects[i]);
      // �������������� ������
      Stmt := FDatabase.Prepare16(UpdateSql);
      try
        // ��������� ��������� �������
        Stmt.Bind_Str16(1, s);
        Stmt.Bind_Int(2, Idx);
        // ��������� ���� ������ � ��������� ������ � �������
        Stmt.Step;
        // ������� ����� �������� � StringList
        StringList[i] := s;
      finally
        Stmt.Free;
      end;
    end;
end;

// ��������� �������� ��������
procedure TfrmMain.mDelStudyClick(Sender: TObject);
const
  DeleteSQL = 'DELETE FROM study WHERE ID=?;';
var
  i, Idx: Integer;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  // ��������� StringList
  StringList := cbStudy.Items;
  // ����������� ���������� ������ ��������
  i := cbStudy.ItemIndex;
  if i < 0 then
    begin
      ShowMessage('�������� ������� �������.');
      Exit;
    end;
  // ������� ������ ������� �� ��������
  if MessageDlg('�������?', mtConfirmation, mbOKCancel, 0) = mrOk then
    begin
      Idx := Integer(StringList.Objects[i]);
      // �������������� ������
      Stmt := FDatabase.Prepare16(DeleteSQL);
      try
        Stmt.Bind_Int(1, Idx);
        // ��������� ���� ������ � ������� ������ �� ����
        Stmt.Step;
        // ������� ����� �������� � StringList
        StringList.Delete(i);
      finally
        Stmt.Free;
      end;
    end;
end;

// ��������� �������������� ������
procedure TfrmMain.mEditGroupClick(Sender: TObject);
const
  UpdateSql = 'UPDATE groups SET gruppa=? WHERE id=?';
var
  i, Idx: Integer;
  s: string;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  // ��������� StringList
  StringList := cbGroups.Items;

  i := cbGroups.ItemIndex;
  s := StringList[i];
  // ������� ������ ��� ��������� ��������
  if InputQuery('��������', '������� ����� ���:', s) then
    begin
      Idx := Integer(StringList.Objects[i]);
      // �������������� ������
      Stmt := FDatabase.Prepare16(UpdateSql);
      try
        Stmt.Bind_Str16(1, s);
        Stmt.Bind_Int(2, Idx);
        // ��������� ���� ������ � ��������� ������ � ����
        Stmt.Step;
        // ������� ����� �������� � StringList
        StringList[i] := s;
      finally
        Stmt.Free;
      end;
    end;
  // ����������� ������
  LoadGroups;
end;

// ������� �������� ����������� ��������������
function TfrmMain.CheckCanEdit(const ACol, ARow: Integer): Boolean;
begin
  Result := (ACol >= 0) and (ARow >= 1) and (Integer(StringGrid.Rows[ARow].Objects[ACol]) > 0);
  if not Result then
    begin
      ShowMessage('������ �������������.');
    end;
end;

// ��������� ��������� ��������� ������
procedure TfrmMain.cbGroupsChange(Sender: TObject);
const
  GetGroupsSql = 'SELECT id, gruppa FROM groups ORDER BY gruppa COLLATE NoCase;';
  GetStudysSql = 'SELECT id, fio FROM study WHERE gruppa=?';
var
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  cbStudy.Clear; // ������� ComboBox �� ����������
  // ����������� StringList � ���������� Combobox ���������
  StringList := cbStudy.Items;
  // �������� ���������� ������
  StringList.BeginUpdate;
  try
    // �������������� ������
    Stmt := FDatabase.Prepare16(GetStudysSql);
    // ������������� ��������� ��� �������
    Stmt.Bind_Str16(1, cbGroups.Text);
    try
      while Stmt.Step = SQLITE_ROW do
        begin
          // ������ �������� � StringList ��������� ��� � ID �������
          StringList.AddObject(
            Stmt.Column_Str16(1), // ���
            TObject(Stmt.Column_Int(0))); // ID, �������� ��� TObject.
        end;
    finally
      Stmt.Free;
    end;
  finally
    // ����������� ���������� ������
    StringList.EndUpdate;
  end;
  btnSelectGroupClick(Self);
end;

// ��������� ��������� ������� ������ � ������ �����
procedure TfrmMain.cbGroupsKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    45: mAddGroupClick(self); // Insert - ���������� ������
    46: mDelGroupClick(self); // Del - �������� ������
  end;
end;

// ��� ������ �������� ������������ ��� ���� ����������
procedure TfrmMain.cbStudyChange(Sender: TObject);
begin
  CalculateStats;
  btnSelectStudyClick(Self);
end;

// ��������� ��������� ������� ������ � ������ ���������
procedure TfrmMain.cbStudyKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    45: mAddStudyClick(self); // Insert - ���������� ��������
    46: mDelStudyClick(self); // Del - �������� ��������
  end;
end;

// ��� ��������� ����� ������������� ����� �� ������ �����
procedure TfrmMain.FormActivate(Sender: TObject);
begin
  cbGroups.SetFocus;
end;

// ��������� ����������� ��� �������� �����
procedure TfrmMain.FormCreate(Sender: TObject);
begin
  // ������ ����� ������ �� � ��������� ��� � ������
  FDatabase := TDISQLite3Database.Create(nil);
  FDatabase.DatabaseName := 'Study.db';

  // ������������� ��������� ��� StringGrid
  StringGrid.ColCount := 5; // 5 �������
  StringGrid.DefaultColWidth := 120; // ������ ������� �� ��������� 120
  StringGrid.DefaultRowHeight := 19; // ������ ���� 19
  // ��������� ����������� �������� ������ �������
  StringGrid.Options := StringGrid.Options + [goColSizing];
  StringGrid.FixedCols := 0;
  StringGrid.RowCount := 2;
  // ��������� ����� �������
  SetGridRow(0, -1, ColumnNames[0], ColumnNames[1], ColumnNames[2], ColumnNames[3], ColumnNames[4]);
  SetGridRow(1);
  // ��������� ������
  LoadGroups;
  // ������������� ������� ���� ��� ���� ����������� ����
  dtpAddDelay.Date := Date;
  dtpFrom.Date := Date;
  dtpTo.Date := Date;
  dtpSearchFrom.Date := Date;
  dtpSearchTo.Date := Date;
  // ������������� ������ ������� ��������
  PageControl1.ActivePageIndex := 0;
  // ��������� �������
  DelaysSelect;
end;

// ��������� ����������� ����� ����������� �����
procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  // ��������� ����� � ����� ������
  FDatabase.Close;
  // ����������� ������ ����� �������� �����
  FDatabase.Free;
  // ��������� ����������
  Application.Terminate;
end;

// ��� ������ ����� ������������� ����� �� ������ �����
procedure TfrmMain.FormShow(Sender: TObject);
begin
  cbGroups.SetFocus;
end;

// �������� ��������
procedure TfrmMain.DelaysSelect;
const
  SelectSQL = 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays ORDER BY gruppa COLLATE NoCase;';
var
  ID, RowCount: Integer;
  Stmt: TDISQLite3Statement;
begin
  // �������������� ������
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    RowCount := 1;
    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        // ��������� StringGrid � �����
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
        Inc(RowCount);
      end;
    StringGrid.RowCount := RowCount;
  finally
    Stmt.Free;
  end;
end;

// ��������� ���������� �������
procedure TfrmMain.btnAddDelayClick(Sender: TObject);
const
  InsertSQL = 'INSERT INTO delays (fio, gruppa, dod, hours_u, hours_n) VALUES (?, ?, ?, ?, ?);';
var
  Stmt: TDISQLite3Statement;
begin
  if ((cbGroups.Text <> '') or (cbStudy.Text <> '')) then
    begin
      // �������������� ������
      Stmt := FDatabase.Prepare16(InsertSQL);
      try
        // ������������� ��������� �������
        Stmt.Bind_Str16(1, cbStudy.Text);
        Stmt.Bind_Str16(2, cbGroups.Text);
        Stmt.Bind_Double(3, DateToJulianDate(dtpAddDelay.Date));
        Stmt.Bind_Int(4, StrToInt(edtHoursU.Text));
        Stmt.Bind_Int(5, StrToInt(edtHoursN.Text));
        // ��������� ������ � ������� ������ � ����
        Stmt.Step;
      finally
        Stmt.Free;
      end;
    end;
  edtHoursU.Text := '';
  edtHoursN.Text := '';
  btnSelectStudyClick(Self);
end;

procedure TfrmMain.btnSearchDelaysClick(Sender: TObject);
const
  SelectSQL = 'SELECT * FROM delays WHERE dod BETWEEN ? and ?;';
var
  ID, RowCount, cHoursU, cHoursN: Integer;
  Stmt: TDISQLite3Statement;
begin
  // �������������� ������
  cHoursU := 0;
  cHoursN := 0;
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    RowCount := 1;
    Stmt.Bind_Double(1, DateToJulianDate(dtpFrom.Date));
    Stmt.Bind_Double(2, DateToJulianDate(dtpTo.Date));
    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
        Inc(RowCount);
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(4));
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(5));
      end;
    StringGrid.RowCount := RowCount;
    StatusBar1.Panels[0].Text := '����� ����� ����������� ���������: �� ������������ - ' + IntToStr(cHoursU) + '; �� �������������� - ' + IntToStr(cHoursN) ;
  finally
    Stmt.Free;
  end;
end;

procedure TfrmMain.btnSelectGroupClick(Sender: TObject);
const
  SelectSQL = 'SELECT * FROM delays WHERE gruppa=?;';
var
  ID, RowCount, cHoursU, cHoursN: Integer;
  Stmt: TDISQLite3Statement;
begin
  cHoursU := 0;
  cHoursN := 0;
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    RowCount := 1;
    Stmt.Bind_Str16(1, cbGroups.Text);
    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
        Inc(RowCount);
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(4));
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(5));
      end;
    StringGrid.RowCount := RowCount;
    StatusBar1.Panels[0].Text := '����� ����� ����������� ���������: �� ������������ - ' + IntToStr(cHoursU) + '; �� �������������� - ' + IntToStr(cHoursN) ;
  finally
    Stmt.Free;
  end;
end;

procedure TfrmMain.btnSelectStudyClick(Sender: TObject);
const
  SelectSQL = 'SELECT * FROM delays WHERE fio=?;';
var
  ID, RowCount, cHoursU, cHoursN: Integer;
  Stmt: TDISQLite3Statement;
begin
  cHoursU := 0;
  cHoursN := 0;
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    RowCount := 1;
    Stmt.Bind_Str16(1, cbStudy.Text);
    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
        Inc(RowCount);
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(4));
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(5));
      end;
    StringGrid.RowCount := RowCount;
    StatusBar1.Panels[0].Text := '����� ����� ����������� ���������: �� ������������ - ' + IntToStr(cHoursU) + '; �� �������������� - ' + IntToStr(cHoursN) ;
  finally
    Stmt.Free;
  end;
end;

procedure TfrmMain.btnFSearchClick(Sender: TObject);
var
  Stmt: TDISQLite3Statement;
  ID, RowCount, cnt, cHoursU, cHoursN: Integer;
  SelectSQL: String;
begin
  cHoursU := 0;
  cHoursN := 0;
  if (cbOnGroup.Checked and cbOnStudy.Checked) then cnt := 3
    else if (cbOnGroup.Checked and not cbOnStudy.Checked) then cnt := 1
      else if (cbOnStudy.Checked and not cbOnGroup.Checked) then cnt := 2
        else cnt := 0;
  case cnt of
    0: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE dod BETWEEN ? and ?;';
    1: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE (dod BETWEEN ? and ?) AND gruppa=?;';
    2: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE (dod BETWEEN ? and ?) AND fio=?;';
    3: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE (dod BETWEEN ? and ?) AND (fio=? and gruppa=?);';
  end;
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    RowCount := 1;
    Stmt.Bind_Double(1, DateToJulianDate(dtpSearchFrom.Date));
    Stmt.Bind_Double(2, DateToJulianDate(dtpSearchTo.Date));
    case cnt of
      0: ;
      1: Stmt.Bind_Str16(3, cbGroups.Text);
      2: Stmt.Bind_Str16(3, cbStudy.Text);
      3: begin
          Stmt.Bind_Str16(3, cbStudy.Text);
          Stmt.Bind_Str16(4, cbGroups.Text);
         end;
    end;
    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
        Inc(RowCount);
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(4));
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(5));
      end;
    StringGrid.RowCount := RowCount;
    StatusBar1.Panels[0].Text := '����� ����� ����������� ���������: �� ������������ - ' + IntToStr(cHoursU) + '; �� �������������� - ' + IntToStr(cHoursN) ;
  finally
    Stmt.Free;
  end;
end;

procedure TfrmMain.DelaysDel;
const
  DeleteSQL = 'DELETE FROM delays WHERE ID=?;';
var
  ID, r: Integer;
  Stmt: TDISQLite3Statement;
begin
  r := StringGrid.Row;
  if not CheckCanEdit(0, r) then
    Exit;

  if MessageDlg('�������?', mtConfirmation, mbOKCancel, 0) = mrOk then
    begin
      Stmt := FDatabase.Prepare16(DeleteSQL);
      try
        ID := Integer(StringGrid.Rows[r].Objects[0]);
        Stmt.Bind_Int(1, ID);
        Stmt.Step;
        DelaysSelect;
      finally
        Stmt.Free;
      end;
    end;
end;

// ��������� ������ ������ � Excel
procedure TfrmMain.btnPrintClick(Sender: TObject);
var
  Stmt: TDISQLite3Statement;
  ID, RowCount, cnt, BeginRow, BeginCol, cHoursU, cHoursN: Integer;
  SelectSQL: String;
  ExcelApp, Cell1, Cell2, ColumnRange: Variant;
begin
  RowCount := 1;
  // ���������� ������ �������� ���� �������, � ������� ����� �������� ������
  BeginCol := 1;
  BeginRow := 5;

  // ��������� �������� � � ����������� �� �� ��������� ���������� ������ ������
  if (cbOnGroup.Checked and cbOnStudy.Checked) then cnt := 3
    else if (cbOnGroup.Checked and not cbOnStudy.Checked) then cnt := 1
      else if (cbOnStudy.Checked and not cbOnGroup.Checked) then cnt := 2
        else cnt := 0;
  case cnt of
    0: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE dod BETWEEN ? and ?;';
    1: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE (dod BETWEEN ? and ?) AND gruppa=?;';
    2: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE (dod BETWEEN ? and ?) AND fio=?;';
    3: SelectSQL := 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays WHERE (dod BETWEEN ? and ?) AND (fio=? and gruppa=?);';
  end;
  // �������������� ������
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    // �������������� ��������� Excel
    ExcelApp :=CreateOleObject('Excel.Application');
    // ��������� ������������ ������� ���� ��������� ������
    ExcelApp.Application.EnableEvents := false;
    // ��������� �����
    ExcelApp.WorkBooks.Add;
    // ������������� ������� ��������
    ColumnRange := ExcelApp.Workbooks[1].WorkSheets[1].Columns;
    // ������������� ������ �������� � ������ ��� ��������
    ColumnRange.Columns[1].ColumnWidth := 20;
    ColumnRange.Columns[2].ColumnWidth := 10;
    ColumnRange.Columns[3].ColumnWidth := 15;
    ColumnRange.Columns[4].ColumnWidth := 15;
    ColumnRange.Columns[5].ColumnWidth := 15;

    // ������� ����� � ������ ��� ������ ������
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 2] := '����� �� ���������';
    // ������ �������� ��� ���������
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[1, 2];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[1, 4];
    // �������� �������� ��������� �����
    ExcelApp.ActiveWorkBook.WorkSheets[1].Range[Cell1, Cell2].Select;
    // ���������� ���������� ������
    ExcelApp.Selection.MergeCells := True;
    // ������������� ������ ����� ���������� �����
    ExcelApp.Selection.Font.Bold:=True;
    // ������������� ������������ �� ������
    ExcelApp.Selection.HorizontalAlignment:=3;

    // ������ �������� ��� ���������
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 1];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 5];
    // �������� �������� ��������� �����
    ExcelApp.ActiveWorkBook.WorkSheets[1].Range[Cell1, Cell2].Select;
    // ��������� ������ ����������
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 1] := '�� ������';
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 2] := 'c:';
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 3] := FormatDateTime('dd.mmm.yyyy', dtpSearchFrom.Date);
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 4] := '��:';
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 5] := FormatDateTime('dd.mmm.yyyy', dtpSearchTo.Date);
    // ������������� ������������ �� ������� ����
    ExcelApp.Selection.HorizontalAlignment := 4;

    // ��������� ����� ��� ������
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 1] := '���';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 2] := '������';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 3] := '����';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 4] := '�����������';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 5] := '������������';
    // ������ �������� ��� ���������
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, BeginCol];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, BeginCol+4];
    // �������� �������� ��������� �����
    ExcelApp.WorkBooks[1].WorkSheets[1].Range[Cell1, Cell2].Select;
    // ������������� ������ ����� ���������� �����
    ExcelApp.Selection.Font.Bold:=True;
    // ������������� ������������ �� ������
    ExcelApp.Selection.HorizontalAlignment:=3;

    // ��������� ��������� ��� �������
    Stmt.Bind_Double(1, DateToJulianDate(dtpSearchFrom.Date));
    Stmt.Bind_Double(2, DateToJulianDate(dtpSearchTo.Date));
    // � ����������� �� ��������� ��������� ��������� �������������� ����
    case cnt of
      0: ;
      1: Stmt.Bind_Str16(3, cbGroups.Text);
      2: Stmt.Bind_Str16(3, cbStudy.Text);
      3: begin
          Stmt.Bind_Str16(3, cbStudy.Text);
          Stmt.Bind_Str16(4, cbGroups.Text);
         end;
    end;

    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        // ��������� StringGrid
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
          // ������� ���������� ������ � Excel
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 1] := Stmt.Column_Str16(1);
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 2] := Stmt.Column_Str16(2);
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 3] := DateToStr(JulianDateToDate(Stmt.Column_Double(3)));
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 4] := Stmt.Column_Str16(4);
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 5] := Stmt.Column_Str16(5);
        Inc(RowCount);
        // ����������� ����� ��������
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(4));
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(5));
      end;
    StringGrid.RowCount := RowCount;
    // ������� ���������� �� ������
    StatusBar1.Panels[0].Text := '����� ����� ����������� ���������: �� ������������ - ' + IntToStr(cHoursU) + '; �� �������������� - ' + IntToStr(cHoursN) ;

    // ������� ������ � ������ � ��������� ��������������
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 1] := '�����:';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 4] := cHoursU;
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 5] := cHoursN;
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, BeginCol];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, BeginCol+4];
    ExcelApp.WorkBooks[1].WorkSheets[1].Range[Cell1, Cell2].Select;
    ExcelApp.Selection.Font.Bold:=True;

    // ���������� Excel
    ExcelApp.Visible := true;
    // ������� ����� � ����������� ����� ��� �� ������ � ������
    ExcelApp:=Unassigned;

  finally
    Stmt.Free;
  end;
end;

initialization
  // �������������� DISQLite3 ���������� DISQLite3
  sqlite3_initialize;

finalization
  // ����������� ������ � ����������� �� ����������
  sqlite3_shutdown;

end.

