unit fMain;

// Подключаем заголовочные файлы для работы с библиотекой SQLite
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
  ColumnNames: array[0..4] of UnicodeString = ('ФИО', 'Группа', 'Дата', 'По уважительной', 'По неуважительной');
  ColumnFields: array[0..4] of UnicodeString = ('fio', 'gruppa', 'dod', 'hours_u', 'hours_n');

// Функция заполнения StringGrid
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
  { запрос для обновления пропуска в выделенном столбце и выделеной строке }
  UpdateSql = 'UPDATE delays SET "%s"=? WHERE ID=?;';
var
  c, r: Integer;
  ID: Integer;
  s: string;
  SQL: UnicodeString;
  Stmt: TDISQLite3Statement;
begin
  c := StringGrid.Col; r := StringGrid.Row;
  // Проверяем возможность редактирования данной ячейки
  if not CheckCanEdit(c, r) then
    Exit;
  if c < 3 then
    Exit;
  s := StringGrid.Cells[c, r];
  // При нажатии ОК в диалоге запроса выполняется обновление данных в базе
  if InputQuery('Update', Format('Новое значение %s:', [ColumnNames[c]]), s) then
    begin
      { подготавливаем запрос, динамически выбирая нужный столбец }
      SQL := Format(UpdateSql, [ColumnFields[c]]);
      Stmt := FDatabase.Prepare16(SQL);
      try
      // устанавливаем параметры для запроса
        Stmt.Bind_Str16(1, s);
        // Получаем ID выделенной ячейки идентифицируя её таким образом
        ID := Integer(StringGrid.Objects[c, r]);
        Stmt.Bind_Int(2, ID);
        // Выполняем один проход и обновляем запись в базе
        Stmt.Step;
        // Заносим изменене в StringGrid
        StringGrid.Cells[c, r] := s;
      finally
        Stmt.Free;
      end;
    end;
end;

procedure TfrmMain.StringGridKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
{ событие на нажатие клавиши на клавиатуре
  46 - код клавиши Del
  116 - код клавиши F5 }
  case Key of
    46: DelaysDel; // Удаляем выделенную запись
    116: DelaysSelect; // Обновляем список пропусков (выбираются все данные)
  end;
end;

procedure TfrmMain.TabSheet3Show(Sender: TObject);
begin
// При переходе на третью вкладку (Статистика) расчитываем прогулы 
  CalculateStats;
end;

procedure TfrmMain.LoadGroups;
const
  SelectSQL = 'SELECT id, gruppa FROM groups ORDER BY gruppa COLLATE NoCase;';
var
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
// Загрузка списка групп из базы и заполнение ими компонента ComboBox
  StringList := cbGroups.Items;
  StringList.BeginUpdate;
  try
    cbGroups.Clear;
    Stmt := FDatabase.Prepare16(SelectSQL);
    try
      while Stmt.Step = SQLITE_ROW do
        begin
          StringList.AddObject(
            Stmt.Column_Str16(1), // Имя
            TObject(Stmt.Column_Int(0))); // ID, хранимое как TObject
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
// Расчёт статистики прогулов
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    cHoursU := 0; // счётчик уважительных прогулов
    cHoursN := 0; // счётчик неуважительных прогулов
    CurrMonth := Date; // Получаем текущую дату

    // Заполняем параметры для запроса
    Stmt.Bind_Str16(1, cbStudy.Text); // Выбранный студент
    Stmt.Bind_Str16(2, cbGroups.Text); // Выбранная группа
    // На основе текущей даты определяем первый и последний день месяца
    // и преобразуем полученные даты в юлианский формат
    Stmt.Bind_Double(3, DateToJulianDate(StartOfTheMonth(CurrMonth)));
    Stmt.Bind_Double(4, DateToJulianDate(EndOfTheMonth(CurrMonth)));
    while Stmt.Step = SQLITE_ROW do
      begin
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(3)); //суммируем прогулы
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(4));
      end;
    lblStatU.Caption := IntToStr(cHoursU);
    lblStatN.Caption := IntToStr(cHoursN);
  finally
    Stmt.Free;
  end;
  // Ниже выполняется та же обработка, но для предыдущего месяца
  Stmt2 := FDatabase.Prepare16(SelectSQL);
  try
    cPrevHoursU := 0; // счётчик уважительных прогулов за прошлый месяц
    cPrevHoursN := 0; // счётчик неуважительных прогулов за прошлый месяц
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

// Процедура удаления выделенной группы
procedure TfrmMain.mDelGroupClick(Sender: TObject);
const
  DeleteSQL = 'DELETE FROM groups WHERE ID=?;';
var
  i, Idx: Integer;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  { Заполняем StringList. }
  StringList := cbGroups.Items;

  i := cbGroups.ItemIndex;
  if i < 0 then
    begin
      ShowMessage('Выберите сначала строчку.');
      Exit;
    end;

  if MessageDlg('Удалить?', mtConfirmation, mbOKCancel, 0) = mrOk then
    begin
      Idx := Integer(StringList.Objects[i]);
      { Подготавливаем запрос на удаление. }
      Stmt := FDatabase.Prepare16(DeleteSQL);
      try
        Stmt.Bind_Int(1, Idx);
        { Одно выполнение запроса и удаление записи из базы }
        Stmt.Step;
        { Удаляем элемент из StringList. }
        StringList.Delete(i);
      finally
        Stmt.Free;
      end;
    end;
  // Перегружаем группы
  LoadGroups;
end;

procedure TfrmMain.mExitClick(Sender: TObject);
begin
  // Пункт меню выход. Завершаем приложение
  Application.Terminate;
end;

procedure TfrmMain.N6Click(Sender: TObject);
begin
  // Показываем модальное окно с помощью
  frmAbout.ShowModal;
end;

// Процедура добавления группы
procedure TfrmMain.mAddGroupClick(Sender: TObject);
const
  InsertSQL = 'INSERT INTO groups (gruppa) VALUES (?);';
var
  NewGroupName: String;
  Stmt: TDISQLite3Statement;
begin
  NewGroupName := ''; // Очищаем переменную
  // Выводим диалог запроса для ввода группы
  if InputQuery('Добавить', 'Введите новую группу:', NewGroupName) then
    begin
      { Подготавливаем запрос }
      Stmt := FDatabase.Prepare16(InsertSQL);
      try
        // Заполняем параметры
        Stmt.Bind_Str16(1, NewGroupName);
        { Делаем один проход и вставляем данные в базу }
        Stmt.Step;
      finally
        Stmt.Free;
      end;
    end;
    // Перегружаем группы
    LoadGroups;
end;

// Процедура добавления студента
procedure TfrmMain.mAddStudyClick(Sender: TObject);
const
  InsertSQL = 'INSERT INTO study (fio, gruppa) VALUES (?, ?);';
var
  Stmt: TDISQLite3Statement;
  s: String;
begin
  s := ''; // Очищаем переменную
  // Выводим запрос на ввод данных для добавления
  if InputQuery('Добавить', 'Введите ФИО студента:', s) then
    begin
      { Подготавливаем запрос }
      Stmt := FDatabase.Prepare16(InsertSQL);
      try
        // Заполняем параметры
        Stmt.Bind_Str16(1, s);
        Stmt.Bind_Str16(2, cbGroups.Text);
        { Делаем один проход и вставляем данные }
        Stmt.Step;
      finally
        Stmt.Free;
      end;
    end;
    cbGroupsChange(cbGroups);
end;

// Процедура редактирования студента
procedure TfrmMain.mEditStudyClick(Sender: TObject);
const
  UpdateSql = 'UPDATE study SET fio=? WHERE id=?';
var
  i, Idx: Integer;
  s: string;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  // Заполняем StringList значениями из ComboBox со списком студентов
  StringList := cbStudy.Items;

  i := cbStudy.ItemIndex; // Индекс студента в списке
  s := StringList[i]; // Присваиваем переменной значение индекса

  // Выводи диалог запроса на редактирование данных
  if InputQuery('Изменить', 'Введите новое имя:', s) then
    begin
      // Присваиваем полученный индекс переменной
      Idx := Integer(StringList.Objects[i]);
      // Подготавливаем запрос
      Stmt := FDatabase.Prepare16(UpdateSql);
      try
        // Заполняем параметры запроса
        Stmt.Bind_Str16(1, s);
        Stmt.Bind_Int(2, Idx);
        // Выполняем один проход и обновляем запись в таблице
        Stmt.Step;
        // Заносим новое значение в StringList
        StringList[i] := s;
      finally
        Stmt.Free;
      end;
    end;
end;

// Процедура удаления студента
procedure TfrmMain.mDelStudyClick(Sender: TObject);
const
  DeleteSQL = 'DELETE FROM study WHERE ID=?;';
var
  i, Idx: Integer;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  // Заполняем StringList
  StringList := cbStudy.Items;
  // присваиваем переменной индекс студента
  i := cbStudy.ItemIndex;
  if i < 0 then
    begin
      ShowMessage('Выберите сначала строчку.');
      Exit;
    end;
  // Выводим диалог запроса на удаление
  if MessageDlg('Удалить?', mtConfirmation, mbOKCancel, 0) = mrOk then
    begin
      Idx := Integer(StringList.Objects[i]);
      // Подготавливаем запрос
      Stmt := FDatabase.Prepare16(DeleteSQL);
      try
        Stmt.Bind_Int(1, Idx);
        // Выполняем один проход и удаляем запись из базы
        Stmt.Step;
        // Заносим новое значение в StringList
        StringList.Delete(i);
      finally
        Stmt.Free;
      end;
    end;
end;

// Процедура редактирования группы
procedure TfrmMain.mEditGroupClick(Sender: TObject);
const
  UpdateSql = 'UPDATE groups SET gruppa=? WHERE id=?';
var
  i, Idx: Integer;
  s: string;
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  // Заполняем StringList
  StringList := cbGroups.Items;

  i := cbGroups.ItemIndex;
  s := StringList[i];
  // Выводим диалог для изменения значения
  if InputQuery('Изменить', 'Введите новое имя:', s) then
    begin
      Idx := Integer(StringList.Objects[i]);
      // Подготавливаем запрос
      Stmt := FDatabase.Prepare16(UpdateSql);
      try
        Stmt.Bind_Str16(1, s);
        Stmt.Bind_Int(2, Idx);
        // Выполняем один проход и обновляем запись в базе
        Stmt.Step;
        // Заносим новое значение в StringList
        StringList[i] := s;
      finally
        Stmt.Free;
      end;
    end;
  // Перегружаем группы
  LoadGroups;
end;

// Функция проверки возможности редактирования
function TfrmMain.CheckCanEdit(const ACol, ARow: Integer): Boolean;
begin
  Result := (ACol >= 0) and (ARow >= 1) and (Integer(StringGrid.Rows[ARow].Objects[ACol]) > 0);
  if not Result then
    begin
      ShowMessage('Нельзя редактировать.');
    end;
end;

// Процедура обработки изменения группы
procedure TfrmMain.cbGroupsChange(Sender: TObject);
const
  GetGroupsSql = 'SELECT id, gruppa FROM groups ORDER BY gruppa COLLATE NoCase;';
  GetStudysSql = 'SELECT id, fio FROM study WHERE gruppa=?';
var
  Stmt: TDISQLite3Statement;
  StringList: TStrings;
begin
  cbStudy.Clear; // Очищаем ComboBox со студентами
  // Ассоциируем StringList с элементами Combobox студентов
  StringList := cbStudy.Items;
  // Начинаем обновление списка
  StringList.BeginUpdate;
  try
    // Подготавливаем запрос
    Stmt := FDatabase.Prepare16(GetStudysSql);
    // Устанавливаем параметры для запроса
    Stmt.Bind_Str16(1, cbGroups.Text);
    try
      while Stmt.Step = SQLITE_ROW do
        begin
          // Создаём объяекты в StringList используя имя и ID объекта
          StringList.AddObject(
            Stmt.Column_Str16(1), // Имя
            TObject(Stmt.Column_Int(0))); // ID, хранится как TObject.
        end;
    finally
      Stmt.Free;
    end;
  finally
    // Заканчиваем обновление списка
    StringList.EndUpdate;
  end;
  btnSelectGroupClick(Self);
end;

// Процедура обработки нажатия клавиш в списке групп
procedure TfrmMain.cbGroupsKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    45: mAddGroupClick(self); // Insert - добавление группы
    46: mDelGroupClick(self); // Del - удаление группы
  end;
end;

// При выборе студента рассчитываем для него статистику
procedure TfrmMain.cbStudyChange(Sender: TObject);
begin
  CalculateStats;
  btnSelectStudyClick(Self);
end;

// Процедура обработки нажатия клавиш в списке студентов
procedure TfrmMain.cbStudyKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    45: mAddStudyClick(self); // Insert - добавление студента
    46: mDelStudyClick(self); // Del - удаление студента
  end;
end;

// При активации формы устанавливаем фокус на список групп
procedure TfrmMain.FormActivate(Sender: TObject);
begin
  cbGroups.SetFocus;
end;

// Процедура выполняется при создании формы
procedure TfrmMain.FormCreate(Sender: TObject);
begin
  // Создаём новый объект БД и связываем его с файлом
  FDatabase := TDISQLite3Database.Create(nil);
  FDatabase.DatabaseName := 'Study.db';

  // Устанавливаем параметры для StringGrid
  StringGrid.ColCount := 5; // 5 колонок
  StringGrid.DefaultColWidth := 120; // ширина колонки по умолчанию 120
  StringGrid.DefaultRowHeight := 19; // высота ряда 19
  // Добавляем возможность изменять ширину колонок
  StringGrid.Options := StringGrid.Options + [goColSizing];
  StringGrid.FixedCols := 0;
  StringGrid.RowCount := 2;
  // Указываем имена колонок
  SetGridRow(0, -1, ColumnNames[0], ColumnNames[1], ColumnNames[2], ColumnNames[3], ColumnNames[4]);
  SetGridRow(1);
  // Загружаем группы
  LoadGroups;
  // Устанавливаем текущую дату для всех компонентов даты
  dtpAddDelay.Date := Date;
  dtpFrom.Date := Date;
  dtpTo.Date := Date;
  dtpSearchFrom.Date := Date;
  dtpSearchTo.Date := Date;
  // Устанавливаем первую вкладку активной
  PageControl1.ActivePageIndex := 0;
  // Загружаем прогулы
  DelaysSelect;
end;

// Процедура выполняется перед уничтожение формы
procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  // Закрываем связь с базой данных
  FDatabase.Close;
  // Освобождаем память после удаления связи
  FDatabase.Free;
  // Завершаем приложение
  Application.Terminate;
end;

// При показе формы устанавливаем фокус на список групп
procedure TfrmMain.FormShow(Sender: TObject);
begin
  cbGroups.SetFocus;
end;

// Загрузка прогулов
procedure TfrmMain.DelaysSelect;
const
  SelectSQL = 'SELECT id, fio, gruppa, dod, hours_u, hours_n FROM delays ORDER BY gruppa COLLATE NoCase;';
var
  ID, RowCount: Integer;
  Stmt: TDISQLite3Statement;
begin
  // Подготавливаем запрос
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    RowCount := 1;
    while Stmt.Step = SQLITE_ROW do
      begin
        ID := Stmt.Column_Int(0);
        // Заполняем StringGrid в цикле
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

// Процедура добавления прогула
procedure TfrmMain.btnAddDelayClick(Sender: TObject);
const
  InsertSQL = 'INSERT INTO delays (fio, gruppa, dod, hours_u, hours_n) VALUES (?, ?, ?, ?, ?);';
var
  Stmt: TDISQLite3Statement;
begin
  if ((cbGroups.Text <> '') or (cbStudy.Text <> '')) then
    begin
      // Подготавливаем запрос
      Stmt := FDatabase.Prepare16(InsertSQL);
      try
        // Устанавливаем параметры запроса
        Stmt.Bind_Str16(1, cbStudy.Text);
        Stmt.Bind_Str16(2, cbGroups.Text);
        Stmt.Bind_Double(3, DateToJulianDate(dtpAddDelay.Date));
        Stmt.Bind_Int(4, StrToInt(edtHoursU.Text));
        Stmt.Bind_Int(5, StrToInt(edtHoursN.Text));
        // Выполняем проход и заносим данные в базу
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
  // Подготавливаем запрос
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
    StatusBar1.Panels[0].Text := 'Общее число отображённых пропусков: По уважительной - ' + IntToStr(cHoursU) + '; По неуважительной - ' + IntToStr(cHoursN) ;
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
    StatusBar1.Panels[0].Text := 'Общее число отображённых пропусков: По уважительной - ' + IntToStr(cHoursU) + '; По неуважительной - ' + IntToStr(cHoursN) ;
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
    StatusBar1.Panels[0].Text := 'Общее число отображённых пропусков: По уважительной - ' + IntToStr(cHoursU) + '; По неуважительной - ' + IntToStr(cHoursN) ;
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
    StatusBar1.Panels[0].Text := 'Общее число отображённых пропусков: По уважительной - ' + IntToStr(cHoursU) + '; По неуважительной - ' + IntToStr(cHoursN) ;
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

  if MessageDlg('Удалить?', mtConfirmation, mbOKCancel, 0) = mrOk then
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

// Процедура вывода данных в Excel
procedure TfrmMain.btnPrintClick(Sender: TObject);
var
  Stmt: TDISQLite3Statement;
  ID, RowCount, cnt, BeginRow, BeginCol, cHoursU, cHoursN: Integer;
  SelectSQL: String;
  ExcelApp, Cell1, Cell2, ColumnRange: Variant;
begin
  RowCount := 1;
  // Координаты левого верхнего угла области, в которую будем выводить данные
  BeginCol := 1;
  BeginRow := 5;

  // Проверяем чекбоксы и в зависимости от их состояния используем нужный запрос
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
  // Подготавливаем запрос
  Stmt := FDatabase.Prepare16(SelectSQL);
  try
    // Инициализираем экземпляр Excel
    ExcelApp :=CreateOleObject('Excel.Application');
    // Запрещаем обрабатывать события пока выводятся данные
    ExcelApp.Application.EnableEvents := false;
    // Добавляем книгу
    ExcelApp.WorkBooks.Add;
    // Устанавливаем диапазн столбцов
    ColumnRange := ExcelApp.Workbooks[1].WorkSheets[1].Columns;
    // Устанавливаем ширину столбцов в нужное нам значение
    ColumnRange.Columns[1].ColumnWidth := 20;
    ColumnRange.Columns[2].ColumnWidth := 10;
    ColumnRange.Columns[3].ColumnWidth := 15;
    ColumnRange.Columns[4].ColumnWidth := 15;
    ColumnRange.Columns[5].ColumnWidth := 15;

    // Выводим текст в первый ряд вторую ячейку
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 2] := 'Отчёт по пропускам';
    // Создаём диапазон для выделения
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[1, 2];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[1, 4];
    // Выделяем диапазон указанных ячеек
    ExcelApp.ActiveWorkBook.WorkSheets[1].Range[Cell1, Cell2].Select;
    // Объединяем выделенные ячейки
    ExcelApp.Selection.MergeCells := True;
    // Устанавливаем жирный шрифт выделенных ячеек
    ExcelApp.Selection.Font.Bold:=True;
    // Устанавливаем выравнивание по центру
    ExcelApp.Selection.HorizontalAlignment:=3;

    // Создаём диапазон для выделения
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 1];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 5];
    // Выделяем диапазон указанных ячеек
    ExcelApp.ActiveWorkBook.WorkSheets[1].Range[Cell1, Cell2].Select;
    // Заполняем ячейки значениями
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 1] := 'За период';
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 2] := 'c:';
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 3] := FormatDateTime('dd.mmm.yyyy', dtpSearchFrom.Date);
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 4] := 'по:';
    ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[3, 5] := FormatDateTime('dd.mmm.yyyy', dtpSearchTo.Date);
    // Устанавливаем выравнивание по правому краю
    ExcelApp.Selection.HorizontalAlignment := 4;

    // Заполняем шапку для вывода
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 1] := 'ФИО';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 2] := 'Группа';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 3] := 'Дата';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 4] := 'Уважительно';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, 5] := 'Неважительно';
    // Создаём диапазон для выделения
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, BeginCol];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow, BeginCol+4];
    // Выделяем диапазон указанных ячеек
    ExcelApp.WorkBooks[1].WorkSheets[1].Range[Cell1, Cell2].Select;
    // Устанавливаем жирный шрифт выделенных ячеек
    ExcelApp.Selection.Font.Bold:=True;
    // Устанавливаем выравнивание по центру
    ExcelApp.Selection.HorizontalAlignment:=3;

    // Заполняем параметры для запроса
    Stmt.Bind_Double(1, DateToJulianDate(dtpSearchFrom.Date));
    Stmt.Bind_Double(2, DateToJulianDate(dtpSearchTo.Date));
    // В зависимости от выбранных чекбоксов заполняем дополнительные поля
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
        // Заполняем StringGrid
        SetGridRow(
          RowCount,
          ID,
          Stmt.Column_Str16(1),
          Stmt.Column_Str16(2),
          DateToStr(JulianDateToDate(Stmt.Column_Double(3))),
          Stmt.Column_Str16(4),
          Stmt.Column_Str16(5));
          // Выводим полученные данные в Excel
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 1] := Stmt.Column_Str16(1);
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 2] := Stmt.Column_Str16(2);
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 3] := DateToStr(JulianDateToDate(Stmt.Column_Double(3)));
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 4] := Stmt.Column_Str16(4);
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 5] := Stmt.Column_Str16(5);
        Inc(RowCount);
        // Расчитываем число прогулов
        cHoursU := cHoursU + StrToInt(Stmt.Column_Str16(4));
        cHoursN := cHoursN + StrToInt(Stmt.Column_Str16(5));
      end;
    StringGrid.RowCount := RowCount;
    // Выводим статистику на панель
    StatusBar1.Panels[0].Text := 'Общее число отображённых пропусков: По уважительной - ' + IntToStr(cHoursU) + '; По неуважительной - ' + IntToStr(cHoursN) ;

    // Выводим строку с итогом и применяем форматирование
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 1] := 'Итого:';
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 4] := cHoursU;
    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, 5] := cHoursN;
    Cell1 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, BeginCol];
    Cell2 := ExcelApp.WorkBooks[1].WorkSheets[1].Cells[BeginRow + RowCount, BeginCol+4];
    ExcelApp.WorkBooks[1].WorkSheets[1].Range[Cell1, Cell2].Select;
    ExcelApp.Selection.Font.Bold:=True;

    // Показываем Excel
    ExcelApp.Visible := true;
    // Удаляем связь с приложением чтобы оно не висело в памяти
    ExcelApp:=Unassigned;

  finally
    Stmt.Free;
  end;
end;

initialization
  // Инициализируем DISQLite3 библиотеку DISQLite3
  sqlite3_initialize;

finalization
  // Освобождаем память и отключаемся от библиотеки
  sqlite3_shutdown;

end.

