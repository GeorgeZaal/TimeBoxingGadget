unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, Menus, Grids, Buttons, StdCtrls, ComCtrls,
  Gauges, MMSystem, ComObj, Math, Vcl.Imaging.pngimage, loadini, Vcl.AppEvnts;

type
  TForm1 = class(TForm)
    SGrid: TStringGrid;
    Image1: TImage;
    Label1: TLabel;
    Timer1: TTimer;
    Memo1: TMemo;
    ComboBox1: TComboBox;
    lb1: TLabel;
    Shape1: TShape;
    lb2: TLabel;
    Stop_Btn: TSpeedButton;
    Pause_Btn: TSpeedButton;
    Start_Btn: TSpeedButton;
    Add_Btn: TSpeedButton;
    Del_Btn: TSpeedButton;
    Copy_Btn: TSpeedButton;
    Up_Btn: TSpeedButton;
    Down_Btn: TSpeedButton;
    Image2: TImage;
    ImgResize: TImage;
    GroupBox1: TGroupBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    Bevel1: TBevel;
    Bevel2: TBevel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    OpenDialog1: TOpenDialog;
    TrayIcon1: TTrayIcon;
    Image3: TImage;
    procedure FormShow(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure SGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SpeedButton10Click(Sender: TObject);
    procedure Image2Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure total_time;
    procedure Del_BtnClick(Sender: TObject);
    procedure Up_BtnClick(Sender: TObject);
    procedure Down_BtnClick(Sender: TObject);
    procedure Copy_BtnClick(Sender: TObject);
    procedure Add_BtnClick(Sender: TObject);
    procedure ImgResizeClick(Sender: TObject);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Label1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure Image1DblClick(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure SpeedButton3MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure TrayIcon1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ApplicationEvents1Minimize(Sender: TObject);
    procedure Image3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  num, min, sec : integer;
  pgbMAX, frmMAX, delen, x : integer;
  excel: variant; // Переменная в которой создаётся объект EXCEL

implementation

{$R *.dfm}


procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // Сохраняем настройки:
  Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'left', inttostr(form1.Left));
  Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'top', inttostr(form1.top));
  if checkbox3.Checked = true then
    Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'ontop', 'on')
  else
    Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'ontop', 'off');
  if checkbox1.Checked = true then
    Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'stretch', 'on')
  else
    Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'stretch', 'off');
  if checkbox2.Checked = true then
    Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'sound', 'on')
  else
    Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'sound', 'off');
  Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'patch_snd', labelededit2.Text);
  Ch_Param(GetCurrentDir() + '\settings.ini', '=', 'patch_img', labelededit1.Text);

  // Закроем все книги:
  excel.Workbooks.Close;
  // Закрываем Excel:
  excel.Application.quit;
  // Освобождаем интерфейсы:
  excel := Unassigned;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  form1.Height := 623;
  form1.left := strtoint(OpenParam(GetCurrentDir() + '\settings.ini', 'left', '='));
  form1.top := strtoint(OpenParam(GetCurrentDir() + '\settings.ini', 'top', '='));
  if OpenParam(GetCurrentDir() + '\settings.ini', 'ontop', '=') = 'on' then
    begin
      form1.FormStyle := fsStayOnTop;
      checkbox3.Checked := true;
    end
    else
    begin
      form1.FormStyle := fsNormal;
      checkbox3.Checked := false;
    end;
  if OpenParam(GetCurrentDir() + '\settings.ini', 'stretch', '=') = 'on' then
    begin
      Image1.Stretch := true;
      checkbox1.Checked := true;
    end
    else
    begin
      Image1.Stretch := false;
      checkbox1.Checked := false;
    end;
  if OpenParam(GetCurrentDir() + '\settings.ini', 'sound', '=') = 'on' then
    checkbox2.Checked := true
    else
    checkbox2.Checked := false;
  labelededit2.Text := OpenParam(GetCurrentDir() + '\settings.ini', 'patch_snd', '=');
  LabeledEdit1.Text := OpenParam(GetCurrentDir() + '\settings.ini', 'patch_img', '=');
  timer1.Interval := strtoint(OpenParam(GetCurrentDir() + '\settings.ini', 'interval', '='));

end;

procedure TForm1.FormShow(Sender: TObject);
var
  sh_cnt, i : integer;
begin
  try
    // -------------------------- ОТКРЫТИЕ EXCEL -----------------------------------
    excel := CreateOleObject('Excel.Application');             // создаем обьект EXCEL
    excel.WorkBooks.Open(GetCurrentDir() + '\Программы.xlsx'); // или загружаем его из директории с программой
    //----------------
    sh_cnt := Excel.Sheets.Count;
    for i := 1 to sh_cnt do
      ComboBox1.Items.Add(excel.WorkBooks[1].Sheets[i].Name);
    ComboBox1.ItemIndex := 0;
  //----------------
  Except
    // обрабатываем ошибки:
    showmessage('Внимание! Произошла ошибка при открытия файла с программами!');
    excel.Workbooks.Close;    // закроем все книги
    excel.Application.quit;   // закрываем Excel
    excel := Unassigned;      // освобождаем интерфейсы
  end;
  // ------------------------ КОНЕЦ РАБОТЫ С EXCEL -------------------------------
  // Загрузка фона:
  //  image1.Picture.LoadFromFile(GetCurrentDir() + '\fone.jpg');
  image1.Picture.LoadFromFile(labelededit1.Text);
  // Загрузка иконок:
  Add_Btn.Glyph.LoadFromFile(GetCurrentDir() + '\Icons\add.bmp');
  Del_Btn.Glyph.LoadFromFile(GetCurrentDir() + '\Icons\del.bmp');
  Copy_Btn.Glyph.LoadFromFile(GetCurrentDir() + '\Icons\copy.bmp');
  Up_Btn.Glyph.LoadFromFile(GetCurrentDir() + '\Icons\up.bmp');
  Down_Btn.Glyph.LoadFromFile(GetCurrentDir() + '\Icons\down.bmp');
  // Установка таблицы:
  SGrid.ColWidths[0] := 20;
  SGrid.ColWidths[1] := 280;
  SGrid.ColWidths[2] := 63;
  SGrid.Cells[1,0] := 'Упражнение';
  SGrid.Cells[2,0] := 'T = 0';
  SGrid.Cells[0,1] := '>>';
  num := 1;
  sec := 60;
  // Установка label-ов на экране таймера:
  memo1.Font.Size := lb1.Font.Size;
  memo1.Font.Name := lb1.Font.Name;
  lb2.Font.Size := lb1.Font.Size;
  lb2.Font.Name := lb1.Font.Name;
  // Загрузить таблицу сразу:
  ComboBox1Change(Form1);
  // Считаем общее время:
  total_time;
end;

procedure TForm1.SpeedButton8Click(Sender: TObject);
var
  i : integer;
begin
  // Делаем запуск с выделенной ячейки:
//  SpeedButton10Click(Form1);                               // Клавиша "Стоп", чтобы остановить таймер и всё очистить
  for i := 1 to SGrid.ColCount do                          // Очищаем 1 колонку с указателем
    SGrid.Cells[0, i] := '';
  SGrid.Cells[0, SGrid.Row] := '>>';                       // Ставим указатель на строчку, где фокус
  num := SGrid.Row;                                        // num - глобальная переменная с номером рабочей строки
  min := strtoint(SGrid.Cells[2, num]) - 1;                // Время устанавливается из рабочей строки
  memo1.Lines.Text := SGrid.Cells[1, num];                 // Указывается задание из 2-й колонки
  // ---------------------------------
  Timer1.Enabled := true;
end;

procedure TForm1.SpeedButton9Click(Sender: TObject);
begin
  Timer1.Enabled := false;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
  addr_snd : pchar;
begin
  // Адрес файла со звуком, преобразуем из String в PChar:
  addr_snd := PChar(labelededit2.Text);

  x := x + 1;

  if SGrid.Cells[2, num] = '' then
    SpeedButton10Click(sender);       //Плохо работает

  sec := sec - 1;
  if sec = 0 then
    begin
      min := min - 1;
      sec := 60;
    end;

  if min <> -1 then
    label1.Caption := inttostr(min) + ':' + inttostr(sec);

  if min = -1 then
    begin
      num := num + 1;
      SGrid.Cells[0, num - 1] := '';
      SGrid.Cells[0, num] := '>>';
      if SGrid.Cells[2, num] <> '' then
        min := strtoint(SGrid.Cells[2, num]) - 1;
      memo1.Lines.Text := SGrid.Cells[1, num];
      if checkbox2.Checked = true then
        PlaySound(addr_snd, 0, SND_FILENAME);
      // Уведомление в трее:
      if trayIcon1.visible = true then
        begin
          trayicon1.balloontitle:=('Начинаем: ' + SGrid.Cells[1, num] + ' - ' + SGrid.Cells[2, num] + ' мин.');

          trayicon1.balloonhint:=('Задача завершена: ' + SGrid.Cells[1, num - 1]);
          trayicon1.showballoonHint;// показываем наше уведомление
        end;
      x := 0;
    end;

  // Вывод задач на label'ы (их два):
  if memo1.Lines.Count = 1 then
     begin
      lb1.Caption := memo1.Lines[0];
      lb2.Visible := false;
     end;
  if memo1.Lines.Count > 1 then
    begin
      lb1.Caption := memo1.Lines[0];
      lb2.Caption := memo1.Lines[1];
    end;
end;

procedure TForm1.total_time;
var
  i, sum : integer;
  t_time_hour, t_time_min : string;
begin
  sum := 0;
  for i := 1 to SGrid.RowCount do
    if SGrid.Cells[2, i] <> '' then
      begin
        // Суммируем время:
        sum := sum + strtoint(SGrid.Cells[2, i]);
        // Если общее время больше часа:
        if sum >= 60 then
        begin
          t_time_hour := inttostr(sum div 60);
          t_time_min := inttostr(sum - strtoint(t_time_hour)*60);
        end
        // Если общее время меньше часа:
        else
        begin
          t_time_hour := '00';
          t_time_min := inttostr(sum);
        end;
        // Если нужно, добавляем нули:
        if length(t_time_hour) < 2 then
          t_time_hour := '0' + t_time_hour;
        if length(t_time_min) < 2 then
          t_time_min := '0' + t_time_min;
        // Выводим время:
        SGrid.Cells[2,0] := t_time_hour + ':' + t_time_min;
      end;
end;

procedure TForm1.TrayIcon1DblClick(Sender: TObject);
begin
form1.show;
trayicon1.visible := false;
end;

procedure TForm1.TrayIcon1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
   trayicon1.hint := memo1.Lines.Text + ' - ' + label1.Caption;
end;

procedure TForm1.SGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  // Раскраска таблицы:
  if (ACol > 0) and (ARow > 0) then
    if ARow mod 2 = 0 then
      begin
        sgrid.Canvas.Brush.Color := $333333; //$00F0F0F0;
        sgrid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
        sgrid.Cells[ACol, ARow]);
      end;
end;

procedure TForm1.SpeedButton10Click(Sender: TObject);
var
  i : integer;
begin
  timer1.Enabled := false;
  label1.Caption := '00:00';
  sec := 60;
  num := 1;
  for i := 1 to SGrid.RowCount do
    SGrid.Cells[0, i] := '';
  SGrid.Cells[0, 1] := '>>';
end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
  OpenDialog1.InitialDir := GetCurrentDir();
  if OpenDialog1.Execute then
    LabeledEdit1.text := OpenDialog1.FileName;
  image1.Picture.LoadFromFile(LabeledEdit1.text);
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
begin
  OpenDialog1.InitialDir := GetCurrentDir();
  if OpenDialog1.Execute then
    LabeledEdit2.text := OpenDialog1.FileName;
end;

procedure TForm1.SpeedButton3MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
   trayicon1.hint := lb1.Caption + ' ' + lb2.Caption + ' - ' + label1.Caption;
end;

procedure TForm1.Image1DblClick(Sender: TObject);
begin
  ImgResize.OnClick(Form1);
end;

procedure TForm1.Image1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  Form1.perform(WM_SysCommand, $F012, 0);
end;

procedure TForm1.Image2Click(Sender: TObject);
begin
  if form1.Height = 623 then
  begin
    form1.Height := 845;
  end
  else
  begin
    form1.Height := 623;
  end;
end;

procedure TForm1.Image3Click(Sender: TObject);
begin
  //form1.hide;
  form1.Visible := false;
  trayicon1.visible := true;
end;

procedure TForm1.ImgResizeClick(Sender: TObject);
begin
  if form1.BorderStyle = bsSingle then
  begin
    form1.BorderStyle := bsNone;
    form1.Height := image1.Height;
    form1.Width := image1.Width - 1;
  end
  else
  begin
    form1.BorderStyle := bsSingle;
    form1.Height := 620;
    form1.Width := 389;
  end;
end;

procedure TForm1.Label1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
   ReleaseCapture;
   Form1.perform(WM_SysCommand,$F012,0);
end;

procedure TForm1.ApplicationEvents1Minimize(Sender: TObject);
begin
//  form1.hide;
//  trayicon1.visible := true;
end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if
    checkbox1.Checked = true then Image1.Stretch := true
  else
    Image1.Stretch := false;
end;

procedure TForm1.CheckBox3Click(Sender: TObject);
begin
  if
    checkbox3.Checked = true then form1.FormStyle := fsStayOnTop
  else
   form1.FormStyle := fsNormal;
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
var
  i : integer;
begin
  // Остановка таймера при смене таблицы:
  if timer1.Enabled = true then
  begin
    SpeedButton10Click(Form1);
    lb1.Caption := 'Начнём что-то новое?';
    // Загрузка в таблицу:
    for i := 1 to 24 do
    begin
      SGrid.Cells[1, i] := excel.WorkBooks[1].WorkSheets[ComboBox1.ItemIndex+1].Cells[i, 1];
      SGrid.Cells[2, i] := excel.WorkBooks[1].WorkSheets[ComboBox1.ItemIndex+1].Cells[i, 2];
    end;
  end
  else
    for i := 1 to 24 do
    begin
      SGrid.Cells[1, i] := excel.WorkBooks[1].WorkSheets[ComboBox1.ItemIndex+1].Cells[i, 1];
      SGrid.Cells[2, i] := excel.WorkBooks[1].WorkSheets[ComboBox1.ItemIndex+1].Cells[i, 2];
     end;
  // Считаем общее время:
  total_time;
end;

// **************************************************************
//
//                 ПРОЦЕДУРЫ ДЛЯ КНОПОК НА ЭКРАНЕ
//
// **************************************************************

// ДОБАВИТЬ СТРОЧКУ:
procedure TForm1.Add_BtnClick(Sender: TObject);
var
  i, sel : integer;
begin
  // Добавление пустой строки:
  //SGrid.RowCount := SGrid.RowCount + 1;
  sel := SGrid.Selection.Top;
  // Сдвиг строчек вниз:
  for i := 25 downto sel + 1 do
    begin
      SGrid.Cells[1, i] := SGrid.Cells[1, i - 1];
      SGrid.Cells[2, i] := SGrid.Cells[2, i - 1];
    end;
  // Собственно пустая строчка:
  SGrid.Cells[1, sel] := '';
  SGrid.Cells[2, sel] := '0';
end;

// КОПИРОВАТЬ СТРОЧКУ:
procedure TForm1.Copy_BtnClick(Sender: TObject);
var
  i, sel : integer;
begin
  // Добавление пустой строки:
  //SGrid.RowCount := SGrid.RowCount + 1;
  sel := SGrid.Selection.Top;
  // Сдвиг строчек вниз:
  for i := 25 downto sel + 1 do
    begin
      SGrid.Cells[1, i] := SGrid.Cells[1, i - 1];
      SGrid.Cells[2, i] := SGrid.Cells[2, i - 1];
    end;
end;

// УДАЛЕНИЕ СТРОЧКИ:
procedure TForm1.Del_BtnClick(Sender: TObject);
var
  i, j : integer;
begin
  for i := SGrid.Selection.Top to SGrid.RowCount-1 do
    for j := 0 to 18 do
      SGrid.Cells[j, i] := SGrid.Cells[j, i+1];
  sgrid.RowCount := sgrid.RowCount-1;
end;

// СТРОЧКА ВВЕРХ:
procedure TForm1.Up_BtnClick(Sender: TObject);
var
  TempList : TStringList;
  i : Integer;
begin
  if SGrid.Selection.Top = 1 then exit else
    begin
    //перемещение строки вверх
      with SGrid do
        if (SGrid.Selection.Top in [0..RowCount - 1]) and
        (SGrid.Selection.Top-1 in [0..RowCount - 1]) then
        begin
          TempList := TStringList.Create;
          TempList.Assign(Rows[SGrid.Selection.Top]);
          if SGrid.Selection.Top > SGrid.Selection.Top-1 then
            for i := SGrid.Selection.Top downto SGrid.Selection.Top-1 + 1 do
              Rows[i].Assign(Rows[i - 1])
          else
           for i := SGrid.Selection.Top to SGrid.Selection.Top-1 - 1 do
             Rows[i].Assign(Rows[i + 1]);
          Rows[SGrid.Selection.Top-1].Assign(TempList);
          TempList.Free;
         end;
  //выделение ячейки после перемещения:
  SGrid.Row := SGrid.Row - 1;
end;
end;

// СТРОЧКА ВНИЗ:
procedure TForm1.Down_BtnClick(Sender: TObject);
var
  TempString : string;
  i : Integer;
begin
  if SGrid.Selection.Top = sgrid.RowCount-1 then
  exit
  else
  begin
  //перемещение строки вниз
    if (SGrid.Selection.Top in [0..SGrid.RowCount - 1]) and
    (SGrid.Selection.Top + 1 in [0..SGrid.RowCount - 1]) then
      begin
        TempString := SGrid.Rows[SGrid.Selection.Top].Text;
        SGrid.Rows[SGrid.Selection.Top].Assign(SGrid.Rows[SGrid.Selection.Top + 1]);
        SGrid.Rows[SGrid.Selection.Top + 1].Text := TempString;
      end;
//выделение ячейки после перемещения:
  SGrid.Row := SGrid.Row + 1;
  end;
end;

end.
