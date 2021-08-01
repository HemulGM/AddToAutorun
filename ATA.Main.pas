unit ATA.Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Winapi.ActiveX, HGM.Controls.VirtualTable, Vcl.Grids, Vcl.ExtCtrls,
  HGM.Controls.Labels, Vcl.Imaging.pngimage, Vcl.StdCtrls, HGM.Button,
  System.ImageList, Vcl.ImgList, System.Win.Registry,
  System.Generics.Collections;

type
  TAutorunFileItem = record
    Name: string;
    FullName: string;
    RootKey: HKEY;
  end;

  TAutoruns = TTableData<TAutorunFileItem>;

  TFormMain = class(TForm, IDropTarget)
    PanelDrag: TPanel;
    LabelExDrag: TLabelEx;
    ImageDrop: TImage;
    PanelClient: TPanel;
    ImageIcon: TImage;
    LabelName: TLabel;
    LabelFullName: TLabel;
    PanelStart: TPanel;
    Label1: TLabel;
    ImageList24: TImageList;
    CheckBoxCurrentUser: TCheckBoxFlat;
    CheckBoxLocalMachine: TCheckBoxFlat;
    Image1: TImage;
    Label2: TLabel;
    Label3: TLabel;
    PanelHintNoSupport: TPanel;
    LabelInfo: TLabel;
    Image2: TImage;
    ButtonFlatCloseHint2: TButtonFlat;
    Panel1: TPanel;
    ButtonFlatAdd: TButtonFlat;
    ButtonFlatCloseFile: TButtonFlat;
    Label5: TLabel;
    PanelAutorun: TPanel;
    ButtonFlatCloseAutoruns: TButtonFlat;
    TableExAutoruns: TTableEx;
    Panel2: TPanel;
    ButtonFlatInfo: TButtonFlat;
    ButtonFlatAutorunList: TButtonFlat;
    PanelInfo: TPanel;
    Label6: TLabel;
    Image3: TImage;
    ButtonFlat1: TButtonFlat;
    PanelSuccess: TPanel;
    LabelSuccess: TLabel;
    Image4: TImage;
    ButtonFlatOk: TButtonFlat;
    TimerAnimate: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CheckBoxCurrentUserClick(Sender: TObject);
    procedure CheckBoxLocalMachineClick(Sender: TObject);
    procedure ButtonFlatAddClick(Sender: TObject);
    procedure ButtonFlatCloseHint2Click(Sender: TObject);
    procedure ButtonFlatCloseFileClick(Sender: TObject);
    procedure TableExAutorunsDrawCellData(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure ButtonFlatCloseAutorunsClick(Sender: TObject);
    procedure ButtonFlatAutorunListClick(Sender: TObject);
    procedure TableExAutorunsItemColClick(Sender: TObject; MouseButton: TMouseButton; const Index: Integer);
    procedure ButtonFlatInfoClick(Sender: TObject);
    procedure ButtonFlatOkClick(Sender: TObject);
    procedure ButtonFlat1Click(Sender: TObject);
    procedure TimerAnimateTimer(Sender: TObject);
  private
    FFile: TAutorunFileItem;
    FAutoruns: TAutoruns;
    FActionInfoBox: TProc;
    FAnimateDir: Boolean;
    function CreateInputFile(FileName: string): TAutorunFileItem;
    procedure Clear;
    function ExistCurrentApp: Integer;
    procedure AddCurrentToAutoRun;
    procedure PanelSplash(Panel: TPanel);
    procedure ShowInfo(Text: string; CloseAction: TProc = nil);
  public
    procedure ShowDragPanel;
    procedure HideDragPanel;
    procedure ProcessFile(FileName: string);
    // IDropTarget
    function DragEnter(const DataObj: IDataObject; grfKeyState: Longint; Pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    {$WARNINGS OFF}
    function DragOver(grfKeyState: Longint; pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    {$WARNINGS ON}
    function DragLeave: HResult; stdcall;
    function Drop(const DataObj: IDataObject; grfKeyState: Longint; Pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    // IUnknown
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
  end;

const
  BinExtArray: TArray<string> = ['.exe', '.bat', '.cmd', '.scr'];

var
  FormMain: TFormMain;

implementation

uses
  Winapi.ShellAPI, HGM.Common.Utils, Winapi.CommCtrl, HGM.ArrayHelper,
  HGM.WinAPI;

{$R *.dfm}

function RunAsAdmin(Window: HWND; FileName: string; Parameters: string): Boolean;
var
  ShellExeInfo: TShellExecuteInfo;
begin
  ZeroMemory(@ShellExeInfo, SizeOf(ShellExeInfo));
  ShellExeInfo.cbSize := SizeOf(TShellExecuteInfo);
  ShellExeInfo.Wnd := Window;
  ShellExeInfo.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
  ShellExeInfo.lpVerb := PChar('runas');
  ShellExeInfo.lpFile := PChar(FileName);
  ShellExeInfo.lpParameters := PChar(Parameters);
  ShellExeInfo.nShow := SW_SHOWNORMAL;
  Result := ShellExecuteEx(@ShellExeInfo);
end;

function GetImageListSH(SHIL_FLAG: Cardinal): HIMAGELIST;
const
  IID_IImageList: TGUID = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';
type
  _SHGetImageList = function(iImageList: integer; const riid: TGUID; var ppv: Pointer): hResult; stdcall;
var
  Handle: THandle;
  SHGetImageList: _SHGetImageList;
begin
  Result := 0;
  Handle := LoadLibrary(shell32);
  if Handle <> S_OK then
  try
    SHGetImageList := GetProcAddress(Handle, PChar(727));
    if Assigned(SHGetImageList) and (Win32Platform = VER_PLATFORM_WIN32_NT) then
      SHGetImageList(SHIL_FLAG, IID_IImageList, Pointer(Result));
  finally
    FreeLibrary(Handle);
  end;
end;

procedure GetIconFromFile(aFile: string; var aIcon: TIcon; SHIL_FLAG: Cardinal);
var
  aImgList: HIMAGELIST;
  SFI: TSHFileInfo;
begin
  SHGetFileInfo(PChar(aFile), FILE_ATTRIBUTE_NORMAL, SFI, SizeOf(TSHFileInfo), SHGFI_ICON or SHGFI_LARGEICON or SHGFI_SHELLICONSIZE or SHGFI_SYSICONINDEX or SHGFI_TYPENAME or SHGFI_DISPLAYNAME);
  if not Assigned(aIcon) then
    aIcon := TIcon.Create;
  aImgList := GetImageListSH(SHIL_FLAG);
  aIcon.Handle := ImageList_GetIcon(aImgList, SFI.iIcon, ILD_NORMAL);
end;


{ TFormMain }

function TFormMain.DragEnter(const DataObj: IDataObject; grfKeyState: Integer; Pt: TPoint; var dwEffect: Integer): HResult;
var
  FmtEtc: TFormatEtc;
  Medium: TStgMedium;
begin
  FmtEtc.cfFormat := CF_HDROP;
  FmtEtc.ptd := nil;
  FmtEtc.dwAspect := DVASPECT_CONTENT;
  FmtEtc.lindex := -1;
  FmtEtc.tymed := TYMED_HGLOBAL;
  if DataObj.GetData(FmtEtc, Medium) <> S_OK then
    Exit(S_FALSE);
  ShowDragPanel;
  dwEffect := DROPEFFECT_LINK;
  Result := S_OK;
end;

function TFormMain.DragLeave: HResult;
begin
  HideDragPanel;
  Result := S_OK;
end;

function TFormMain.DragOver(grfKeyState: Integer; Pt: TPoint; var dwEffect: Integer): HResult;
begin
  ShowDragPanel;
  dwEffect := DROPEFFECT_LINK;
  Result := S_OK;
end;

function TFormMain.Drop(const DataObj: IDataObject; grfKeyState: Integer; Pt: TPoint; var dwEffect: Integer): HResult;
var
  FmtEtc: TFormatEtc;
  Medium: TStgMedium;
  i: Integer;
  FileNameLength: Integer;
  FileName: string;
  FileList: TStringList;
begin
  //Структура дропа файлов
  FmtEtc.cfFormat := CF_HDROP;
  FmtEtc.ptd := nil;
  FmtEtc.dwAspect := DVASPECT_CONTENT;
  FmtEtc.lindex := -1;
  FmtEtc.tymed := TYMED_HGLOBAL;
  //Если структура нужная нам
  if DataObj.GetData(FmtEtc, Medium) = S_OK then
  begin
    FileList := TStringList.Create;
    try
      for i := 0 to DragQueryFile(Medium.hGlobal, $FFFFFFFF, nil, 0) - 1 do
      begin
        FileNameLength := DragQueryFile(Medium.hGlobal, i, nil, 0);
        SetLength(FileName, FileNameLength);
        DragQueryFile(Medium.hGlobal, i, PChar(FileName), FileNameLength + 1);
        //Только исполнительные файлы

        if TArrayHelp.InArray<string>(BinExtArray, AnsiLowerCase(ExtractFileExt(FileName))) then
          FileList.Add(FileName)
        else if AnsiLowerCase(ExtractFileExt(FileName)) = '.lnk' then
        begin
          FileName := GetFileNameFromLink(FileName);
          if TArrayHelp.InArray<string>(BinExtArray, AnsiLowerCase(ExtractFileExt(FileName))) then
            FileList.Add(FileName);
        end;
      end;
      if FileList.Count > 0 then
        ProcessFile(FileList[0])
      else
        ProcessFile('');
    finally
      DragFinish(Medium.hGlobal);
      ReleaseStgMedium(Medium);
      FileList.Free;
    end;
    Result := S_OK;
  end
  else
    Result := S_FALSE;
  HideDragPanel;
end;

procedure TFormMain.FormCreate(Sender: TObject);
var
  FileName: string;
begin
  FAutoruns := TAutoruns.Create(TableExAutoruns);
  RegisterDragDrop(Handle, Self);
  PanelSplash(PanelStart);
  if ParamCount > 0 then
  begin
    FileName := ParamStr(1);
    if FileExists(FileName) then
    begin
      if TArrayHelp.InArray<string>(BinExtArray, AnsiLowerCase(ExtractFileExt(FileName))) then
        ProcessFile(FileName)
      else if AnsiLowerCase(ExtractFileExt(FileName)) = '.lnk' then
      begin
        FileName := GetFileNameFromLink(FileName);
        if TArrayHelp.InArray<string>(BinExtArray, AnsiLowerCase(ExtractFileExt(FileName))) then
          ProcessFile(FileName);
      end;
    end;
  end;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  RevokeDragDrop(Handle);
  FAutoruns.Free;
end;

procedure TFormMain.HideDragPanel;
begin
  TimerAnimate.Enabled := False;
  PanelDrag.Hide;
end;

function TFormMain.CreateInputFile(FileName: string): TAutorunFileItem;
begin
  Result.Name := GetFileDescription(FileName, GetFileNameWoE(FileName));
  Result.FullName := FileName;
end;

function TFormMain.ExistCurrentApp: Integer;
var
  Reg: TRegistry;
  List: TStringList;
  FN: string;

  function CheckReg: Boolean;
  var
    i: Integer;
  begin
    Result := False;
    Reg.OpenKeyReadOnly('Software\Microsoft\Windows\CurrentVersion\Run');
    Reg.GetValueNames(List);
    for i := 0 to List.Count - 1 do
    begin
      if Pos(FN, AnsiLowerCase(Reg.ReadString(List[i]))) <> 0 then
        Exit(True);
    end;
  end;

begin
  FN := AnsiLowerCase(FFile.FullName);
  Reg := TRegistry.Create(KEY_READ);
  List := TStringList.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if CheckReg then
      Result := 1
    else
    begin
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if CheckReg then
        Result := 2
      else
        Result := 0;
    end;
  finally
    begin
      List.Free;
      Reg.Free;
    end;
  end;
end;

procedure TFormMain.AddCurrentToAutoRun;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_WRITE);
  try
    try
      if CheckBoxLocalMachine.Checked then
        Reg.RootKey := HKEY_LOCAL_MACHINE
      else
        Reg.RootKey := HKEY_CURRENT_USER;
      Reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\Run', True);
      Reg.WriteString(FFile.Name, FFile.FullName);
      LabelSuccess.Caption := 'Приложение успешно добавлено в автозагрузку';
      PanelSplash(PanelSuccess);
    except
      on E: Exception do
      begin
        if E is ERegistryException then
        begin
          ShowInfo('Приложение не было добавлено в автозапуск, т.к. у пользователя нет прав. Запустите приложение от имени администратора');
        end
        else
        begin
          ShowInfo('Произошла неизвестная ошибка и приложение не было добавлено в автозапуск. Попробуйте запустить приложение от имени администратора');
        end;
      end;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TFormMain.ButtonFlatCloseAutorunsClick(Sender: TObject);
begin
  Clear;
end;

procedure TFormMain.ButtonFlatCloseFileClick(Sender: TObject);
begin
  Clear;
end;

procedure TFormMain.ButtonFlatCloseHint2Click(Sender: TObject);
begin
  if Assigned(FActionInfoBox) then
  begin
    FActionInfoBox();
    FActionInfoBox := nil;
  end
  else
    PanelHintNoSupport.Hide;
end;

procedure TFormMain.ButtonFlatInfoClick(Sender: TObject);
begin
  PanelSplash(PanelInfo);
end;

procedure TFormMain.ButtonFlatOkClick(Sender: TObject);
begin
  PanelSuccess.Hide;
end;

procedure TFormMain.ShowInfo(Text: string; CloseAction: TProc);
begin
  FActionInfoBox := CloseAction;
  LabelInfo.Caption := Text;
  PanelSplash(PanelHintNoSupport);
end;

procedure TFormMain.ButtonFlat1Click(Sender: TObject);
begin
  PanelInfo.Hide;
end;

procedure TFormMain.ButtonFlatAddClick(Sender: TObject);
begin
  case ExistCurrentApp of
    -1:
      ShowInfo('Нет доступа к реестру');
    0:
      AddCurrentToAutoRun;
    1:
      ShowInfo('Приложение уже добавлено в автозагрузку с текущим пользователем');
    2:
      ShowInfo('Приложение уже добавлено в автозагрузку с любым пользователем');
  end;
end;

procedure TFormMain.PanelSplash(Panel: TPanel);
begin
  Panel.Show;
  Panel.BringToFront;
  Repaint;
end;

procedure TFormMain.ButtonFlatAutorunListClick(Sender: TObject);
var
  Reg: TRegistry;
  List: TStringList;
  Item: TAutorunFileItem;

  procedure AddItems;
  var
    i: Integer;
  begin
    Reg.OpenKeyReadOnly('Software\Microsoft\Windows\CurrentVersion\Run');
    Reg.GetValueNames(List);
    for i := 0 to List.Count - 1 do
    begin
      Item.Name := List[i];
      Item.FullName := Reg.ReadString(Item.Name);
      Item.RootKey := Reg.RootKey;
      FAutoruns.Add(Item);
    end;
  end;

begin
  Reg := TRegistry.Create(KEY_READ);
  List := TStringList.Create;
  FAutoruns.Clear;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    AddItems;
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    AddItems;
  finally
    begin
      List.Free;
      Reg.Free;
    end;
  end;
  PanelSplash(PanelAutorun);
end;

procedure TFormMain.CheckBoxCurrentUserClick(Sender: TObject);
begin
  if CheckBoxCurrentUser.Checked then
  begin
    CheckBoxLocalMachine.Checked := False;
  end
  else
    CheckBoxCurrentUser.Checked := True;
end;

procedure TFormMain.CheckBoxLocalMachineClick(Sender: TObject);
begin
  if CheckBoxLocalMachine.Checked then
  begin
    CheckBoxCurrentUser.Checked := False;
  end
  else
    CheckBoxLocalMachine.Checked := True;
end;

procedure TFormMain.Clear;
begin
  LabelName.Caption := '';
  LabelFullName.Caption := '';
  ImageIcon.Picture.Assign(nil);
  PanelSplash(PanelStart);
end;

procedure TFormMain.ProcessFile(FileName: string);
var
  Icon: TIcon;
begin
  if FileName = '' then
  begin
    ShowInfo('Выбранный вами файл не является исполнительным и не может быть добавлен в Автозагрузку. ' + #13#10 + 'Возможно, это ссылка-протокол, которая используется сторонними программами.', Clear);
    Exit;
  end;
  FFile := CreateInputFile(FileName);
  LabelName.Caption := FFile.Name;
  LabelFullName.Caption := FFile.FullName;
  Icon := TIcon.Create;
  try
    GetIconFromFile(FFile.FullName, Icon, SHIL_EXTRALARGE);
    ImageIcon.Picture.Assign(Icon);
  finally
    Icon.Free;
  end;
  PanelSplash(PanelClient);
end;

procedure TFormMain.ShowDragPanel;
begin
  TimerAnimate.Enabled := True;
  PanelSplash(PanelDrag);
end;

procedure TFormMain.TableExAutorunsDrawCellData(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  R: TRect;
  S: string;
begin
  if not FAutoruns.IndexIn(ARow) then
    Exit;
  with TableExAutoruns.Canvas do
  begin
    case ACol of
      0:
        begin
          Font.Size := 12;
          Font.Color := clBlack;
          R := Rect;
          R.Inflate(-10, 0);
          R.Height := Rect.Height div 2;
          S := FAutoruns[ARow].Name;
          TextRect(R, S, [tfSingleLine, tfLeft, tfVerticalCenter, tfEndEllipsis]);

          Font.Size := 10;
          Font.Color := $00585858;
          R.Offset(0, Rect.Height div 2);
          S := FAutoruns[ARow].FullName;
          TextRect(R, S, [tfSingleLine, tfLeft, tfVerticalCenter, tfEndEllipsis]);
        end;
      1:
        begin
          ImageList24.Draw(TableExAutoruns.Canvas, Rect.CenterPoint.X - 12, Rect.CenterPoint.Y - 12, 3, True);
        end;
    end;
  end;
end;

procedure TFormMain.TableExAutorunsItemColClick(Sender: TObject; MouseButton: TMouseButton; const Index: Integer);
var
  Id: Integer;
  Reg: TRegistry;
begin
  if Index <> 1 then
    Exit;
  Id := TableExAutoruns.ItemIndex;
  if not FAutoruns.IndexIn(Id) then
    Exit;
  if MessageBox(Handle, 'Вы действительно хотите удалить выбранный элемент из автозагрузки?', 'Внимание', MB_ICONWARNING or MB_YESNO) <> ID_YES then
    Exit;
  Reg := TRegistry.Create(KEY_ALL_ACCESS);
  try
    Reg.RootKey := FAutoruns[Id].RootKey;
    if Reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\Run\', False) then
      if Reg.DeleteValue(FAutoruns[Id].Name) then
        FAutoruns.Delete(Id);
  finally
    Reg.Free;
  end;
end;

procedure TFormMain.TimerAnimateTimer(Sender: TObject);
begin
  if FAnimateDir then
  begin //вниз
    ImageDrop.Top := ImageDrop.Top + 7;
    if ImageDrop.BoundsRect.Bottom >= LabelExDrag.BoundsRect.Bottom then
    begin
      FAnimateDir := False;
    end;
  end
  else //вверх
  begin
    ImageDrop.Top := ImageDrop.Top - 4;
    if ImageDrop.BoundsRect.Top <= LabelExDrag.BoundsRect.Top then
    begin
      FAnimateDir := True;
    end;
  end;
  PanelDrag.Repaint;
end;

function TFormMain._AddRef: Integer;
begin
  Result := S_FALSE;
end;

function TFormMain._Release: Integer;
begin
  Result := S_FALSE;
end;

end.

