unit ATA.Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Winapi.ActiveX, Winapi.ShellAPI, HGM.Controls.VirtualTable, Vcl.Grids,
  Vcl.ExtCtrls, HGM.Controls.Labels, Vcl.Imaging.pngimage, Vcl.StdCtrls,
  HGM.Button, System.ImageList, Vcl.ImgList, System.Win.Registry;

type
  TFile = record
    Name: string;
    FullName: string;
    RootKey: HKEY;
  end;

  TAutoruns = TTableData<TFile>;

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
    Label4: TLabel;
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
  private
    FFile: TFile;
    FAutoruns: TAutoruns;
    function CreateInputFile(FileName: string): TFile;
    procedure Clear;
    function ExistCurrentApp: Integer;
    procedure AddCurrentToAutoRun;
    procedure PanelSplash(Panel: TPanel);
  public
    procedure ShowDragPanel;
    procedure HideDragPanel;
    procedure ProcessFile(FileName: string);
    // IDropTarget
    function DragEnter(const DataObj: IDataObject; grfKeyState: Longint; Pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    function DragOver(grfKeyState: Longint; Pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    function DragLeave: HResult; stdcall;
    function Drop(const DataObj: IDataObject; grfKeyState: Longint; Pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    // IUnknown
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
  end;

var
  FormMain: TFormMain;
  CF_IDLIST: Integer = 0;

implementation

uses
  HGM.Common.Utils, ComObj, Commctrl, ShlObj;

{$R *.dfm}

const
  IID_IImageList: TGUID = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';

function GetImageListSH(SHIL_FLAG: Cardinal): HIMAGELIST;
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
  HideDragPanel;
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
      try
        for i := 0 to DragQueryFile(Medium.hGlobal, $FFFFFFFF, nil, 0) - 1 do
        begin
          FileNameLength := DragQueryFile(Medium.hGlobal, i, nil, 0);
          SetLength(FileName, FileNameLength);
          DragQueryFile(Medium.hGlobal, i, PChar(FileName), FileNameLength + 1);
          //Только EXE-файлы
          if AnsiLowerCase(ExtractFileExt(FileName)) = '.exe' then
            FileList.Add(FileName)
          else if AnsiLowerCase(ExtractFileExt(FileName)) = '.lnk' then
          begin
            FileName := GetFileNameFromLink(FileName);
            if AnsiLowerCase(ExtractFileExt(FileName)) = '.exe' then
              FileList.Add(FileName);
          end;
        end;
      finally
        DragFinish(Medium.hGlobal);
      end;
    finally
      ReleaseStgMedium(Medium);
    end;
    try
      if FileList.Count > 0 then
        ProcessFile(FileList[0])
      else
        ProcessFile('');
    finally
      FileList.Free;
    end;
    Result := S_OK;
  end
  else
    Result := S_FALSE;
end;

procedure TFormMain.FormCreate(Sender: TObject);
var FileName: string;
begin
  OleInitialize(nil);
  FAutoruns := TAutoruns.Create(TableExAutoruns);
  CF_IDLIST := RegisterClipboardFormat(CFSTR_SHELLIDLIST);
  RegisterDragDrop(Handle, Self);
  PanelSplash(PanelStart);
  if ParamCount > 0 then
  begin
    FileName := ParamStr(1);
    if FileExists(FileName) then
    begin
      if AnsiLowerCase(ExtractFileExt(FileName)) = '.exe' then
        ProcessFile(FileName)
      else if AnsiLowerCase(ExtractFileExt(FileName)) = '.lnk' then
      begin
        FileName := GetFileNameFromLink(FileName);
        if AnsiLowerCase(ExtractFileExt(FileName)) = '.exe' then
          ProcessFile(FileName);
      end;
    end;
  end;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  FAutoruns.Clear;
  FAutoruns.Free;
end;

procedure TFormMain.HideDragPanel;
begin
  PanelDrag.Hide;
end;

function TFormMain.CreateInputFile(FileName: string): TFile;
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
  Result := -1;
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
    if CheckBoxLocalMachine.Checked then
      Reg.RootKey := HKEY_LOCAL_MACHINE
    else
      Reg.RootKey := HKEY_CURRENT_USER;
    Reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\Run', True);
    Reg.WriteString(FFile.Name, FFile.FullName);
    MessageBox(Handle, 'Приложение успешно добавлено в автозагрузку', '', MB_ICONINFORMATION or MB_OK);
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
  Clear;
end;

procedure TFormMain.ButtonFlatInfoClick(Sender: TObject);
begin
  PanelSplash(PanelInfo);
end;

procedure TFormMain.ButtonFlatAddClick(Sender: TObject);
begin
  case ExistCurrentApp of
    -1:
      MessageBox(Handle, 'Нет доступа к реестру', 'Ошибка', MB_ICONERROR or MB_OK);
    0:
      AddCurrentToAutoRun;
    1:
      MessageBox(Handle, 'Приложение уже добавлено в автозагрузку с текущим пользователем', '', MB_ICONWARNING or MB_OK);
    2:
      MessageBox(Handle, 'Приложение уже добавлено в автозагрузку с любым пользователем', '', MB_ICONWARNING or MB_OK);
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
  Item: TFile;

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
  Clear;
  if FileName = '' then
  begin
    PanelSplash(PanelHintNoSupport);
    Exit;
  end;
  FFile := CreateInputFile(FileName);
  FAutoruns.Add(FFile);
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

function TFormMain._AddRef: Integer;
begin
  Result := S_FALSE;
end;

function TFormMain._Release: Integer;
begin
  Result := S_FALSE;
end;

end.

