unit EMF550P.Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Xml.XMLIntf, Vcl.StdCtrls,
  Xml.XMLDoc, System.Generics.Collections, TableDraw, Vcl.Buttons,
  Vcl.ExtCtrls, Vcl.ComCtrls, ShellAPI, ShlObj, ActiveX, System.Win.ComObj, IniFiles,
  LabelButton, Vcl.ImgList, System.ImageList, Vcl.Grids, Xml.xmldom, MSXML;

type

  //�������� ������2 �����������
  TESItem = record
   IdInfoOR:string;
   RES_ES:Integer;
   KOD_ERR:Integer;
   NAM_ERR:string;
  end;
  TESItems = TList<TESItem>;

  //��������� �����������
  TES550PStruct = record
   Output:string;                //�������� ����
   IDNOR:string;                 //
   FILE_ES:string;               //
   SIZE_ES:Cardinal;             //
   DATE_ES:TDate;                //
   RECNO_ES:Integer;             //
   ES_REC:TESItems;              //�������� ������2 �����������
   DATE_KVIT:TDate;              //
   TIME_KVIT:TTime;              //
   OPER:string;                  //
   TEL_OPER:string;              //
  end;

  //��������� �������� � ������� �������� ������
  TInputFileState = (ifsWait, ifsProcces, ifsError, ifsReady);

  //������ �������� ������
  TInputFile = record
   FileName:TFileName;     //��� �����
   State:TInputFileState;  //���������
   Output:TFileName;       //�������� ����
  end;
  TFiles = TTableData<TInputFile>;

  //��� ���. ��������
  TNoteEvent = (neInfo, neWarning, neError);

  //���. ���������
  TNoteItem = record
   Text:string;
   Event:TNoteEvent;
   ItemIndex:Integer;
  end;
  TNotes = TTableData<TNoteItem>;

  TFormMain = class(TForm, IDropTarget)
    XMLDocument: TXMLDocument;
    TableExFiles: TTableEx;
    PanelSaveTo: TPanel;
    Label1: TLabel;
    EditSaveTo: TEdit;
    ImageList: TImageList;
    ComboBoxSave: TComboBox;
    PanelMenu: TPanel;
    ButtonProcess: TButton;
    ButtonAddFiles: TButton;
    PanelReq: TPanel;
    DateTimePickerTime: TDateTimePicker;
    DateTimePickerDate: TDateTimePicker;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    EditINN: TEdit;
    EditKPP: TEdit;
    Label5: TLabel;
    EditOper: TEdit;
    Label6: TLabel;
    EditTel: TEdit;
    LabelWait: TLabel;
    FileOpenDialogInput: TFileOpenDialog;
    Label7: TLabel;
    PanelDrag: TPanel;
    LabelExDrag: TLabelEx;
    TimerUpdateDT: TTimer;
    ButtonHelp: TButton;
    PanelClient: TPanel;
    TableExNote: TTableEx;
    Splitter1: TSplitter;
    Label8: TLabelEx;
    SpeedButtonSaveTo: TButton;
    procedure ButtonProcessClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TableExFilesGetData(FCol, FRow: Integer; var Value: string);
    procedure ComboBoxSaveChange(Sender: TObject);
    procedure SpeedButtonSaveToClick(Sender: TObject);
    procedure DateTimePickerDateClick(Sender: TObject);
    procedure DateTimePickerTimeClick(Sender: TObject);
    procedure TableExFilesItemColClick(Sender: TObject; MouseButton: TMouseButton; const Index: Integer);
    procedure ButtonAddFilesClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure TimerUpdateDTTimer(Sender: TObject);
    procedure ButtonHelpClick(Sender: TObject);
    procedure TableExNoteGetData(FCol, FRow: Integer; var Value: string);
    procedure TableExNoteDrawCellData(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure TableExFilesDblClick(Sender: TObject);
    procedure TableExNoteDblClick(Sender: TObject);
  private
    FFiles:TFiles;                                                              //������ �������� XML-������
    FNotes:TNotes;                                                              //��������� ��� ������������
    FSkipXSDValidCheck:Boolean;                                                 //�������� ������ ��� ������ � �������� �� XSD-�����
    FSkipXSDValid:Boolean;                                                      //��������� ����� �� XSD-�����
    FExistsCLASS_DOMXML:Boolean;                                                //���������� ����� MSXML
    CLASS_DOMDocumentMy:TGUID;                                                  //������ MSXML
    CLASS_XMLSchemaCacheMy:TGUID;
    procedure Process;                                                          //������������ �����������
    function CheckES550PNode(Data: IXMLNode): TESItem;                          //�������� ����� "������2"
    function CreateKVIT(Struct: TES550PStruct): Boolean;                        //������������ � ��������� XML-���� �����������
    procedure SaveSettings;                                                     //��������� ���������
    procedure LoadSettings;                                                     //���������
    procedure ProcessFiles(List:TStrings);                                      //���������� ������ � ������
    procedure ShowDragPanel;                                                    //�������� ������ �����
    procedure HideDragPanel;                                                    //������
    function Note(Text: string; Event: TNoteEvent; ItemIndex:Integer):Boolean;  //�������� � ������� ���������� ������
    procedure IniXML;                                                           //������������� ������ MSXML
  public
    // IDropTarget
    function DragEnter(const DataObj:IDataObject; grfKeyState:Longint; Pt:TPoint; var dwEffect:Longint):HResult; stdcall;
    function DragOver(grfKeyState:Longint; Pt:TPoint; var dwEffect:Longint):HResult; stdcall;
    function DragLeave:HResult; stdcall;
    function Drop(const DataObj:IDataObject; grfKeyState:Longint; Pt:TPoint; var dwEffect:Longint):HResult; stdcall;
    // IUnknown
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
  end;

var
  FormMain: TFormMain;

  function InputFileStateStr(Value:TInputFileState):string;
  function ESItem(aIdInfoOR:string; aRES_ES:Integer):TESItem;
  function CreateInputFile(FileName:string):TInputFile;
  function EventNoteStr(Value:TNoteEvent):string;
  function OpenFolderAndSelectFile(const FileName:string):Boolean;
  function AdvSelectDirectory(const Caption:string; const Root:WideString; var Directory:string; EditBox:Boolean = False; ShowFiles:Boolean = False; AllowCreateDirs:Boolean = True):Boolean;
  function GetFileSize(FileName:string):Integer;

implementation

{$R *.dfm}

function OpenFolderAndSelectFile(const FileName:string):Boolean;
var IIDL:PItemIDList;
begin
 Result:=False;
 IIDL:=ILCreateFromPath(PChar(FileName));
 if IIDL <> nil then
  try
   Result:=SHOpenFolderAndSelectItems(IIDL, 0, nil, 0) = S_OK;
  finally
   ILFree(IIDL);
  end;
end;

function ESItem(aIdInfoOR:string; aRES_ES:Integer):TESItem;
begin
 Result.IdInfoOR:=aIdInfoOR;
 Result.RES_ES:=aRES_ES;
end;

function AdvSelectDirectory(const Caption:string; const Root:WideString; var Directory:string; EditBox:Boolean = False; ShowFiles:Boolean = False; AllowCreateDirs:Boolean = True):Boolean;

function SelectDirCB(Wnd: HWND; uMsg: UINT; lParam, lpData: lParam): Integer; stdcall;
begin
 case uMsg of
  BFFM_INITIALIZED: SendMessage(Wnd, BFFM_SETSELECTION, Ord(True), Integer(lpData));
 end;
 Result:= 0;
end;

var
 WindowList: Pointer;
 BrowseInfo: TBrowseInfo;
 Buffer: PChar;
 RootItemIDList, ItemIDList: PItemIDList;
 ShellMalloc: IMalloc;
 IDesktopFolder: IShellFolder;
 Eaten, Flags: LongWord;

const
 BIF_USENEWUI = $0040;
 BIF_NOCREATEDIRS = $0200;

begin
 Result:= False;
 if not DirectoryExists(Directory) then Directory:= '';
 FillChar(BrowseInfo, SizeOf(BrowseInfo), 0);
 if (ShGetMalloc(ShellMalloc) = S_OK) and (ShellMalloc <> nil) then
  begin
   Buffer:= ShellMalloc.Alloc(MAX_PATH);
   try
    RootItemIDList:= nil;
    if Root <> '' then
     begin
      SHGetDesktopFolder(IDesktopFolder);
      IDesktopFolder.ParseDisplayName(Application.Handle, nil, POleStr(Root), Eaten, RootItemIDList, Flags);
     end;
    OleInitialize(nil);
    with BrowseInfo do
     begin
      hwndOwner:= Application.Handle;
      pidlRoot:= RootItemIDList;
      pszDisplayName := Buffer;
      lpszTitle:= PChar(Caption);
      ulFlags:= BIF_RETURNONLYFSDIRS or BIF_USENEWUI or
         BIF_EDITBOX * Ord(EditBox) or BIF_BROWSEINCLUDEFILES * Ord(ShowFiles) or
         BIF_NOCREATEDIRS * Ord(not AllowCreateDirs);
      lpfn:=@SelectDirCB;
      if Directory <> '' then lParam := Integer(PChar(Directory));
     end;
    WindowList:= DisableTaskWindows(0);
    try
     ItemIDList:= ShBrowseForFolder(BrowseInfo);
    finally
     EnableTaskWindows(WindowList);
    end;
    Result:= ItemIDList <> nil;
    if Result then
     begin
      ShGetPathFromIDList(ItemIDList, Buffer);
      ShellMalloc.Free(ItemIDList);
      Directory:= Buffer;
     end;
   finally
    ShellMalloc.Free(Buffer);
   end;
  end;
end;

function InputFileStateStr(Value:TInputFileState):string;
begin
 case Value of
  ifsWait: Result:='��������';
  ifsProcces: Result:='���������';
  ifsReady: Result:='�����';
  ifsError: Result:='������';
 end;
end;

function EventNoteStr(Value:TNoteEvent):string;
begin
 case Value of
  neInfo: Result:='����������';
  neWarning: Result:='��������������';
  neError: Result:='������';
 end;
end;

function GetFileSize(FileName:string):Integer;
var FS:TFileStream;
begin
 Result:=-1;
 try
  FS:=TFileStream.Create(FileName, fmOpenRead);
  Result:=FS.Size;
  FS.Free;
 except

 end;
end;

function CreateInputFile(FileName:string):TInputFile;
begin
 Result.FileName:=FileName;
 Result.State:=ifsWait;
 Result.Output:='';
end;

function TFormMain.CheckES550PNode(Data:IXMLNode):TESItem;
begin
 Result:=ESItem(Data.ChildNodes['�����������'].Text, 0);
end;

function TFormMain.CreateKVIT(Struct:TES550PStruct):Boolean;
var Root, Recs, Rec:IXMLNode;
    XML:TXMLDocument;
    i:Integer;
    LSep:Char;
begin
 Result:=False;
 //������� "/" ��� ���������� ����. ������.
 LSep:=FormatSettings.DateSeparator;
 FormatSettings.DateSeparator:='/';
 //
 XML:=TXMLDocument.Create(nil);
 XML.Active:=True;
 XML.Version:='1.0';
 XML.Encoding:='UTF-8';
 XML.Options:=[doNodeAutoIndent];
 try
  begin
   if XML.Active then
    begin
     //������ �����
     Root:=XML.AddChild('KVIT');
     //������� ��������� ������
     Root.AddChild('IDNOR').Text:=Struct.IDNOR;
     Root.AddChild('ES').Text:=Struct.FILE_ES;
     Root.AddChild('SIZE_ES').Text:=IntToStr(Struct.SIZE_ES);
     Root.AddChild('DATE_ES').Text:=FormatDateTime('ddddd', Struct.DATE_ES);
     //�������� ������ �������
     Recs:=Root.AddChild('RECNO_ES');
     Recs.Attributes['nRe�']:=IntToStr(Struct.RECNO_ES);
     for i:= 0 to Struct.ES_REC.Count-1 do
      begin
       Rec:=Recs.AddChild('ES_REC');
       Rec.Attributes['IdInfoOR']:=Struct.ES_REC[i].IdInfoOR;
       Rec.AddChild('RES_ES').Text:=IntToStr(Struct.ES_REC[i].RES_ES);
      end;
     //������� ��������� ������
     Root.AddChild('DATE_KVIT').Text:=FormatDateTime('ddddd', Struct.DATE_KVIT);
     Root.AddChild('TIME_KVIT').Text:=FormatDateTime('HH:NN:SS', Struct.TIME_KVIT);
     Root.AddChild('OPER').Text:=Struct.OPER;
     Root.AddChild('TEL_OPER').Text:=Struct.TEL_OPER;
     Result:=True;
     //��������� ����
     XML.SaveToFile(Struct.Output);
     XML.Active:=False;
    end;
  end;
 except
  on E:Exception do Note(E.Message, neError, -1);
 end;
 //����� �����������
 FormatSettings.DateSeparator:=LSep;
end;

function TFormMain.DragEnter(const dataObj:IDataObject; grfKeyState:Longint; pt:TPoint; var dwEffect:Longint):HResult;
var FmtEtc:TFormatEtc;
    Medium:TStgMedium;
begin
 //����� ������, ��������. ���� �� �����, �� ������ �� ����
 FmtEtc.cfFormat:=CF_HDROP;
 FmtEtc.ptd:=nil;
 FmtEtc.dwAspect:=DVASPECT_CONTENT;
 FmtEtc.lindex:=-1;
 FmtEtc.tymed:=TYMED_HGLOBAL;
 if DataObj.GetData(FmtEtc, Medium) <> S_OK then Exit(S_FALSE);
 ShowDragPanel;
 dwEffect:=DROPEFFECT_MOVE;
 Result:=S_OK;
end;

function TFormMain.DragOver(grfKeyState: Longint; pt: TPoint; var dwEffect: Longint):HResult;
begin
 ShowDragPanel;
 dwEffect:=DROPEFFECT_MOVE;
 Result:=S_OK;
end;

function TFormMain.DragLeave:HResult;
begin
 HideDragPanel;
 Result:=S_OK;
end;

function TFormMain._AddRef:Integer;
begin
 Result:=S_FALSE;
end;

function TFormMain._Release:Integer;
begin
 Result:=S_FALSE;
end;

function TFormMain.Drop(const DataObj:IDataObject; grfKeyState:Longint; pt:TPoint; var dwEffect:Longint):HResult;
var FmtEtc:TFormatEtc;
    Medium:TStgMedium;
    i:Integer;
    FileNameLength:Integer;
    FileName:string;
    FileList:TStringList;
begin
 HideDragPanel;
 //��������� ����� ������
 FmtEtc.cfFormat:=CF_HDROP;
 FmtEtc.ptd:=nil;
 FmtEtc.dwAspect:=DVASPECT_CONTENT;
 FmtEtc.lindex:=-1;
 FmtEtc.tymed:=TYMED_HGLOBAL;
 //���� ��������� ������ ���
 if DataObj.GetData(FmtEtc, Medium) = S_OK then
  begin
   FileList:=TStringList.Create;
   try
    try
     for i:=0 to DragQueryFile(Medium.hGlobal, $FFFFFFFF, nil, 0)-1 do
      begin
       FileNameLength:=DragQueryFile(Medium.hGlobal, i, nil, 0);
       SetLength(FileName, FileNameLength);
       DragQueryFile(Medium.hGlobal, i, PChar(FileName), FileNameLength + 1);
       //������ XML-�����
       if AnsiLowerCase(ExtractFileExt(FileName)) = '.xml' then FileList.Add(FileName);
      end;
    finally
     DragFinish(Medium.hGlobal);
    end;
   finally
    ReleaseStgMedium(Medium);
   end;
   try
    ProcessFiles(FileList);
   finally
    FileList.Free;
   end;
   Result:=S_OK;
  end
 else Result:=S_FALSE;
end;

procedure TFormMain.ButtonAddFilesClick(Sender: TObject);
begin
 if FileOpenDialogInput.Execute then ProcessFiles(FileOpenDialogInput.Files);
end;

procedure TFormMain.ButtonHelpClick(Sender: TObject);
begin
 MessageBox(Application.Handle, '��������� ��� ������������ ����������� � �������� (����������) ������������ ��������� ��������� ������������ (�������� ��������� �����������), ��� (�������� ���) ���������������� ���������� ����� ������ � 550-�'+#13+#10+''+#13+#10+'�����������: �������� ������� 2018 (�)'+#13+#10+''+#13+#10+'����������:'+#13+#10+'1. ��� ������ ������ ���������� ��������� ��������� � ����� ����� ���� ���������.'+
  #13#10+'2. ����� ��������� ���������� �������� ����� � �������. ��� ����� ������� ����� ���������:'+#13+#10+' �) ������ ������ "�������� ����..." � ������� ����������� �����'+#13+#10+' �) �������� ����������� ����� � ���������� Windows � ���������� � ���� ���������'+#13+#10+'3. ����� ��������� ���������� ������ ���������� ������� ���� ��������� �������� ����� (����� ���� ���������)'+#13+#10+'4. ������ ������ "������������"'+#13+#10+'5. ������'+#13+#10+''+#13+#10+'����� ��� ��������� � ������ �� ����� ���� ��������� �������� � ����� ��������������', '����������', MB_ICONINFORMATION or MB_OK);
end;

procedure TFormMain.ButtonProcessClick(Sender: TObject);
begin
 //�������� ���������
 if EditINN.Text = '' then if Note('������� ��� �����������', neWarning, -1) then Exit;
 if EditKPP.Text = '' then if Note('������� ��� �����������', neWarning, -1) then Exit;
 if EditOper.Text = '' then if Note('������� ��������� (���)', neWarning, -1) then Exit;
 if EditTel.Text = '' then if Note('������� ���������� ������� ���������', neWarning, -1) then Exit;
 //
 LabelWait.Visible:=True;
 try
  Process;
 except
  on E:Exception do Note(E.Message, neError, -1);
 end;
 LabelWait.Visible:=False;
end;

function TFormMain.Note(Text:string; Event:TNoteEvent; ItemIndex:Integer):Boolean;
var NItem:TNoteItem;
begin
 NItem.Text:=Text;
 NItem.Event:=Event;
 NItem.ItemIndex:=ItemIndex;
 FNotes.Insert(0, NItem);
 FNotes.UpdateTable;
 Result:=True;
end;

procedure TFormMain.Process;
var Nodes:IXMLNodeList;
    i, n:Integer;
    Struct:TES550PStruct;
    FItem:TInputFile;
    BufStr, SDir:string;
var FXMLDocument:IXMLDOMDocument2;
    FXMLDOMSchema:IXMLDOMSchemaCollection2;
begin
 //���� ������� �����
 if FFiles.Count <= 0 then
  begin
   Note('������� �������� ����� �� � ������', neInfo, -1);
   Exit;
  end;
 //���� ������ ����� ���������� � �������, �� ��������� ��� �������������
 if ComboBoxSave.ItemIndex = 1 then
  if not DirectoryExists(EditSaveTo.Text) then
   begin
    Note('������� ���������� �� ������ ��� ������ �������!', neWarning, -1);
    Exit;
   end;
 Note('��������� ��������...', neInfo, -1);
 //������������� �������� ����������
 if FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocumentMy) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCacheMy) as IXMLDOMSchemaCollection2;
  except
   FXMLDocument:=nil;
  end;
 //���� ���-�� �� ���, �� ���������� ������������. ���������� ��� �������� ��� ��������� ������.
 if FXMLDocument = nil then
  begin
   if FSkipXSDValidCheck then
    begin
     FSkipXSDValidCheck:=True;
     if MessageBox(Application.Handle, '��������� �� ������� ���������������� �������� ����� �� XSD-�����. ���� ����������, �������� ����� ���������.'+#13+#10+'�������� ���������� ��������?', '��������', MB_ICONEXCLAMATION or MB_YESNO or MB_DEFBUTTON1) = ID_NO then
      begin
       FSkipXSDValid:=True;
      end
     else Exit;
    end;
  end
 else
  begin
   FSkipXSDValid:=False;
   FXMLDOMSchema.Add('', 'FSFM550P.xsd');
   FXMLDocument.Async:=False;
   FXMLDocument.ValidateOnParse:=True;
  end;
 //�������� ��������� � �������� ���������
 Struct.ES_REC:=TESItems.Create;
 //��� � ���
 Struct.IDNOR:=Format('%s_%s', [EditINN.Text, EditKPP.Text]);
 //��������
 Struct.OPER:=EditOper.Text;
 //����. ���.
 Struct.TEL_OPER:=EditTel.Text;
 //�� ������
 for i:= 0 to FFiles.Count-1 do
  begin
   FItem:=FFiles[i];
   //������� � ���������
   FItem.State:=ifsProcces;
   //�������
   FFiles[i]:=FItem;
   FFiles.UpdateTable;
   Application.ProcessMessages;
   //�������� �� XSD-�����
   try
    if not FSkipXSDValid then
     begin
      FXMLDocument.Schemas:=FXMLDOMSchema;
      FXMLDocument.Load(FItem.FileName);
      if FXMLDocument.ParseError.ErrorCode <> 0 then
       begin
        Note(FXMLDocument.ParseError.Reason, neError, i);
        FItem.State:=ifsError;
       end;
      FXMLDocument.Abort;
     end;
   except
    on E:Exception do
     begin
      Note(E.Message, neError, i);
      FItem.State:=ifsError;
     end;
   end;
   //��������� ���� XML
   if (FItem.State <> ifsError) then
    begin
     XMLDocument.Active:=False;
     XMLDocument.FileName:=FItem.FileName;
     try
      XMLDocument.Active:=True;
     except
      on E:Exception do
       begin
        Note(E.Message, neError, i);
        FItem.State:=ifsError;
       end;
     end;
    end;
   //���� ������� � �� ��
   if (FItem.State <> ifsError) and (XMLDocument.Active) then
    begin
     try
      //������� ������ ����������� ��������� ES_REC ���������
      Struct.ES_REC.Clear;
      //��� �����
      Struct.FILE_ES:=ExtractFileName(FItem.FileName);
      //�������� ���� ���������
      BufStr:=XMLDocument.DocumentElement.ChildNodes['���������'].ChildNodes['�������������'].Text;
      BufStr:=StringReplace(BufStr, '/', '.', [rfReplaceAll]);
      Struct.DATE_ES:=StrToDate(BufStr);
      //������ �����
      Struct.SIZE_ES:=GetFileSize(XMLDocument.FileName);
      //������ ��� �������� ���������
      Nodes:=XMLDocument.DocumentElement.ChildNodes['�����������'].ChildNodes;
      if Nodes <> nil then
       for n:= 0 to Nodes.Count-1 do
        begin
         //���� �� ���������, �� �� ���������� ���.
         if Nodes[n].IsTextElement then Continue;
         //��� ����� ������ �������
         Struct.ES_REC.Add(CheckES550PNode(Nodes[n]));
        end;
      //��, ���� ��� ������ �� �����
      XMLDocument.Active:=False;
      //�������� ���-�� �������
      Struct.RECNO_ES:=Struct.ES_REC.Count;
      //���� ���� ��������������� �������
      if DateTimePickerDate.Checked then
       Struct.DATE_KVIT:=DateTimePickerDate.Date
       //� ���� ���
      else Struct.DATE_KVIT:=Now;
      //���� ���� ��������������� �������
      if DateTimePickerTime.Checked then
       Struct.TIME_KVIT:=DateTimePickerTime.Time
       //� ���� ���
      else Struct.TIME_KVIT:=Now;
     except
      on E:Exception do
       begin
        Note('������ �������� XML: '+E.Message, neError, i);
        FItem.State:=ifsError;
       end;
     end;
     Application.ProcessMessages;
     //������������ ����������� � ����������
     try
      //���� ����������
      case ComboBoxSave.ItemIndex of
       0:SDir:=ExtractFileDir(FItem.FileName);
       1:SDir:=EditSaveTo.Text;
      end;
      //��������� ������������ ����������� �� �������. ������: UV_6661049239-666101001_CB_ES550P_20180302_001
      Struct.Output:=SDir+'\'+Format('UV_%s-%s_%s', [EditINN.Text, EditKPP.Text, ExtractFileName(FItem.FileName)]);
      //������ XML-���� � ���������
      if CreateKVIT(Struct) then
       begin
        //���� �����
        FItem.State:=ifsReady;
        FItem.Output:=Struct.Output;
       end
      else
       begin
        Note('����������� ������', neError, i);
        FItem.State:=ifsError;
       end;
     except
      on E:Exception do
       begin
        Note('������ ��� �������� ��� ���������� �����������: '+E.Message, neError, i);
        FItem.State:=ifsError;
       end;
     end;
    end;
   //���� ��������� �������� "� ���������", �� ������ ��������� ����������� ������
   if FItem.State = ifsProcces then
    begin
     Note('����������� ������', neError, i);
     FItem.State:=ifsError;
    end;
   //��������� ������� � ������
   FFiles[i]:=FItem;
   FFiles.UpdateTable;
   Application.ProcessMessages;
  end;
 //��������� ������� ��������
 FXMLDocument:=nil;
 FXMLDOMSchema:=nil;
 //������
 Note('��������� ���������', neInfo, -1);
end;

procedure TFormMain.ProcessFiles(List: TStrings);
var i:Integer;
    LastFound:string; //��������� ��������� ����-��������
    Found:Integer;    //����� ������� ���������

//����� ����� � ������
function FileAlreadyExists(FN:string):Boolean;
var j:Integer;
begin
 Result:=False;
 for j:= 0 to FFiles.Count-1 do
  if AnsiLowerCase(FFiles[j].FileName) = AnsiLowerCase(FN) then Exit(True);
end;

begin
 Found:=0;
 for i:= 0 to List.Count-1 do
  //���� ����� �� ������ ��� � ����� ������, �� ���������, � ��������� ������ ������� ���
  if not FileAlreadyExists(List[i]) then FFiles.Add(CreateInputFile(List[i])) else begin Inc(Found); LastFound:=List[i]; end;

 //������� ������������ ��������� ����������
 if Found > 0 then
  begin
   if Found = 1 then LastFound:='���� "'+LastFound+'" ��� ������� � ������ � �� ��� ��������.'
   else LastFound:='� ������ ��� ������ ��������� ����� ('+IntToStr(Found)+') ��� �� ���� ���������';
   Note(LastFound, neInfo, -1);
  end;
end;

procedure TFormMain.ComboBoxSaveChange(Sender: TObject);
begin
 EditSaveTo.Visible:=ComboBoxSave.ItemIndex = 1;
 SpeedButtonSaveTo.Visible:=ComboBoxSave.ItemIndex = 1;
end;

procedure TFormMain.DateTimePickerDateClick(Sender: TObject);
begin
 if not DateTimePickerDate.Checked then
  begin
   DateTimePickerDate.DateTime:=Now;
   DateTimePickerDate.Checked:=False;
  end;
end;

procedure TFormMain.DateTimePickerTimeClick(Sender: TObject);
begin
 if not DateTimePickerTime.Checked then
  begin
   DateTimePickerTime.DateTime:=Now;
   DateTimePickerTime.Checked:=False;
  end;
end;

procedure TFormMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 try
  SaveSettings;
 except

 end;
end;

procedure TFormMain.LoadSettings;
var Ini:TIniFile;
begin
 if not FileExists(ExtractFilePath(Application.ExeName)+'\config.ini') then
  begin
   Note('����� � ����������� �� �������. ����� ����������� ��������� �� ���������.', neWarning, -1);
   Note('�� �������� ������� ��������� � ����� ����� ����', neWarning, -1);
   //������������ ���� ������
   FileClose(FileCreate(ExtractFilePath(Application.ExeName)+'\config.ini'));
  end;

 Ini:=TIniFile.Create(ExtractFilePath(Application.ExeName)+'\config.ini');
 WindowState:=TWindowState(Ini.ReadInteger('General', 'WindowState', Ord(wsNormal)));
 if WindowState = wsNormal then
  begin
   ClientHeight:=Ini.ReadInteger('General', 'ClientHeight', ClientHeight);
   ClientWidth:=Ini.ReadInteger('General', 'ClientWidth', ClientWidth);
   Left:=Ini.ReadInteger('General', 'Left', Left);
   Top:=Ini.ReadInteger('General', 'Top', Top);
  end;
 EditINN.Text:=Ini.ReadString('General', 'INN', '');
 EditKPP.Text:=Ini.ReadString('General', 'KPP', '');
 EditOper.Text:=Ini.ReadString('General', 'OPER', '');
 EditTel.Text:=Ini.ReadString('General', 'TEL', '');
 ComboBoxSave.ItemIndex:=Ini.ReadInteger('General', 'SaveTo', 0);
 ComboBoxSaveChange(nil);
 EditSaveTo.Text:=Ini.ReadString('General', 'SaveToPath', '');
 DateTimePickerDate.DateTime:=Now;
 DateTimePickerDate.Checked:=False;
 DateTimePickerTime.DateTime:=Now;
 DateTimePickerTime.Checked:=False;
 Ini.Free;
end;

procedure TFormMain.SaveSettings;
var Ini:TIniFile;
begin
 Ini:=TIniFile.Create(ExtractFilePath(Application.ExeName)+'\config.ini');
 Ini.WriteInteger('General', 'WindowState', Ord(WindowState));
 if WindowState = wsNormal then
  begin
   Ini.WriteInteger('General', 'ClientHeight', ClientHeight);
   Ini.WriteInteger('General', 'ClientWidth', ClientWidth);
   Ini.WriteInteger('General', 'Left', Left);
   Ini.WriteInteger('General', 'Top', Top);
  end;
 Ini.WriteString('General', 'INN', EditINN.Text);
 Ini.WriteString('General', 'KPP', EditKPP.Text);
 Ini.WriteString('General', 'OPER', EditOper.Text);
 Ini.WriteString('General', 'TEL', EditTel.Text);
 Ini.WriteInteger('General', 'SaveTo', ComboBoxSave.ItemIndex);
 Ini.WriteString('General', 'SaveToPath', EditSaveTo.Text);
 Ini.Free;
end;

procedure TFormMain.ShowDragPanel;
begin
 PanelDrag.Show;
 PanelDrag.BringToFront;
end;

procedure TFormMain.IniXML;
var FXMLDocument:IXMLDOMDocument2;
    FXMLDOMSchema:IXMLDOMSchemaCollection2;
begin
 FExistsCLASS_DOMXML:=False;
 if not FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocument) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCache) as IXMLDOMSchemaCollection2;
   CLASS_DOMDocumentMy:=CLASS_DOMDocument;
   CLASS_XMLSchemaCacheMy:=CLASS_XMLSchemaCache;
   FExistsCLASS_DOMXML:=True;
   Note('��������������� MSXML '+CLASS_DOMDocument.ToString, neInfo, -1);
  except

  end;
 if not FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocument40) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCache40) as IXMLDOMSchemaCollection2;
   CLASS_DOMDocumentMy:=CLASS_DOMDocument40;
   CLASS_XMLSchemaCacheMy:=CLASS_XMLSchemaCache40;
   FExistsCLASS_DOMXML:=True;
   Note('��������������� MSXML 4.0', neInfo, -1);
  except

  end;
 if not FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocument60) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCache60) as IXMLDOMSchemaCollection2;
   CLASS_DOMDocumentMy:=CLASS_DOMDocument60;
   CLASS_XMLSchemaCacheMy:=CLASS_XMLSchemaCache60;
   FExistsCLASS_DOMXML:=True;
   Note('��������������� MSXML 6.0', neInfo, -1);
  except

  end;
 if not FExistsCLASS_DOMXML then
  begin
   Note('MSXML �� ���������������', neError, -1);
  end;
 FXMLDocument:=nil;
 FXMLDOMSchema:=nil;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
 FSkipXSDValidCheck:=True;
 FSkipXSDValid:=False;
 FExistsCLASS_DOMXML:=False;
 OleInitialize(nil);
 OleCheck(RegisterDragDrop(Handle, Self));
 //������
 FFiles:=TFiles.Create(TableExFiles);
 FNotes:=TNotes.Create(TableExNote);
 //���� ������
 with TableExFiles do
  begin
   AddColumn('�', 30);
   AddColumn('����', 260);
   AddColumn('������', 90);
   AddColumn('�������� ����', 450);
   Columns[AddColumn('', 100, True)].ShowButtonOnlySelect:=True;
   AddColumn('', 10);
  end;
 FFiles.UpdateTable;
 with TableExNote do
  begin
   AddColumn('', 30);
   AddColumn('�������', 130);
   AddColumn('����� ���������', 300);
  end;
 FNotes.UpdateTable;
 //
 Note('������ ����! ������ �����...', neInfo, -1);
 //������������� MSXML
 IniXML;
 //�������� ��������
 try
  LoadSettings;
 except
  Note('������ ��� �������� ���������', neError, -1);
 end;
 //���������
 Note('�������� ����� � ������. ������ "�������� �����..." ��� ������ ���������� ���� � ����', neInfo, -1);
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
 RevokeDragDrop(Handle);
 OleUninitialize;
end;

procedure TFormMain.HideDragPanel;
begin
 PanelDrag.Hide;
end;

procedure TFormMain.SpeedButtonSaveToClick(Sender: TObject);
var SDir:string;
begin
 if AdvSelectDirectory('', '', SDir, True, False, True) then EditSaveTo.Text:=SDir;
end;

procedure TFormMain.TableExFilesDblClick(Sender: TObject);
begin
 if not IndexInList(TableExFiles.ItemIndex, FFiles.Count) then Exit;
 //���� ���� �������� ����, �� ��������� ���
 if FileExists(FFiles[TableExFiles.ItemIndex].Output) then
  begin
   OpenFolderAndSelectFile(FFiles[TableExFiles.ItemIndex].Output);
  end
 else
  //���� ��� � ���������� �������� ����, �� ��������� ���
  if FileExists(FFiles[TableExFiles.ItemIndex].FileName) then
   begin
    OpenFolderAndSelectFile(FFiles[TableExFiles.ItemIndex].FileName);
   end;
end;

procedure TFormMain.TableExFilesGetData(FCol, FRow: Integer; var Value: string);
begin
 Value:='';
 if not IndexInList(FRow, FFiles.Count) then Exit;
 case FCol of
  0:Value:=IntToStr(FRow+1);
  1:Value:=ExtractFileName(FFiles[FRow].FileName);
  2:Value:=InputFileStateStr(FFiles[FRow].State);
  3:Value:=ExtractFileName(FFiles[FRow].Output);
  4:Value:='�������';
 end;
end;

procedure TFormMain.TableExFilesItemColClick(Sender: TObject; MouseButton: TMouseButton; const Index: Integer);
begin
 if not IndexInList(TableExFiles.ItemIndex, FFiles.Count) then Exit;
 case Index of
  4:FFiles.Delete(TableExFiles.ItemIndex);
 end;
end;

procedure TFormMain.TableExNoteDblClick(Sender: TObject);
begin
 if not IndexInList(TableExNote.ItemIndex, FNotes.Count) then Exit;
 if not IndexInList(FNotes[TableExNote.ItemIndex].ItemIndex, FFiles.Count) then Exit;
 TableExFiles.ItemIndex:=FNotes[TableExNote.ItemIndex].ItemIndex;
end;

procedure TFormMain.TableExNoteDrawCellData(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
 if not IndexInList(ARow, FNotes.Count) then Exit;
 if ACol <> 0 then Exit;
 with TableExNote.Canvas do
  begin
   case FNotes[ARow].Event of
    neInfo:ImageList.Draw(TableExNote.Canvas, Rect.Left, Rect.Top, 1);
    neWarning:ImageList.Draw(TableExNote.Canvas, Rect.Left, Rect.Top, 2);
    neError:ImageList.Draw(TableExNote.Canvas, Rect.Left, Rect.Top, 3);
   end;
  end;
end;

procedure TFormMain.TableExNoteGetData(FCol, FRow: Integer; var Value: string);
begin
 Value:='';
 if not IndexInList(FRow, FNotes.Count) then Exit;
 case FCol of
  1:Value:=EventNoteStr(FNotes[FRow].Event);
  2:Value:=FNotes[FRow].Text;
 end;
end;

procedure TFormMain.TimerUpdateDTTimer(Sender: TObject);
begin
 if not DateTimePickerDate.Checked then
  begin
   DateTimePickerDate.DateTime:=Now;
   DateTimePickerDate.Checked:=False;
  end;
 if not DateTimePickerTime.Checked then
  begin
   DateTimePickerTime.DateTime:=Now;
   DateTimePickerTime.Checked:=False;
  end;
end;

end.
