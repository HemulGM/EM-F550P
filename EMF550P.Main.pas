unit EMF550P.Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Xml.XMLIntf, Vcl.StdCtrls,
  Xml.XMLDoc, System.Generics.Collections, TableDraw, Vcl.Buttons,
  Vcl.ExtCtrls, Vcl.ComCtrls, ShellAPI, ShlObj, ActiveX, System.Win.ComObj, IniFiles,
  LabelButton, Vcl.ImgList, System.ImageList, Vcl.Grids, Xml.xmldom, MSXML;

type

  //Элементы Раздел2 уведомления
  TESItem = record
   IdInfoOR:string;
   RES_ES:Integer;
   KOD_ERR:Integer;
   NAM_ERR:string;
  end;
  TESItems = TList<TESItem>;

  //Структура уведомления
  TES550PStruct = record
   Output:string;                //Конечный файл
   IDNOR:string;                 //
   FILE_ES:string;               //
   SIZE_ES:Cardinal;             //
   DATE_ES:TDate;                //
   RECNO_ES:Integer;             //
   ES_REC:TESItems;              //Элементы Раздел2 уведомления
   DATE_KVIT:TDate;              //
   TIME_KVIT:TTime;              //
   OPER:string;                  //
   TEL_OPER:string;              //
  end;

  //Состояние элемента в таблице исходных файлов
  TInputFileState = (ifsWait, ifsProcces, ifsError, ifsReady);

  //Список исходных файлов
  TInputFile = record
   FileName:TFileName;     //Имя файла
   State:TInputFileState;  //Состояние
   Output:TFileName;       //Конечный файл
  end;
  TFiles = TTableData<TInputFile>;

  //Тип инф. собщения
  TNoteEvent = (neInfo, neWarning, neError);

  //Инф. сообщения
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
    FFiles:TFiles;                                                              //Список исходных XML-файлов
    FNotes:TNotes;                                                              //Сообщения для пользователя
    FSkipXSDValidCheck:Boolean;                                                 //Выдавать вопрос при ошибке в проверке по XSD-схеме
    FSkipXSDValid:Boolean;                                                      //Проверять файлы по XSD-схеме
    FExistsCLASS_DOMXML:Boolean;                                                //Установлен класс MSXML
    CLASS_DOMDocumentMy:TGUID;                                                  //Классы MSXML
    CLASS_XMLSchemaCacheMy:TGUID;
    procedure Process;                                                          //Сформировать уведомления
    function CheckES550PNode(Data: IXMLNode): TESItem;                          //Проверка ветки "Раздел2"
    function CreateKVIT(Struct: TES550PStruct): Boolean;                        //Сформировать и сохранить XML-файл уведомления
    procedure SaveSettings;                                                     //Сохранить настройки
    procedure LoadSettings;                                                     //Загрузить
    procedure ProcessFiles(List:TStrings);                                      //Добавление файлов в список
    procedure ShowDragPanel;                                                    //Показать панель дропа
    procedure HideDragPanel;                                                    //Скрыть
    function Note(Text: string; Event: TNoteEvent; ItemIndex:Integer):Boolean;  //Добавить в область информации запись
    procedure IniXML;                                                           //Инициализация класса MSXML
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
  ifsWait: Result:='Ожидание';
  ifsProcces: Result:='Обработка';
  ifsReady: Result:='Готов';
  ifsError: Result:='Ошибка';
 end;
end;

function EventNoteStr(Value:TNoteEvent):string;
begin
 case Value of
  neInfo: Result:='Информация';
  neWarning: Result:='Предупреждение';
  neError: Result:='Ошибка';
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
 Result:=ESItem(Data.ChildNodes['НомерЗаписи'].Text, 0);
end;

function TFormMain.CreateKVIT(Struct:TES550PStruct):Boolean;
var Root, Recs, Rec:IXMLNode;
    XML:TXMLDocument;
    i:Integer;
    LSep:Char;
begin
 Result:=False;
 //Требуют "/" для разделения даты. Делаем.
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
     //Корень файла
     Root:=XML.AddChild('KVIT');
     //Добавим начальные данные
     Root.AddChild('IDNOR').Text:=Struct.IDNOR;
     Root.AddChild('ES').Text:=Struct.FILE_ES;
     Root.AddChild('SIZE_ES').Text:=IntToStr(Struct.SIZE_ES);
     Root.AddChild('DATE_ES').Text:=FormatDateTime('ddddd', Struct.DATE_ES);
     //Заполним данные записей
     Recs:=Root.AddChild('RECNO_ES');
     Recs.Attributes['nReс']:=IntToStr(Struct.RECNO_ES);
     for i:= 0 to Struct.ES_REC.Count-1 do
      begin
       Rec:=Recs.AddChild('ES_REC');
       Rec.Attributes['IdInfoOR']:=Struct.ES_REC[i].IdInfoOR;
       Rec.AddChild('RES_ES').Text:=IntToStr(Struct.ES_REC[i].RES_ES);
      end;
     //Добавим последние данные
     Root.AddChild('DATE_KVIT').Text:=FormatDateTime('ddddd', Struct.DATE_KVIT);
     Root.AddChild('TIME_KVIT').Text:=FormatDateTime('HH:NN:SS', Struct.TIME_KVIT);
     Root.AddChild('OPER').Text:=Struct.OPER;
     Root.AddChild('TEL_OPER').Text:=Struct.TEL_OPER;
     Result:=True;
     //Сохраняем файл
     XML.SaveToFile(Struct.Output);
     XML.Active:=False;
    end;
  end;
 except
  on E:Exception do Note(E.Message, neError, -1);
 end;
 //Вернём разделитель
 FormatSettings.DateSeparator:=LSep;
end;

function TFormMain.DragEnter(const dataObj:IDataObject; grfKeyState:Longint; pt:TPoint; var dwEffect:Longint):HResult;
var FmtEtc:TFormatEtc;
    Medium:TStgMedium;
begin
 //Перед дропом, проверим. Если не файлы, то запрет на дроп
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
 //Структура дропа файлов
 FmtEtc.cfFormat:=CF_HDROP;
 FmtEtc.ptd:=nil;
 FmtEtc.dwAspect:=DVASPECT_CONTENT;
 FmtEtc.lindex:=-1;
 FmtEtc.tymed:=TYMED_HGLOBAL;
 //Если структура нужная нам
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
       //Только XML-файлы
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
 MessageBox(Application.Handle, 'Программа для формирования уведомления о принятии (непринятии) электронного сообщения кредитной организацией (филиалом кредитной организации), НФО (филиалом НФО) предусмотренного Положением Банка России № 550-П'+#13+#10+''+#13+#10+'Разработчик: Геннадий Малинин 2018 (с)'+#13+#10+''+#13+#10+'Инструкция:'+#13+#10+'1. Для начала работы необходимо настроить реквизиты в левой части окна программы.'+
  #13#10+'2. После настройки необходимо добавить файлы в таблицу. Это можно сделать двумя способами:'+#13+#10+' А) Нажать кнопку "Добавить файл..." и выбрать необходимые файлы'+#13+#10+' Б) Выделить необходимые файлы в проводнике Windows и перетащить в окно программы'+#13+#10+'3. После успешного пополнения списка необходимо указать куда сохранять конечные файлы (внизу окна программы)'+#13+#10+'4. Нажать кнопку "Сформировать"'+#13+#10+'5. Готово'+#13+#10+''+#13+#10+'Файлы уже имеющиеся в списке не могут быть добавлены повторно и будут игнорироваться', 'Информация', MB_ICONINFORMATION or MB_OK);
end;

procedure TFormMain.ButtonProcessClick(Sender: TObject);
begin
 //Проверим реквизиты
 if EditINN.Text = '' then if Note('Укажите ИНН организации', neWarning, -1) then Exit;
 if EditKPP.Text = '' then if Note('Укажите КПП организации', neWarning, -1) then Exit;
 if EditOper.Text = '' then if Note('Укажите оператора (ФИО)', neWarning, -1) then Exit;
 if EditTel.Text = '' then if Note('Укажите контактный телефон оператора', neWarning, -1) then Exit;
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
 //Если таблица пуста
 if FFiles.Count <= 0 then
  begin
   Note('Сначала добавьте файлы ЭС в список', neInfo, -1);
   Exit;
  end;
 //Если выбран режим сохранения в каталог, то проверяем его существование
 if ComboBoxSave.ItemIndex = 1 then
  if not DirectoryExists(EditSaveTo.Text) then
   begin
    Note('Каталог сохранения не указан или указан неверно!', neWarning, -1);
    Exit;
   end;
 Note('Обработка запущена...', neInfo, -1);
 //Инициализация проверки документов
 if FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocumentMy) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCacheMy) as IXMLDOMSchemaCollection2;
  except
   FXMLDocument:=nil;
  end;
 //Если что-то не так, то спрашиваем пользователя. Пропускать без проверки или закончить работу.
 if FXMLDocument = nil then
  begin
   if FSkipXSDValidCheck then
    begin
     FSkipXSDValidCheck:=True;
     if MessageBox(Application.Handle, 'Программе не удалось инициализировать проверку файла по XSD-схеме. Если продолжить, проверка будет пропущена.'+#13+#10+'Прервать выполнение операции?', 'Внимание', MB_ICONEXCLAMATION or MB_YESNO or MB_DEFBUTTON1) = ID_NO then
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
 //Создадим структуру и заполним реквизиты
 Struct.ES_REC:=TESItems.Create;
 //ИНН и КПП
 Struct.IDNOR:=Format('%s_%s', [EditINN.Text, EditKPP.Text]);
 //Оператор
 Struct.OPER:=EditOper.Text;
 //Конт. тел.
 Struct.TEL_OPER:=EditTel.Text;
 //По списку
 for i:= 0 to FFiles.Count-1 do
  begin
   FItem:=FFiles[i];
   //Элемент в обработке
   FItem.State:=ifsProcces;
   //Обновим
   FFiles[i]:=FItem;
   FFiles.UpdateTable;
   Application.ProcessMessages;
   //Проверка по XSD-схеме
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
   //Загружаем файл XML
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
   //Если открыли и всё ок
   if (FItem.State <> ifsError) and (XMLDocument.Active) then
    begin
     try
      //Очистим список проверяемых элементов ES_REC структуры
      Struct.ES_REC.Clear;
      //Имя файла
      Struct.FILE_ES:=ExtractFileName(FItem.FileName);
      //Забираем дату сообщения
      BufStr:=XMLDocument.DocumentElement.ChildNodes['СлужЧасть'].ChildNodes['ДатаСообщения'].Text;
      BufStr:=StringReplace(BufStr, '/', '.', [rfReplaceAll]);
      Struct.DATE_ES:=StrToDate(BufStr);
      //Размер файла
      Struct.SIZE_ES:=GetFileSize(XMLDocument.FileName);
      //Читаем все элементы сообщения
      Nodes:=XMLDocument.DocumentElement.ChildNodes['ИнформЧасть'].ChildNodes;
      if Nodes <> nil then
       for n:= 0 to Nodes.Count-1 do
        begin
         //Если он текстовый, то мы пропускаем его.
         if Nodes[n].IsTextElement then Continue;
         //Нам нужны только разделы
         Struct.ES_REC.Add(CheckES550PNode(Nodes[n]));
        end;
      //Всё, файл нам больше не нужен
      XMLDocument.Active:=False;
      //Сохраним кол-во записей
      Struct.RECNO_ES:=Struct.ES_REC.Count;
      //Если дата устанавливается вручную
      if DateTimePickerDate.Checked then
       Struct.DATE_KVIT:=DateTimePickerDate.Date
       //И если нет
      else Struct.DATE_KVIT:=Now;
      //Если дата устанавливается вручную
      if DateTimePickerTime.Checked then
       Struct.TIME_KVIT:=DateTimePickerTime.Time
       //И если нет
      else Struct.TIME_KVIT:=Now;
     except
      on E:Exception do
       begin
        Note('Ошибка парсинга XML: '+E.Message, neError, i);
        FItem.State:=ifsError;
       end;
     end;
     Application.ProcessMessages;
     //Формирование уведомления и сохранение
     try
      //Путь сохранения
      case ComboBoxSave.ItemIndex of
       0:SDir:=ExtractFileDir(FItem.FileName);
       1:SDir:=EditSaveTo.Text;
      end;
      //Формируем наименование уведомления по правилу. Пример: UV_6661049239-666101001_CB_ES550P_20180302_001
      Struct.Output:=SDir+'\'+Format('UV_%s-%s_%s', [EditINN.Text, EditKPP.Text, ExtractFileName(FItem.FileName)]);
      //Создаём XML-файл и сохраняем
      if CreateKVIT(Struct) then
       begin
        //Файл готов
        FItem.State:=ifsReady;
        FItem.Output:=Struct.Output;
       end
      else
       begin
        Note('Неизвестная ошибка', neError, i);
        FItem.State:=ifsError;
       end;
     except
      on E:Exception do
       begin
        Note('Ошибка при создании или сохранении уведомления: '+E.Message, neError, i);
        FItem.State:=ifsError;
       end;
     end;
    end;
   //Если состояние элемента "В обработке", то значит произошла неизвестная ошибка
   if FItem.State = ifsProcces then
    begin
     Note('Неизвестная ошибка', neError, i);
     FItem.State:=ifsError;
    end;
   //Обновляем элемент в списке
   FFiles[i]:=FItem;
   FFiles.UpdateTable;
   Application.ProcessMessages;
  end;
 //Освободим объекты проверки
 FXMLDocument:=nil;
 FXMLDOMSchema:=nil;
 //Готово
 Note('Обработка завершена', neInfo, -1);
end;

procedure TFormMain.ProcessFiles(List: TStrings);
var i:Integer;
    LastFound:string; //Последний найденный файл-дубликат
    Found:Integer;    //Будем считать дубликаты

//Поиск файла в списке
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
  //Если файла из списка нет в нашем списке, то добавляем, в противном случае считаем его
  if not FileAlreadyExists(List[i]) then FFiles.Add(CreateInputFile(List[i])) else begin Inc(Found); LastFound:=List[i]; end;

 //Покажем пользователю результат добавления
 if Found > 0 then
  begin
   if Found = 1 then LastFound:='Файл "'+LastFound+'" уже имеется в списке и не был добавлен.'
   else LastFound:='В списке уже имются некоторые файлы ('+IntToStr(Found)+') они не были добавлены';
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
   Note('Файла с настройками не найдено. Будут установлены настройки по умолчанию.', neWarning, -1);
   Note('Не забудьте указать реквизиты в левой части окна', neWarning, -1);
   //Пересоздадим файл конфиг
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
   Note('Инициализирован MSXML '+CLASS_DOMDocument.ToString, neInfo, -1);
  except

  end;
 if not FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocument40) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCache40) as IXMLDOMSchemaCollection2;
   CLASS_DOMDocumentMy:=CLASS_DOMDocument40;
   CLASS_XMLSchemaCacheMy:=CLASS_XMLSchemaCache40;
   FExistsCLASS_DOMXML:=True;
   Note('Инициализирован MSXML 4.0', neInfo, -1);
  except

  end;
 if not FExistsCLASS_DOMXML then
  try
   FXMLDocument:=CreateComObject(CLASS_DOMDocument60) as IXMLDomDocument2;
   FXMLDOMSchema:=CreateComObject(CLASS_XMLSchemaCache60) as IXMLDOMSchemaCollection2;
   CLASS_DOMDocumentMy:=CLASS_DOMDocument60;
   CLASS_XMLSchemaCacheMy:=CLASS_XMLSchemaCache60;
   FExistsCLASS_DOMXML:=True;
   Note('Инициализирован MSXML 6.0', neInfo, -1);
  except

  end;
 if not FExistsCLASS_DOMXML then
  begin
   Note('MSXML не инициализирован', neError, -1);
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
 //Списки
 FFiles:=TFiles.Create(TableExFiles);
 FNotes:=TNotes.Create(TableExNote);
 //Поля таблиц
 with TableExFiles do
  begin
   AddColumn('№', 30);
   AddColumn('Файл', 260);
   AddColumn('Статус', 90);
   AddColumn('Конечный файл', 450);
   Columns[AddColumn('', 100, True)].ShowButtonOnlySelect:=True;
   AddColumn('', 10);
  end;
 FFiles.UpdateTable;
 with TableExNote do
  begin
   AddColumn('', 30);
   AddColumn('Событие', 130);
   AddColumn('Текст сообщения', 300);
  end;
 FNotes.UpdateTable;
 //
 Note('Добрый день! Сейчас начнём...', neInfo, -1);
 //Инициализация MSXML
 IniXML;
 //Загрузка настроек
 try
  LoadSettings;
 except
  Note('Ошибка при загрузке настройке', neError, -1);
 end;
 //Подсказки
 Note('Добавьте файлы в список. Кнопка "Добавить файлы..." или просто перетащите файл в окно', neInfo, -1);
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
 //Если есть конечный файл, то открываем его
 if FileExists(FFiles[TableExFiles.ItemIndex].Output) then
  begin
   OpenFolderAndSelectFile(FFiles[TableExFiles.ItemIndex].Output);
  end
 else
  //Если нет и существует исходный файл, то открываем его
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
  4:Value:='Удалить';
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
