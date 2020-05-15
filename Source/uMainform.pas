{***************************************************************************}
{                                                                           }
{           Visualizzatore Fattura Elettronica "LITE"                       }
{                                                                           }
{           Copyright (C) 2018 Giancarlo Oneglio                            }
{                                                                           }
{           giancarlo.oneglio@gmail.com                                     }
{                                                                           }
{                                                                           }
{***************************************************************************}
{                                                                           }
{  This file is part of VisualizzatoreFatturaElettronicaLITE                }
{                                                                           }
{  Licensed under the GNU Lesser General Public License, Version 3;         }
{  you may not use this file except in compliance with the License.         }
{                                                                           }
{  VisualizzatoreFatturaElettronicaLITE is free software:                   }
{  you can redistribute it and/or modify                                    }
{  it under the terms of the GNU Lesser General Public License as published }
{  by the Free Software Foundation, either version 3 of the License, or     }
{  (at your option) any later version.                                      }
{                                                                           }
{  VisualizzatoreFatturaElettronicaLITE is distributed in the hope          }
{  that it will be useful, but WITHOUT ANY WARRANTY;                        }
{  without even the implied warranty of                                     }
{  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the            }
{  GNU Lesser General Public License for more details.                      }
{                                                                           }
{  You should have received a copy of the GNU Lesser General Public License }
{  along with VisualizzatoreFatturaElettronicaLITE                          }
{  If not, see <http://www.gnu.org/licenses/>.                              }
{                                                                           }
{***************************************************************************}

unit uMainform;

interface

uses
    Winapi.Windows,
    Winapi.Messages,

    System.SysUtils,
    System.Variants,
    System.Classes,
    System.IOUtils,
    System.Types,

    Vcl.Controls,
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.StdCtrls,
    Vcl.Buttons,
    Vcl.ComCtrls,
    Vcl.OleCtrls,
    Vcl.Menus,
    Vcl.Imaging.jpeg,
    Vcl.AppEvnts,
    Vcl.Imaging.pngimage,
    Vcl.ExtCtrls,
    Vcl.FileCtrl,
    Vcl.Samples.Spin,

    Registry,
    IniFiles,
    SHDocVw,
    ActiveX,
    Xml.xmldom,
    Xml.XMLIntf,
    Xml.Win.msxmldom,
    Xml.XMLDoc

    ;

type
  TAllegato = record
    FileName: string;
    FileType: string;
    Compressione: string;
    Data: string;
    Button: TButton;
    procedure DumpAndOpen;
    procedure Clear;
  end;

    TMainform = class(TForm)
        XMLDocument1: TXMLDocument;
        WebBrowserFEpreview: TWebBrowser;
        StatusBar1: TStatusBar;
        Splitter1: TSplitter;
        PanelSX: TPanel;
        GroupBox3: TGroupBox;
        BitBtnShowXML: TBitBtn;
        BitBtnPrintPreview: TBitBtn;
        BitBtnPrint: TBitBtn;
        ComboBoxStile: TComboBox;
        Label1: TLabel;
        PopupMenuMainform: TPopupMenu;
        Informazioni1: TMenuItem;
        N1: TMenuItem;
        OpenDialogFeFile: TOpenDialog;
        SpeedButton1: TSpeedButton;
        gbAllegati: TGroupBox;
        procedure BitBtnPrintPreviewClick(Sender: TObject);
        procedure FormCreate(Sender: TObject);
        procedure AllegatoButtonClick(Sender: TObject);
        procedure BitBtnPrintClick(Sender: TObject);
        procedure WebBrowserFEpreviewDocumentComplete(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
        procedure ApplicationEvents1Message(var Msg: tagMSG; var Handled: Boolean);
        procedure Informazioni1Click(Sender: TObject);
        procedure BitBtnShowXMLClick(Sender: TObject);
        procedure SpeedButton1Click(Sender: TObject);
        procedure ComboBoxStileSelect(Sender: TObject);
        procedure FormDestroy(Sender: TObject);
    private
        { Private declarations }
        currDir,config_file: string;
        dirIn,dirOut: string;
        dirInSearch, dirOutSearch: Boolean;
        file_attuale:string;
        FAllegati: TArray<TAllegato>;
        procedure ApriFatturaXML(XMLDoc: string);
        function GetFontSize: integer;
        procedure SetFontSize(Size: integer);
        procedure SetOpticalZoom(Value: integer);
        procedure WMDropFiles(var Msg: TMessage); message WM_DROPFILES;
        procedure ParseAllegati(const AXMLDoc: TXMLDocument);
    public
        { Public declarations }
    end;

    function CoInternetSetFeatureEnabled(FeatureEntry: DWORD; dwFlags: DWORD; fEnable: BOOL): HRESULT; stdcall; external 'urlmon.dll';

var
    Mainform: TMainform;

const
    SET_FEATURE_ON_PROCESS = $00000002;
    FEATURE_DISABLE_NAVIGATION_SOUNDS = 21;
implementation

{$R *.dfm}

uses ShellAPI, uInformazioni, NetEncoding;

procedure WB_LoadHTML(WebBrowser: TWebBrowser; HTMLCode: string);
var
    sl: TStringList;
    ms: TMemoryStream;
begin
    WebBrowser.Navigate('about:blank');
    while WebBrowser.ReadyState < READYSTATE_INTERACTIVE do
    begin
        Application.ProcessMessages;
    end;

    if Assigned(WebBrowser.Document) then
    begin
        sl := TStringList.Create;
        try
            ms := TMemoryStream.Create;
            try
                sl.Text := HTMLCode;
                sl.SaveToStream(ms);
                ms.Seek(0, 0);
                (WebBrowser.Document as IPersistStreamInit).Load(TStreamAdapter.Create(ms));
            finally
                ms.Free;
            end;
        finally
            sl.Free;
        end;
    end;
end;

procedure TMainform.AllegatoButtonClick(Sender: TObject);
var
  LButton: TButton;
  LAllegato: TAllegato;
begin
  LButton := Sender as TButton;
  LAllegato := FAllegati[LButton.Tag];

  LAllegato.DumpAndOpen;
end;

procedure TMainform.ApplicationEvents1Message(var Msg: tagMSG; var Handled: Boolean);
begin
  if (Msg.message=WM_RBUTTONDOWN) and IsChild(WebBrowserFEpreview.Handle,Msg.hwnd) then
   begin
       Handled:=true;
   end;
end;

procedure TMainform.ApriFatturaXML(XMLDoc: string);
var
  LStyleSheetNode: IXMLNode;
  LTempFileName, LStyleSheet: string;
  LInvoiceXML: string;
  LStartIndex: Integer;
  LEndIndex: Integer;
  LNodeIndex: Integer;
begin
  file_attuale := XMLDoc;
  try
    LStyleSheet := 'fe_templates\' + ComboBoxStile.Items[ComboBoxStile.ItemIndex];
    LInvoiceXML := TFile.ReadAllText(XMLDoc);

    repeat
      LStartIndex := LInvoiceXML.IndexOf('<?xml-stylesheet');
      if LStartIndex <> -1 then
      begin
        LEndIndex := LInvoiceXML.IndexOf('>', LStartIndex) + 1;
        Delete(LInvoiceXML, LStartIndex+1, LEndIndex - LStartIndex);
      end;
    until LStartIndex = -1;

    XMLDocument1.Active := false;
    XMLDocument1.Xml.Text := LInvoiceXML;
    XMLDocument1.Active := true;

    LStyleSheetNode := XMLDocument1.CreateNode('xml-stylesheet', ntProcessingInstr, 'type="text/xsl" href="' + currDir+LStyleSheet + '"');

    LNodeIndex := 1;
    if LInvoiceXML.IndexOf('<?xml version=') = -1 then
      LNodeIndex := 0;

    XMLDocument1.ChildNodes.Insert(LNodeIndex, LStyleSheetNode);
    LTempFileName := TPath.GetTempFileName + '.xml';
    XMLDocument1.Xml.SaveToFile(LTempFileName);
    WebBrowserFEpreview.Navigate(LTempFileName);
    ParseAllegati(XMLDocument1);
    StatusBar1.Panels[0].Text := ExtractFileName(XMLDoc);
  except
    on e:Exception do
    begin
      ShowMessage('Errore durante apertura file: ' + XMLDoc + sLineBreak
        + '[' + e.Message + ']');
    end;
  end;
  LInvoiceXML:='';
  XMLDocument1.Active := false;
end;

procedure TMainform.BitBtnPrintPreviewClick(Sender: TObject);
var
    vaIn, vOut: OleVariant;
begin
    WebBrowserFEpreview.ControlInterface.ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, vaIn, vOut);
end;

procedure TMainform.BitBtnShowXMLClick(Sender: TObject);
begin
  if OpenDialogFeFile.Execute then ApriFatturaXML(OpenDialogFeFile.FileName);
end;

procedure TMainform.ComboBoxStileSelect(Sender: TObject);
begin
    if (file_attuale<>'') then ApriFatturaXML(file_attuale);
end;

procedure TMainform.BitBtnPrintClick(Sender: TObject);
var
    vIn, vOut: OleVariant;
begin
    WebBrowserFEpreview.ControlInterface.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, vIn, vOut)
end;

procedure TMainform.FormCreate(Sender: TObject);
var
    RegistryEntry: TRegistry;
    template,fn_exe, fn_xsl: string;
    IniCfg:TIniFile;
begin
    FAllegati := [];

    DragAcceptFiles(Handle, True);
    CoInternetSetFeatureEnabled(FEATURE_DISABLE_NAVIGATION_SOUNDS,SET_FEATURE_ON_PROCESS,True);
    currDir:=ExtractFilePath(ParamStr(0));
    config_file:=ChangeFileExt(ParamStr(0),'.cfg');

    ComboBoxStile.Clear;
    for fn_xsl in TDirectory.GetFiles(currDir+'fe_templates')  do
    begin
        if LowerCase(ExtractFileExt(fn_xsl))='.xsl' then
        begin
            ComboBoxStile.Items.Add(ExtractFileName(fn_xsl));
        end;
    end;

    fn_exe := ExtractFileName(ParamStr(0));

    Informazioni:=TInformazioni.Create(Application);

    dirIn:=currDir+'fe_fornitori';
    dirInSearch:=TRUE;
    dirOut:=currDir+'fe_clienti';
    dirOutSearch:=TRUE;

	//utilizzare IEFIX.reg per aggiungere nel registro la compatibilità del twebbrowser con la versione IE presente
	//oppure gestire il controllo da programma
	
	//    RegistryEntry := TRegistry.Create;
	//    RegistryEntry.RootKey := HKEY_CURRENT_USER;
	//    RegistryEntry.OpenKeyReadOnly('Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION');
	//    if RegistryEntry.ValueExists(fn_exe) = false then
	//    begin
	//        ShowMessage('compatibilità IE di questo eseguibile non definita nel registro');
	//        WebBrowserFEpreview.Navigate('http://www.mybrowserinfo.com/');
	//    end;
	//    RegistryEntry.Free;

    IniCfg:=TIniFile.Create(currDir+'config.ini');
    template:=IniCfg.ReadString('config','template','');
    IniCfg.Free;

    if template<>'' then
        ComboBoxStile.ItemIndex:=ComboBoxStile.Items.IndexOf(template)
    else
        ComboBoxStile.ItemIndex:=0;

    if (ParamCount > 0) {and SameText(ExtractFileExt(ParamStr(1)), '.xml')} then
      ApriFatturaXML(ParamStr(1));
end;

procedure TMainform.FormDestroy(Sender: TObject);
begin
    DragAcceptFiles(Handle, False);
end;

procedure TMainform.SetOpticalZoom(Value : integer);
var vaIn, vaOut : OleVariant;
begin
  vaIn := null;
  vaOut := null;
  WebBrowserFEpreview.ExecWB(OLECMDID_OPTICAL_GETZOOMRANGE,OLECMDEXECOPT_DONTPROMPTUSER,vaIn,vaOut);
  if Value < LoWord(DWORD(vaOut)) then
      vaIn := LoWord(DWORD(vaOut))
    else
      if Value > HiWord(DWORD(vaOut)) then
        vaIn := HiWord(DWORD(vaOut))
      else
        vaIn := Value;
  WebBrowserFEpreview.ExecWB(OLECMDID_OPTICAL_ZOOM,OLECMDEXECOPT_DONTPROMPTUSER,vaIn,vaOut);
end;

procedure TMainform.SpeedButton1Click(Sender: TObject);
var
    IniCfg:TIniFile;
begin
    try
        IniCfg:=TIniFile.Create(currDir+'config.ini');
        IniCfg.WriteString('config','template',ComboBoxStile.Items[ComboBoxStile.ItemIndex]);
        IniCfg.Free;
        ShowMessage('Preferenza salvata');
    except
        ShowMessage('Errore durante la scrittura');
    end;
end;

procedure TMainform.WebBrowserFEpreviewDocumentComplete(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
begin
    SetOpticalZoom(100);
end;

procedure TMainform.WMDropFiles(var Msg: TMessage);
var
  LHandleDrop: THandle;
  LCount: Integer;
  LNameLength: Integer;
  LIndex: Integer;
  LFileName: string;

begin
  LHandleDrop := Msg.wParam;
  LCount := DragQueryFile(LHandleDrop, $FFFFFFFF, nil, 0);

  for LIndex := 0 to LCount-1 do begin
    LNameLength := DragQueryFile(LHandleDrop, LIndex, nil, 0);
    SetLength(LFileName, LNameLength);
    DragQueryFile(LHandleDrop, LIndex, PWideChar(LFileName), LNameLength+1);

    ApriFatturaXML(LFileName);
  end;

  DragFinish(LHandleDrop);
end;

procedure TMainform.SetFontSize(Size: integer);
var
  vIn, vOut: Olevariant;
begin
    vIn := Size;
    WebBrowserFEpreview.ControlInterface.ExecWB(OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, vIn, vOut);
end;

function TMainform.GetFontSize: integer;
var
  vIn, vOut: Olevariant;
begin
    result := 0;
    vIn := null;
    WebBrowserFEpreview.ControlInterface.ExecWB(OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, vIn, vOut);
    result := vOut;
end;

procedure TMainform.Informazioni1Click(Sender: TObject);
begin
    Informazioni:=TInformazioni.Create(Application);
    Informazioni.ShowModal;
    Informazioni.Free;
end;

procedure TMainform.ParseAllegati(const AXMLDoc: TXMLDocument);
var
  LAllegati: IDOMNodeList;
  LAllegato: TAllegato;
  LButton: TButton;
  LAllegatoNode: IDOMNode;
  LAllegatoElement: IDOMElement;
  LList: IDOMNodeList;
  LIndex: Integer;
  LNomeAttachment: string;
  LFormatoAttachment: string;
  LAlgoritmoCompressione: string;
  LData: string;
begin
  LAllegati := AXMLDoc.DOMDocument.getElementsByTagName('Allegati');
  for LAllegato in  FAllegati do
    LAllegato.Button.Free;
  FAllegati := [];

  for LIndex := 0 to LAllegati.length - 1 do
  begin
    LAllegato.Clear;
    LAllegatoNode := LAllegati.item[LIndex];
    LAllegatoElement := LAllegatoNode as IDOMElement;
    if Assigned(LAllegatoElement) then
    begin
      LList := LAllegatoElement.getElementsByTagName('NomeAttachment');
      LNomeAttachment := '';
      if (LList.length = 1) and (LList.item[0].hasChildNodes) then
        LNomeAttachment := LList.item[0].firstChild.nodeValue;

      LList := LAllegatoElement.getElementsByTagName('FormatoAttachment');
      LFormatoAttachment := '';
      if (LList.length = 1) and (LList.item[0].hasChildNodes) then
        LFormatoAttachment := LList.item[0].firstChild.nodeValue;

      LList := LAllegatoElement.getElementsByTagName('AlgoritmoCompressione');
      LAlgoritmoCompressione := '';
      if (LList.length = 1) and (LList.item[0].hasChildNodes) then
        LAlgoritmoCompressione := LList.item[0].firstChild.nodeValue;

      LList := LAllegatoElement.getElementsByTagName('Attachment');
      LData := '';
      if (LList.length = 1) and (LList.item[0].hasChildNodes) then
        LData := LList.item[0].firstChild.nodeValue;


      LAllegato.FileName := LNomeAttachment;
      LAllegato.FileType := LFormatoAttachment;
      LAllegato.Compressione := LAlgoritmoCompressione;
      LAllegato.Data := LData;

      LButton := TButton.Create(Self);
      try
        LButton.Caption := LNomeAttachment;
        gbAllegati.InsertControl(LButton);
        LButton.Align := TAlign.alTop;

        FAllegati := FAllegati + [LAllegato];
        LButton.Tag := Length(FAllegati) -1;
        LButton.OnClick := AllegatoButtonClick;
      except
        LButton.Free;
        raise;
      end;
    end;
  end;

end;

{ TAllegato }

procedure TAllegato.Clear;
begin
  FileName := '';
  FileType := '';
  Compressione := '';
  Data := '';
  if Assigned(Button) then
    FreeAndNil(Button);
end;

procedure TAllegato.DumpAndOpen;
var
  LTempFileName: string;
  LBytes: TBytes;
  LBytesStream: TBytesStream;
begin
  if Compressione <> '' then
    raise Exception.Create('Compressione ' + Compressione + ' non supportata, allegato: ' + FileName);
  
  LTempFileName := TPath.Combine(TPath.GetTempPath, ExtractFileName(FileName));
  LBytes := TNetEncoding.Base64.DecodeStringToBytes(Data);
  LBytesStream := TBytesStream.Create(LBytes);
  try
    LBytesStream.SaveToFile(LTempFileName);
    ShellExecute(0, nil
      , PWideChar(LTempFileName)
      , nil // PWideChar(TPath.Combine(BASE_FOLDER, FatturePassiveNomeFile.AsString))
      , nil, SW_NORMAL);
  finally
    LBytesStream.Free;
  end;
end;

end.
