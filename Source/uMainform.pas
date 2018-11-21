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
    TMainform = class(TForm)
        XMLDocument1: TXMLDocument;
        WebBrowserFEpreview: TWebBrowser;
        StatusBar1: TStatusBar;
        Splitter1: TSplitter;
        PanelSX: TPanel;
        GroupBox1: TGroupBox;
        GroupBox2: TGroupBox;
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
        ListBoxFEin: TListBox;
        ListBoxFEout: TListBox;
        Aggiornalistadocumenti1: TMenuItem;
        SpeedButton1: TSpeedButton;
        procedure BitBtnPrintPreviewClick(Sender: TObject);
        procedure FormCreate(Sender: TObject);
        procedure BitBtnPrintClick(Sender: TObject);
        procedure WebBrowserFEpreviewDocumentComplete(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
        procedure ApplicationEvents1Message(var Msg: tagMSG; var Handled: Boolean);
        procedure Informazioni1Click(Sender: TObject);
        procedure BitBtnShowXMLClick(Sender: TObject);
        procedure Aggiornalistadocumenti1Click(Sender: TObject);
        procedure ListBoxFEinDblClick(Sender: TObject);
        procedure ListBoxFEoutDblClick(Sender: TObject);
        procedure SpeedButton1Click(Sender: TObject);
        procedure ComboBoxStileSelect(Sender: TObject);
    private
        { Private declarations }
        currDir,config_file: string;
        dirIn,dirOut: string;
        dirInSearch, dirOutSearch: Boolean;
        file_attuale:string;
        procedure ApriFatturaXML(XMLDoc: string);
        function GetFontSize: integer;
        procedure SetFontSize(Size: integer);
        procedure SetOpticalZoom(Value: integer);
        procedure ReadXmlList(dirIn, dirOut: string);
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

uses uInformazioni;

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

procedure TMainform.Aggiornalistadocumenti1Click(Sender: TObject);
begin
    WebBrowserFEpreview.Navigate('about:blank');
    ReadXmlList(dirIn,dirOut);
    ShowMessage('Lista aggiornata');
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
    Node: IXMLNode;
    XMLTXT: WideString;
    fn_tmp,style: string;
    TagBegin, TagEnd, TagLength: integer;
    txt: WideString;
begin
    file_attuale:=XMLDoc;
    try
        style:='fe_templates\'+ComboBoxStile.Items[ComboBoxStile.ItemIndex];
        txt := TFile.ReadAllText(XMLDoc);

        repeat
            TagBegin := Pos(WideString('<?xml-stylesheet'), WideString(txt));
            TagEnd := Pos('>', txt, TagBegin + 1);
            TagLength := TagEnd - TagBegin + 1;
            Delete(txt, TagBegin, TagLength);
        until TagLength>0;

        XMLDocument1.Active := false;
        XMLDocument1.Xml.Text := txt;
        XMLDocument1.Active := true;

        Node := XMLDocument1.CreateNode('xml-stylesheet', ntProcessingInstr, 'type="text/xsl" href="' + currDir+style + '"');
        XMLDocument1.ChildNodes.Insert(1, Node);
        fn_tmp := System.IOUtils.TPath.GetTempFileName + '.xml';
        XMLDocument1.Xml.SaveToFile(fn_tmp);
        WebBrowserFEpreview.Navigate(fn_tmp);

        StatusBar1.Panels[0].Text := ExtractFileName(XMLDoc);

    except
        on e:Exception do
        begin
            ShowMessage('Errore durante apertura file: '+XMLDoc+#13+#10+'['+e.Message+']');
        end;
    end;
    txt:='';
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

    ReadXmlList(dirIn,dirOut);
end;

procedure TMainform.ReadXmlList(dirIn,dirOut:string);
var
    Files: TStringDynArray;
    fn_xml,fn_ext:string;
    i:integer;
begin
    Self.Cursor := crHourGlass;

    try
        if dirInSearch=true then
            Files := TDirectory.GetFiles(dirIn, '*.*', TSearchOption.soAllDirectories)
        else
            Files := TDirectory.GetFiles(dirIn, '*.*', TSearchOption.soTopDirectoryOnly);

        ListBoxFEin.Clear;
        for i := 0 to Length(Files) - 1 do
        begin
            fn_xml:=Files[i];
            fn_ext:=LowerCase(ExtractFileExt(Files[i]));
            if (fn_ext='.xml') then
            begin
                ListBoxFEin.Items.Add(ExtractFileName(fn_xml));
            end;
        end;

    except
        self.Cursor:=crDefault;
        MessageBox(0, PChar('La directory '+dirIn+' non è accessibile!'), 'Attenzione', MB_ICONWARNING or MB_OK or MB_TOPMOST or MB_TASKMODAL or MB_DEFBUTTON1);
    end;

    try
        if dirOutSearch=true then
            Files := TDirectory.GetFiles(dirOut, '*.*', TSearchOption.soAllDirectories)
        else
            Files := TDirectory.GetFiles(dirOut, '*.*', TSearchOption.soTopDirectoryOnly);

        ListBoxFEout.Clear;
        for i := 0 to Length(Files) - 1 do
        begin
            fn_xml:=Files[i];
            fn_ext:=LowerCase(ExtractFileExt(Files[i]));
            if (fn_ext='.xml')  then
            begin
                ListBoxFEout.Items.Add(ExtractFileName(fn_xml));
            end;
        end;
    except
        self.Cursor:=crDefault;
        MessageBox(0, PChar('La directory '+dirOut+' non è accessibile!'), 'Attenzione', MB_ICONWARNING or MB_OK or MB_TOPMOST or MB_TASKMODAL or MB_DEFBUTTON1);
    end;
    self.Cursor:=crDefault;
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

procedure TMainform.ListBoxFEinDblClick(Sender: TObject);
begin
    ApriFatturaXML(dirIn+'\'+ListBoxFEin.Items[ListBoxFEin.ItemIndex]);
end;

procedure TMainform.ListBoxFEoutDblClick(Sender: TObject);
begin
    ApriFatturaXML(dirOut+'\'+ListBoxFEout.Items[ListBoxFEout.ItemIndex]);
end;

end.
