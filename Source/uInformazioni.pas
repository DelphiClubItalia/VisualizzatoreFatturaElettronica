unit uInformazioni;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,ShellAPI, ShlObj, ActiveX,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Imaging.pngimage;

type
  TInformazioni = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    BitBtn1: TBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    Image1: TImage;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Informazioni: TInformazioni;

implementation

{$R *.dfm}

procedure TInformazioni.BitBtn1Click(Sender: TObject);
begin
  Close;
end;

end.
