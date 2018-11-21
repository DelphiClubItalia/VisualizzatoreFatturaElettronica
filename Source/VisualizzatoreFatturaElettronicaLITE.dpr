program VisualizzatoreFatturaElettronicaLITE;

uses
  Vcl.Forms,
  uMainform in 'uMainform.pas' {Mainform},
  Vcl.Themes,
  Vcl.Styles,
  uInformazioni in 'uInformazioni.pas' {Informazioni};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainform, Mainform);
  Application.CreateForm(TInformazioni, Informazioni);
  Application.Run;
end.
