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
