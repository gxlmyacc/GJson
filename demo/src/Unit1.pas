unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, GJsonIntf, StdCtrls;

type
  TForm1 = class(TForm)
    btn1: TButton;
    dlgOpen1: TOpenDialog;
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btn1Click(Sender: TObject);
var
  o: IJsonObject;
begin
  if not dlgOpen1.Execute then Exit;
  o := SOFile(dlgOpen1.FileName);
  ShowMessage(o.A['keywords'].S[1]);
end;

end.
 