unit nn;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses TeamsGraphDemo;

procedure TForm1.Button1Click(Sender: TObject);
var
  Teams: TTeamsConnector;
  Ok: Boolean;
begin
  Teams := TTeamsConnector.Create(
    '���_CLIENT_ID',
    '���_CLIENT_SECRET',
    '���_TENANT_ID'
  );
  try
    Ok := Teams.SendMessageToChannel(
      '���_TEAM_ID',
      '���_CHANNEL_ID',
      '�������� ��������� �� Delphi!'
    );
    if Ok then
      ShowMessage('��������� ����������!')
    else
      ShowMessage('������ ��������.');
  finally
    Teams.Free;
  end;
end;

end.
