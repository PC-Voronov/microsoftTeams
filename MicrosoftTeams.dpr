program MicrosoftTeams;

uses
  Vcl.Forms,
  nn in 'nn.pas' {Form1},
  TeamsGraphDemo in 'TeamsGraphDemo.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
