program COMTest;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  CommDev in 'CommDev.pas',
  CommTypes in 'CommTypes.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
