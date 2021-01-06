program IFTest;

uses
  Forms,
  uTestForm in 'uTestForm.pas' {Form1},
  IFUniversal_TLB in '..\..\..\..\..\Program Files (x86)\Borland\Delphi7\Imports\IFUniversal_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
//  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
