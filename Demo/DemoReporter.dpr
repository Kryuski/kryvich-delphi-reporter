program DemoReporter;

uses
  Forms,
  uMain in 'uMain.pas' {fReporterDemoMain},
  KDR.Reporter in '..\KDR.Reporter.pas',
  KDR.DbReporter in '..\KDR.DbReporter.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfReporterDemoMain, fReporterDemoMain);
  Application.Run;
end.

