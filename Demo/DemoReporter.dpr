program DemoReporter;

uses
  Forms,
  uMain in 'uMain.pas' {fReporterDemoMain},
  uFreeReporter in '..\uFreeReporter.pas',
  uDbFreeReporter in '..\uDbFreeReporter.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfReporterDemoMain, fReporterDemoMain);
  Application.Run;
end.

