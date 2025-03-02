program SimpleConsoleDemoApp;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Datasnap.DBClient,
  Winapi.ShellAPI,
  KDR.Reporter in '..\KDR.Reporter.pas',
  KDR.DbReporter in '..\KDR.DbReporter.pas';

const
  TEMPLATE_FILE_NAME = 'Templates\Simple Customer List Report.template.htm';
  REPORT_FILE_NAME = 'Output\Simple Customer List Report.htm';

var
  Customer: TClientDataSet;
  Reporter: TDBReporter;
begin
  try
    Customer := TClientDataSet.Create(nil);
    Reporter := TDBReporter.Create;
    try
      Customer.LoadFromFile('customer.xml');
      Customer.Name := 'Customer';
      Reporter.RegisterDataSet(Customer);
      Reporter.CreateReport(TEMPLATE_FILE_NAME, REPORT_FILE_NAME);
      ShellExecute(0, 'open', '"'+REPORT_FILE_NAME+'"', nil, nil, 0);
    finally
      Reporter.Free;
      Customer.Free;
    end;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
