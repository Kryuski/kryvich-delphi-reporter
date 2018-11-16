unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Grids, DBGrids, ADODB, ExtCtrls, Buttons, DBClient;

type
  TfReporterDemoMain = class(TForm)
    gbTemplates: TGroupBox;
    lbTemplates: TListBox;
    dsOrders: TDataSource;
    dsItems: TDataSource;
    PartSales: TADOQuery;
    PartSalesPartNo: TFloatField;
    PartSalesDescription: TWideStringField;
    PartSalesVendorNo: TFloatField;
    PartSalesVendorName: TWideStringField;
    PartSalesListPrice: TFloatField;
    PartSalesSaleYear: TSmallintField;
    PartSalesOrderNo: TFloatField;
    PartSalesSaleDate: TDateTimeField;
    PartSalesQty: TIntegerField;
    PartSalesAmount: TFloatField;
    bCreateReport: TBitBtn;
    bPrintReport: TBitBtn;
    bViewTemplate: TBitBtn;
    GroupBox1: TGroupBox;
    imgDelphiLogo: TImage;
    bClose: TBitBtn;
    Customer: TClientDataSet;
    Orders: TClientDataSet;
    Employee: TClientDataSet;
    Items: TClientDataSet;
    Parts: TClientDataSet;
    OrderHistory: TADOQuery;
    procedure FormShow(Sender: TObject);
    procedure bCreateReportClick(Sender: TObject);
    procedure bViewTemplateClick(Sender: TObject);
    procedure lbTemplatesDblClick(Sender: TObject);
    procedure bPrintReportClick(Sender: TObject);
    procedure bCloseClick(Sender: TObject);
    procedure imgDelphiLogoClick(Sender: TObject);
  private
    SumOrderSubtotal: real;

    procedure CreateReport(Print: Boolean);
    procedure PopulateTemplateList;
    function GetCustomTagValueForInvoice(const Tag: AnsiString; var Value: string): Boolean;
  public
  end;

var
  fReporterDemoMain: TfReporterDemoMain;

implementation

uses ShellAPI, uDbFreeReporter;

{$R *.dfm}

procedure TfReporterDemoMain.bPrintReportClick(Sender: TObject);
begin
  CreateReport(True);
end;

procedure TfReporterDemoMain.bViewTemplateClick(Sender: TObject);
var
  TemplatePath: string;
begin
  TemplatePath := ExtractFilePath(Application.ExeName) + 'Templates\'
    + lbTemplates.Items[lbTemplates.ItemIndex];
  ShellExecute(Application.MainForm.Handle, nil, pChar(TemplatePath),
    nil, nil, SW_MAXIMIZE);
end;

procedure TfReporterDemoMain.CreateReport(Print: Boolean);
const
  sOperation: array[Boolean] of string = ('open', 'print');
var
  TemplateName, OutputName, sTemplate: string;
  Reporter: TDbFreeReporter;
begin
  if lbTemplates.Count <= 0 then
    Exit;
  sTemplate := lbTemplates.Items[lbTemplates.ItemIndex];
  TemplateName := ExtractFilePath(Application.ExeName) + 'Templates\'
    + sTemplate;
  OutputName := ExtractFilePath(Application.ExeName) + 'Output\'
    + StringReplace(sTemplate, '.template.', '.', [rfIgnoreCase]);

  Reporter := TDbFreeReporter.Create;

  if pos('Simple Customer List Report', sTemplate) = 1 then
    Customer.Open
  else if pos('Invoice Report', sTemplate) = 1 then begin
    Orders.Open;
    Customer.Open;
    Employee.Open;
    Items.Open;
    Parts.Open;
    Reporter.OnGetCustomTagValue := GetCustomTagValueForInvoice;
    SumOrderSubtotal := 0;
    Orders.Locate('OrderNo', 1004, []);
  end else if pos('Order History Report', sTemplate) = 1 then begin
    OrderHistory.Open;
    Employee.Open;
  end else if pos('Part Sales Report', sTemplate) = 1 then begin
    PartSales.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0;'+
      'Data Source=dbdemos.mdb'; // <--- Change path to the database if needed
    PartSales.Open;
  end;

  try
    Reporter.AddDataSet(Orders);
    Reporter.AddDataSet(Customer);
    Reporter.AddDataSet(Employee);
    Reporter.AddDataSet(Items);
    Reporter.AddDataSet(Parts);
    Reporter.AddDataSet(OrderHistory);
    Reporter.AddDataSet(PartSales);
    Reporter.CreateReport(TemplateName, OutputName);
    if Application.MessageBox(pChar('Report created and saved in file'#13+
      '"' + OutputName + '".'#13#13'Do you want to ' + sOperation[Print] + ' it?'),
      'Confirm', MB_ICONQUESTION+MB_YESNO+MB_DEFBUTTON1+MB_APPLMODAL) <> mrYes
    then
      Exit;
    ShellExecute(Application.MainForm.Handle, pChar(sOperation[Print]), pChar(OutputName),
      nil, nil, SW_MAXIMIZE);
  finally
    Reporter.Free;
    Orders.Close;
    Customer.Close;
    Employee.Close;
    Items.Close;
    Parts.Close;
    OrderHistory.Close;
    PartSales.Close;
  end;
end;

procedure TfReporterDemoMain.bCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfReporterDemoMain.bCreateReportClick(Sender: TObject);
begin
  CreateReport(False);
end;

procedure TfReporterDemoMain.FormShow(Sender: TObject);
begin
  PopulateTemplateList;
end;

function TfReporterDemoMain.GetCustomTagValueForInvoice(const Tag: AnsiString;
  var Value: string): Boolean;
var
  r: real;
begin
  Result := True;
  if Tag = 'PartTotal' then begin
    r := Items['Qty'] * Parts['ListPrice'] * (100 - Items['Discount']) / 100;
    SumOrderSubtotal := SumOrderSubtotal + r;
    Value := Format('%m', [r]);
  end else if Tag = 'OrderSubtotal' then
    Value := Format('%m', [SumOrderSubtotal])
  else if Tag = 'OrderTotal' then begin
    r := SumOrderSubtotal * (1 + Orders['TaxRate'] / 100) + Orders['Freight'];
    Value := Format('%m', [r]);
  end else
    Result := False;
end;

procedure TfReporterDemoMain.imgDelphiLogoClick(Sender: TObject);
begin
  ShellExecute(Application.MainForm.Handle, '',
    pChar('https://www.embarcadero.com/products/delphi'), nil, nil, SW_MAXIMIZE);
end;

procedure TfReporterDemoMain.lbTemplatesDblClick(Sender: TObject);
begin
  bCreateReport.Click;
end;

procedure TfReporterDemoMain.PopulateTemplateList;
var
  tpath: string;
  sr: TSearchRec;
begin
  lbTemplates.Clear;
  tpath := ExtractFilePath(Application.ExeName) + 'Templates\';

  if FindFirst(tpath + '*.template.rtf', faAnyFile - faDirectory, sr) = 0 then begin
    repeat
      lbTemplates.Items.Add(sr.Name);
    until FindNext(sr) <> 0;
    FindClose(sr);
  end;

  if FindFirst(tpath + '*.template.htm', faAnyFile - faDirectory, sr) = 0 then begin
    repeat
      lbTemplates.Items.Add(sr.Name);
    until FindNext(sr) <> 0;
    FindClose(sr);
  end;

  lbTemplates.Sorted := True;
  if lbTemplates.Count > 0 then
    lbTemplates.ItemIndex := 0;
end;

end.

