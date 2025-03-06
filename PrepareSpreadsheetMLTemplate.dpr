(* Utility for preparing SpreadsheetML XML file created in Excel
  for use by Kryvich's Delphi Reporter as a template.
  Copyright (C) 2025 Kryvich
  https://github.com/Kryuski/kryvich-delphi-reporter
  Licensed under LGPLv3 or later: http://www.gnu.org/licenses/lgpl.html

  Neslib.Xml library is used.
  https://github.com/neslib/Neslib.Xml
*)

program PrepareSpreadsheetMLTemplate;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  Windows, System.SysUtils, Neslib.Xml, Neslib.Xml.Types;

var
  Doc: IXMLDocument;
  DocIsModified: Boolean;

function CheckProgID: Boolean;
var
  node: TXmlNode;
begin
  node := Doc.DocumentElement;
  Result := (node.Value = 'Workbook')
    and (node.AttributeByName('xmlns') <> nil)
    and (node.AttributeByName('xmlns').Value
      = 'urn:schemas-microsoft-com:office:spreadsheet');
end;

procedure RemoveTableAttributes;
const
  ATTR1 = 'ss:ExpandedColumnCount';
  ATTR2 = 'ss:ExpandedRowCount';
var
  worksheet, table: TXmlNode;
begin
  worksheet := Doc.DocumentElement.ElementByName('Worksheet');
  while worksheet <> nil do begin
    table := worksheet.ElementByName('Table');
    while table <> nil do begin
      if table.AttributeByName(ATTR1) <> nil then begin
        table.RemoveAttribute(ATTR1);
        DocIsModified := True;
      end;
      if table.AttributeByName(ATTR2) <> nil then begin
        table.RemoveAttribute(ATTR2);
        DocIsModified := True;
      end;
      table := table.NextSiblingByName('Table');
    end;
    worksheet := worksheet.NextSiblingByName('Worksheet');
  end;
end;

procedure RestoreDataNumberType;
var
  styles, style, numberFormat, worksheet, table, row, cell, data: TXmlNode;
  attrType, attrStyle: TXmlAttribute;
  styleID, sFormat: string;
begin
  styles := Doc.DocumentElement.ElementByName('Styles');
  worksheet := Doc.DocumentElement.ElementByName('Worksheet');
  while worksheet <> nil do begin
    table := worksheet.ElementByName('Table');
    while table <> nil do begin
      row := table.ElementByName('Row');
      while row <> nil do begin
        cell := row.ElementByName('Cell');
        while cell <> nil do begin
          data := cell.ElementByName('Data');
          if (data <> nil)
            and (Pos('\{', data.Text) > 0)
          then begin
            attrType := data.AttributeByName('ss:Type');
            if (attrType <> nil)
              and ((attrType.Value = 'String')
                or (attrType.Value = 'Error'))
            then begin
              attrStyle := cell.AttributeByName('ss:StyleID');
              if attrStyle <> nil then begin
                styleID := attrStyle.Value;
                style := styles.ElementByAttribute('ss:ID', styleID);
                if style <> nil then begin
                  numberFormat := style.ElementByName('NumberFormat');
                  if numberFormat <> nil then begin
                    sFormat := numberFormat.AttributeByName('ss:Format').Value;
                    if (sFormat = 'Standard')
                      or (Pos('0', sFormat) > 0)
                      or (Pos('#', sFormat) > 0)
                    then
                      attrType.Value := 'Number'
                    else
                      attrType.Value := 'DateTime';
                    DocIsModified := True;
                  end;
                end;
              end;
            end;
          end;
          cell := cell.NextSiblingByName('Cell');
        end;
        row := row.NextSiblingByName('Row');
      end;
      table := table.NextSiblingByName('Table');
    end;
    worksheet := worksheet.NextSiblingByName('Worksheet');
  end;
end;

var
  FileName, BackupFileName: string;
begin
  try
    if ParamCount < 1 then begin
      Writeln('Prepares SpreadsheetML file from Excel for use by KDR as a template.');
      Writeln('Usage: ', ExtractFileName(ParamStr(0)),
        ' template-in-Spreadsheet-2003-format.xml');
      Write('Press ENTER to continue...');
      Readln;
      Exit;
    end;
    FileName := ParamStr(1);
    Doc := TXmlDocument.Create;
    Doc.Load(FileName);
    DocIsModified := False;
    if not CheckProgID then
      raise Exception.Create('This is not a SpreadsheetML XML file.');
    RemoveTableAttributes;
    RestoreDataNumberType;
    if DocIsModified then begin
      BackupFileName := FileName + '.bak';
      if FileExists(BackupFileName) then
        DeleteFile(BackupFileName);
      RenameFile(FileName, BackupFileName);
      Doc.Save(FileName);
    end else
      Writeln('No changes have been made to the template.');
  except
    on E: Exception do begin
      Writeln(E.ClassName, ': ', E.Message);
      Write('Press ENTER to continue...');
      Readln;
    end;
  end;
end.
