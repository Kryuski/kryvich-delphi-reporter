(* Клас для стварэння справаздач у фарматах RTF і HTML на падставе шаблонаў.
  Class to create RTF & HTML reports based on file templates.
  Copyright (C) 2007-2018 Kryvich
  https://sites.google.com/site/kryvich/reporter
  Licensed under LGPLv3 or later: http://www.gnu.org/licenses/lgpl.html
*)

unit uDbFreeReporter;

interface

uses Classes, DB, uFreeReporter;

type
  // Calls when the reporter meets with unknown tag in the template
  // Tag - name of the unknown tag
  // Value - a return value of the tag
  TGetCustomTagValueFunc = function(const Tag: AnsiString; var Value: string):
    Boolean of object;

  // Database-related reporter which supports an evaluation of values of
  // the database fields
  TDbFreeReporter = class(TFreeReporter)
  private
    // Datasets used in the template
    DataSets: TList;    // list of TDataSet
    // Nested Dataset queue
    Cycles: TList;      // list of TCycleData
    // Current cycle index, -1 if none
    iCurrCycle: Integer;
    // Aggregates data
    Aggregates: TList;  // list of TAggregates

    // 'for' tag event handler
    function ForTag(const Param: AnsiString): Boolean;
    // 'end' tag event handler
    function EndTag: Boolean;
    // Find dataset DataSetName in supported datamodule, form or other component
    function FindDataSet(const DataSetName: string): TDataSet;
    // Update aggregates values for given cycle index
    procedure UpdateAggregatesFor(iCycle: Integer);
    // Is it any opened cycle for the DataSet
    function CycleExistsFor(DataSet: TDataSet): Boolean;
  protected
    // Return a value for the field of the template
    function GetTagValue(const TagName: AnsiString): string; override;
  public
    OnGetCustomTagValue: TGetCustomTagValueFunc;

    constructor Create;
    destructor Destroy; override;
    procedure AddDataSet(DataSet: TDataSet);
    // Return formatted string for the given field value
    function FormatFieldValue(Field: TField; Value: Variant): string;
  end;

implementation

uses SysUtils, Variants,
{$IFDEF UNICODE}
  AnsiStrings
{$ELSE}
  StrUtils
{$ENDIF}
  ;

resourcestring
  rsNoDataset = 'Dataset "%s" not found.'#13'Check template!';
  rsNoDatasetField = 'Dataset field "%s" not found.'#13'Check template!';
  rsNoTag = 'Unknown tag "%s" found.'#13'Check template!';
  rsBadCycleNesting = 'Cycle "%s" can''t be closed before nested cycle "%s"'#13+
    'Check template!';
  rsNoCycleForDataset = 'Unable to evaluate the aggregate "%s" because'#13 +
    'there is no cycle for the dataset "%s".'#13'Check template!';
  rsClosedDataset = 'Unable to retrieve data from the closed dataset "%s"';
  rsBadAggregateTag = 'Unable to evaluate the aggregate field tag "%s"'#13 +
    'Check template!';

type
  TDbReporterException = class(Exception);

  // Aggregates data for a field
  PAggregates = ^TAggregates;
  TAggregates = record
    Field: TField;
    Sum: Variant;
    Count: Integer;
    Min: Variant;
    Max: Variant;
  end;

  // Cycle information
  PCycleData = ^TCycleData;
  TCycleData = record
    // The cycle is not completed?
    Active: Boolean;
    // Index of the parent cycle
    ParentInd: Integer;
    // Iterated dataset
    DataSet: TDataSet;
    // Iterated field, nil allowed
    Field: TField;
    // Initial value of iterated field
    InitFieldValue: Variant;
    // Range of indexes of associated aggregates in TDbFreeReporter.Aggregates
    iAgg1, iAgg2: Integer;
    // Count of iterated records
    RecordCount: Integer;
  end;

{ TDBFreeReporter }

procedure TDbFreeReporter.AddDataSet(DataSet: TDataSet);
begin
  DataSets.Add(DataSet);
end;

constructor TDbFreeReporter.Create;
begin
  inherited;
  OnForTag := ForTag;
  OnEndTag := EndTag;
  DataSets := TList.Create;
  Cycles := TList.Create;
  iCurrCycle := -1;
  Aggregates := TList.Create;
end;

function TDbFreeReporter.CycleExistsFor(DataSet: TDataSet): Boolean;
var
  i: Integer;
  cd: PCycleData;
begin
  for i := 0 to Cycles.Count - 1 do begin
    cd := PCycleData(Cycles[i]);
    if cd.Active and (cd.DataSet = DataSet) then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

destructor TDbFreeReporter.Destroy;
var
  i: Integer;
begin
  DataSets.Free;
  for i := 0 to Cycles.Count - 1 do
    Dispose(PCycleData(Cycles[i]));
  Cycles.Free;
  for i := 0 to Aggregates.Count - 1 do
    Dispose(PAggregates(Aggregates[i]));
  Aggregates.Free;
  inherited;
end;

function TDbFreeReporter.EndTag: Boolean;

  // Is it a parent cycle that indicates an end of the record range
  function StopParentCyclesFor(DataSet: TDataSet): Boolean;
  var
    i: Integer;
    cd: PCycleData;
  begin
    Result := False;
    for i := iCurrCycle-1 downto 0 do begin
      cd := PCycleData(Cycles[i]);
      Result := cd.Active and (cd.DataSet = DataSet) and (cd.Field <> nil)
        and (cd.InitFieldValue <> cd.Field.Value);
      if Result then
        Exit;
    end;
  end;

  // Update aggregates values for all cycles that iterate DataSet records
  procedure UpdateAggregatesForDataSet(DataSet: TDataSet);
  var
    i: Integer;
    cd: PCycleData;
  begin
    for i := iCurrCycle downto 0 do begin
      cd := PCycleData(Cycles[i]);
      if cd.Active and (cd.DataSet = DataSet) then
        UpdateAggregatesFor(i);
    end;
  end;

var
  cd, cd1: PCycleData;
  ags: PAggregates;
  i: Integer;
begin
  Assert(iCurrCycle >= 0, 'FreeReporter: Internal error #02');
  cd := PCycleData(Cycles[iCurrCycle]);
  // Delete all nested cycles if any
  for i := Cycles.Count-1 downto iCurrCycle+1 do begin
    cd1 := PCycleData(Cycles[i]);
    if cd1.ParentInd = iCurrCycle then begin
      if cd1.Active then
        raise TDbReporterException.CreateFmt(rsBadCycleNesting,
          [cd.DataSet.Name, cd1.DataSet.Name]);
      Dispose(cd1);
      Cycles.Delete(i);
    end;
  end;
  // Delete aggregates of every nested cycle if any
  for i := Aggregates.Count-1 downto cd.iAgg2+1 do begin
    ags := PAggregates(Aggregates[i]);
    Dispose(ags);
    Aggregates.Delete(i);
  end;
  repeat
    cd.DataSet.Next;
    Result := cd.DataSet.Eof or StopParentCyclesFor(cd.DataSet);
    if Result then begin
      if not cd.DataSet.Eof and CycleExistsFor(cd.DataSet) then
        cd.DataSet.Prior; // Position will be advanced by the parent cycle
      Break;
    end;
    UpdateAggregatesForDataSet(cd.DataSet)
  until (cd.Field = nil) or (cd.InitFieldValue <> cd.Field.Value);
  if Result then begin
    // Activate parent cycle if exists
    cd.Active := False;
    iCurrCycle := cd.ParentInd;
  end else if cd.Field <> nil then
    cd.InitFieldValue := cd.Field.Value;
end;

function TDbFreeReporter.FindDataSet(const DataSetName: string): TDataSet;
var
  i: Integer;
begin
  for i := 0 to DataSets.Count - 1 do begin
    Result := TDataSet(DataSets[i]);
    if SameText(Result.Name, DataSetName) then
      Exit;
  end;
  Result := nil;
end;

function TDbFreeReporter.FormatFieldValue(Field: TField;
  Value: Variant): string;
var
  FmtStr: string;

  function IntegerFieldText(fld: TIntegerField): string;
  begin
    if (DocType = dtXML)
      or (FmtStr = '')
    then
      Result := IntToStr(Value)
    else
      Result := FormatFloat(FmtStr, Value);
  end;

  function FloatFieldText(fld: TFloatField): string;
  var
    format: TFloatFormat;
    formatSettings: TFormatSettings;
    digits: Integer;
  begin
    if DocType = dtXml then begin
      format := ffGeneral;
      digits := 0;
      formatSettings.ThousandSeparator := #0;
      formatSettings.DecimalSeparator := '.';
      Result := FloatToStrF(Value, format, fld.Precision, digits, formatSettings);
    end else if FmtStr = '' then begin
      if fld.currency then begin
        format := ffCurrency;
        digits := {$IFDEF Unicode}FormatSettings.{$ENDIF}CurrencyDecimals;
      end else begin
        format := ffGeneral;
        digits := 0;
      end;
      Result := FloatToStrF(Value, format, fld.Precision, digits);
    end else
      Result := FormatFloat(FmtStr, Value);
  end; 

  function DateTimeFieldText(fld: TDateTimeField): string;
  var
    f: string;
  begin
    if DocType = dtXml then
      f := 'yyyy-mm-ddThh:nn:ss.zzz'
    else if fld.DisplayFormat <> '' then
      f := fld.DisplayFormat
    else
      case fld.DataType of
        ftDate: f := {$IFDEF Unicode}FormatSettings.{$ENDIF}ShortDateFormat;
        ftTime: f := {$IFDEF Unicode}FormatSettings.{$ENDIF}LongTimeFormat;
      end;
    DateTimeToString(Result, f, Value);
  end; 

begin
  if VarIsNull(Value) then
    Result := ''
  else if Assigned(Field.onGetText) then
    Field.onGetText(Field, Result, True)
  else if Field is TNumericField then begin
    if TNumericField(Field).DisplayFormat <> '' then
      FmtStr := TNumericField(Field).DisplayFormat
    else
      FmtStr := TNumericField(Field).EditFormat;
    if Field is TFloatField then
      Result := FloatFieldText(TFloatField(Field))
    else if Field is TIntegerField then
      Result := IntegerFieldText(TIntegerField(Field))
    else
      Result := VarToStr(Value);
  end else if Field is TDateTimeField then
    Result := DateTimeFieldText(TDateTimeField(Field))
  else
    Result := VarToStr(Value);
end;

function TDbFreeReporter.ForTag(const Param: AnsiString): Boolean;
var
  ds: TDataSet;
  fld: TField;
  cd: PCycleData;
  ags: PAggregates;
  i: Integer;
  ft: TFieldType;
  s: AnsiString;
begin
  i := {$IFDEF Unicode}AnsiStrings.{$ENDIF}AnsiPos('.', Param);
  if i = 0 then
    s := Param
  else
    s := Copy(Param, 1, i-1);
  ds := FindDataSet(string(s));
  if ds = nil then
    raise TDbReporterException.CreateFmt(rsNoDataset, [s]);
  if not ds.Active then
    raise TDbReporterException.CreateFmt(rsClosedDataset, [s]);
  if i > 0 then begin
    s := Copy(Param, i+1, MaxInt);
    fld := ds.FindField(string(s));
    if fld = nil then
      raise TDbReporterException.CreateFmt(rsNoDatasetField, [Param]);
  end else
    fld := nil;
  if not CycleExistsFor(ds) then
    ds.First;
  Result := not ds.Eof;
  if Result then begin
    // Create cycle
    New(cd);
    FillChar(cd^, SizeOf(TCycleData), 0);
    cd.Active := True;
    cd.ParentInd := iCurrCycle;
    cd.DataSet := ds;
    cd.Field := fld;
    if Fld <> nil then
      cd.InitFieldValue := Fld.Value;
    // Init aggregates
    cd.iAgg1 := Aggregates.Count;
    for i := 0 to ds.FieldCount - 1 do begin
      ft := ds.Fields[i].DataType;
      if not (ft in ftNonTextTypes) then begin
        New(ags);
        FillChar(ags^, SizeOf(TAggregates), 0);
        ags.Field := ds.Fields[i];
        ags.Min := Null;
        ags.Max := Null;
        Aggregates.Add(ags);
      end;
    end;
    cd.iAgg2 := Aggregates.Count-1;
    Cycles.Add(cd);
    iCurrCycle := Cycles.Count-1;
    UpdateAggregatesFor(iCurrCycle);
  end;
end;

function TDbFreeReporter.GetTagValue(const TagName: AnsiString): string;
type
  TAggregateFunc = (afNone, afSum, afAvg, afCount, afMin, afMax);

  function GetAggregateFunc(s: AnsiString): TAggregateFunc;
  const
    sAggregateFunc: array[TAggregateFunc] of AnsiString = (
      '', 'sum', 'avg', 'count', 'min', 'max'
    );
  begin
    s := LowerCase(s);
    Result := high(TAggregateFunc);
    while Result > low(TAggregateFunc) do
      if s = sAggregateFunc[Result] then
        Exit
      else
        Dec(Result);
  end;

  function GetAggregateFuncValue(Field: TField; AggregateFunc: TAggregateFunc): string;
  var
    i: Integer;
    ags: PAggregates;
  begin
    for i := Aggregates.Count-1 downto 0 do begin
      ags := PAggregates(Aggregates[i]);
      if ags.Field = Field then begin
        case AggregateFunc of
          afSum: Result := FormatFieldValue(Field, ags.Sum);
          afAvg: begin
            if not (VarIsNull(ags.Sum) or VarIsNull(ags.Count))
              and (ags.Count > 0)
            then begin
              if VarIsFloat(ags.Sum) then
                Result := FormatFieldValue(Field, ags.Sum / ags.Count)
              else
                Result := FormatFieldValue(Field, ags.Sum div ags.Count);
            end else
              Result := '';
          end;
          afCount: Result := IntToStr(ags.Count);
          afMin: Result := FormatFieldValue(Field, ags.Min);
          afMax: Result := FormatFieldValue(Field, ags.Max);
        end;
        Exit;
      end;
    end;
    raise TDbReporterException.CreateFmt(rsBadAggregateTag, [TagName]);
  end;

  function GetCycleRecordCountFor(DataSet: TDataSet): Integer;
  var
    i: Integer;
    cd: PCycleData;
  begin
    for i := Cycles.Count-1 downto 0 do begin
      cd := PCycleData(Cycles[i]);
      if cd.DataSet = DataSet then begin
        Result := cd.RecordCount;
        Exit;
      end;
    end;
    raise TDbReporterException.CreateFmt(rsNoCycleForDataset, [TagName,
      DataSet.Name]);
  end;

var
  s, slcFieldName: AnsiString;
  i: Integer;
  ds: TDataSet;
  fld: TField;
  AggregateFunc: TAggregateFunc;
begin
  if (TagName <> '') and (DataSets.Count > 0) then begin
    slcFieldName := LowerCase(TagName);
    // It's a aggregate?
    i := {$IFDEF Unicode}AnsiStrings.{$ENDIF}AnsiPos('(', slcFieldName);
    if (i > 0)
      and ({$IFDEF Unicode}AnsiStrings.{$ENDIF}AnsiPos(')', slcFieldName)
        = Length(slcFieldName))
    then begin
      // Sum(Table.Field)
      s := Copy(slcFieldName, 1, i-1);
      AggregateFunc := GetAggregateFunc(s);
      if AggregateFunc <> afNone then begin
        s := Copy(slcFieldName, i+1, Length(slcFieldName)-i-1);
        // Table.Field
        i := {$IFDEF Unicode}AnsiStrings.{$ENDIF}AnsiPos('.', s);
        if i > 0 then begin
          ds := FindDataSet(string(Copy(s, 1, i-1)));
          if ds <> nil then begin
            Delete(s, 1, i);
            if s = '*' then begin
              i := GetCycleRecordCountFor(ds);
              Result := IntToStr(i);
              Exit;
            end else begin
              fld := ds.FindField(string(s));
              if fld <> nil then begin
                Result := GetAggregateFuncValue(fld, AggregateFunc);
                Exit;
              end;
            end;
          end;
        end;
      end;
    end;
    // It's a database field?
    i := {$IFDEF Unicode}AnsiStrings.{$ENDIF}AnsiPos('.', TagName);
    if i > 0 then begin
      s := Copy(TagName, 1, i-1);
      ds := FindDataSet(string(s));
      if ds <> nil then begin
        if not ds.Active then
          raise TDbReporterException.CreateFmt(rsClosedDataset, [ds.Name]);
        s := Copy(TagName, i+1, MaxInt);
        fld := ds.FindField(string(s));
        if fld <> nil then begin
          Result := FormatFieldValue(fld, fld.Value);
          Exit;
        end;
      end else
        raise TDbReporterException.CreateFmt(rsNoDataset, [s]); 
    end else begin
      // It's a date/time function?
      if slcFieldName = 'version' then begin
        Result := Version;
        Exit;
      end else if slcFieldName = 'date' then begin
        Result := DateToStr(Date);
        Exit;
      end else if slcFieldName = 'time' then begin
        Result := TimeToStr(Time);
        Exit;
      end else if slcFieldName = 'now' then begin
        Result := DateTimeToStr(Now);
        Exit;
      end;
    end;
  end;

  if not Assigned(OnGetCustomTagValue) or
    not OnGetCustomTagValue(TagName, Result)
  then
    raise TDbReporterException.CreateFmt(rsNoTag, [TagName]);
end;

procedure TDbFreeReporter.UpdateAggregatesFor(iCycle: Integer);
var
  cd: PCycleData;
  ags: PAggregates;
  i: Integer;
  val: Variant;
begin
  cd := PCycleData(Cycles[iCycle]);
  Inc(cd.RecordCount);
  for i := cd.iAgg1 to cd.iAgg2 do begin
    ags := PAggregates(Aggregates[i]);
    val := ags.Field.Value;
    if not VarIsNull(val) then begin
      Inc(ags.Count);
      if VarIsNull(ags.Min) or (val < ags.Min) then
        ags.Min := val;
      if VarIsNull(ags.Max) or (val > ags.Max) then
        ags.Max := val;
      if (ags.Field.DataType in [ftSmallint, ftInteger, ftWord, ftFloat,
        ftCurrency, ftDate, ftTime, ftDateTime, ftAutoInc, ftLargeint, ftTimeStamp])
      then
        ags.Sum := ags.Sum + val;
    end;
  end;
end;

end.
