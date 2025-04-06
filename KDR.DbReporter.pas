(* Клас для стварэння справаздач у фарматах RTF, HTML і XML на базе шаблонаў.
  Class to create RTF, HTML & XML reports based on file templates.
  Copyright (C) 2007-2025 Kryvich
  https://sites.google.com/site/kryvich/reporter
  Licensed under LGPLv3 or later: http://www.gnu.org/licenses/lgpl.html
*)

unit KDR.DbReporter;

interface

uses Classes, DB, KDR.Reporter;

type
  // Fires when the reporter encounters an unknown tag in the template.
  // Name - name of the tag
  // Value - return value of the tag
  TGetCustomTagValueFunc
    = function(const Name: string; out Value: string): Boolean of object;

  // Database-aware reporter that can iterate over database tables and fields
  TDbReporter = class(TReporter)
  private
    Datasets: TList;    // Datasets used in the template (list of TDataSet)
    Loops: TList;       // Nested Dataset queue (list of TLoopData)
    iCurrLoop: Integer; // Index of the current loop, -1 if none
    Aggregates: TList;  // List of TAggregates

    // 'for' tag event handler
    function ForTag(const Param: string): Boolean;
    // 'end' tag event handler
    function EndTag: Boolean;
    // Finds a dataset in the list
    function FindDataSet(const Name: string): TDataSet;
    // Updates aggregate values for a loop
    procedure UpdateAggregatesFor(iLoop: Integer);
    // Checks if any loop is open on the DataSet
    function LoopExistsFor(DataSet: TDataSet): Boolean;
  protected
    // Returns the value for a template field
    function GetTagValue(const TagName: string): string; override;
  public
    OnGetCustomTagValue: TGetCustomTagValueFunc;

    constructor Create;
    destructor Destroy; override;
    // Registers a dataset for use by the reporter
    procedure RegisterDataSet(DataSet: TDataSet);
    // Returns a formatted string for the specified field value
    function FormatFieldValue(Field: TField; Value: Variant): string;
  end;

implementation

uses SysUtils, Variants, NetEncoding;

resourcestring
  rsUnnamedDataset = 'Cannot refer an unnamed dataset.'#13+
    'Assign a value to the Name property of the dataset.';
  rsNoDataset = 'Dataset "%s" not found.'#13'Check template!';
  rsNoDatasetField = 'Dataset field "%s" not found.'#13'Check template!';
  rsNoTag = 'Found unknown tag "%s".'#13'Check template!';
  rsBadLoopNesting = 'Loop "%s" cannot be closed before nested loop "%s"'#13+
    'Check template!';
  rsNoLoopForDataset =
    'Cannot evaluate aggregate "%s" because there is no loop for dataset "%s"'#13+
    'Check template!';
  rsClosedDataset = 'Unable to retrieve data from closed dataset "%s"';
  rsBadAggregateTag = 'Unable to evaluate aggregate field tag "%s"'#13 +
    'Check template!';

type
  TDbReporterException = class(Exception);

  // Aggregated values for a database field
  PAggregates = ^TAggregates;
  TAggregates = record
    Field: TField;
    Sum: Variant;
    Count: Integer;
    Min: Variant;
    Max: Variant;
  end;

  // Loop parameters
  PLoopData = ^TLoopData;
  TLoopData = record
    Active: Boolean;         // The loop is not completed?
    ParentInd: Integer;      // Index of the parent loop
    DataSet: TDataSet;       // Iterated dataset
    Field: TField;           // Iterated field, nil allowed
    InitFieldValue: Variant; // Initial value of the iterated field
    iAgg1, iAgg2: Integer;   // Range of indexes of associated aggregates
                             // in TDbReporter.Aggregates
    RecordCount: Integer;    // Count of iterated records
  end;

{ TDBReporter }

constructor TDbReporter.Create;
begin
  inherited;
  OnForTag := ForTag;
  OnEndTag := EndTag;
  Datasets := TList.Create;
  Loops := TList.Create;
  iCurrLoop := -1;
  Aggregates := TList.Create;
end;

procedure TDbReporter.RegisterDataSet(DataSet: TDataSet);
begin
  if DataSet.Name = '' then
    raise TDbReporterException.Create(rsUnnamedDataset);
  Datasets.Add(DataSet);
end;

function TDbReporter.LoopExistsFor(DataSet: TDataSet): Boolean;
var
  i: Integer;
  loopData: PLoopData;
begin
  for i := 0 to Loops.Count - 1 do begin
    loopData := PLoopData(Loops[i]);
    if loopData.Active
      and (loopData.DataSet = DataSet)
    then
      Exit(True);
  end;
  Result := False;
end;

destructor TDbReporter.Destroy;
var
  i: Integer;
begin
  Datasets.Free;
  for i := 0 to Loops.Count - 1 do
    Dispose(PLoopData(Loops[i]));
  Loops.Free;
  for i := 0 to Aggregates.Count - 1 do
    Dispose(PAggregates(Aggregates[i]));
  Aggregates.Free;
  inherited;
end;

function TDbReporter.EndTag: Boolean;

  // Is there a parent loop in the same dataset that needs to be terminated?
  function StopParentLoopsFor(DataSet: TDataSet): Boolean;
  var
    i: Integer;
    loopData: PLoopData;
  begin
    for i := iCurrLoop-1 downto 0 do begin
      loopData := PLoopData(Loops[i]);
      Result := loopData.Active
        and (loopData.DataSet = DataSet)
        and (loopData.Field <> nil)
        and (loopData.InitFieldValue <> loopData.Field.Value);
      if Result then
        Exit;
    end;
    Result := False;
  end;

  // Updates aggregate values for all loops that iterate over records of DataSet
  procedure UpdateAggregatesForDataSet(DataSet: TDataSet);
  var
    i: Integer;
    loopData: PLoopData;
  begin
    for i := iCurrLoop downto 0 do begin
      loopData := PLoopData(Loops[i]);
      if loopData.Active and (loopData.DataSet = DataSet) then
        UpdateAggregatesFor(i);
    end;
  end;

var
  loopData, loopData1: PLoopData;
  ags: PAggregates;
  i: Integer;
begin
  Assert(iCurrLoop >= 0, 'Reporter: Internal error #02');
  loopData := PLoopData(Loops[iCurrLoop]);
  // Delete all nested loops if any
  for i := Loops.Count-1 downto iCurrLoop+1 do begin
    loopData1 := PLoopData(Loops[i]);
    if loopData1.ParentInd = iCurrLoop then begin
      if loopData1.Active then
        raise TDbReporterException.CreateFmt(rsBadLoopNesting,
          [loopData.DataSet.Name, loopData1.DataSet.Name]);
      Dispose(loopData1);
      Loops.Delete(i);
    end;
  end;
  // Delete aggregates of every nested loop if any
  for i := Aggregates.Count-1 downto loopData.iAgg2+1 do begin
    ags := PAggregates(Aggregates[i]);
    Dispose(ags);
    Aggregates.Delete(i);
  end;
  repeat
    loopData.DataSet.Next;
    Result := loopData.DataSet.Eof or StopParentLoopsFor(loopData.DataSet);
    if Result then begin
      if not loopData.DataSet.Eof
        and LoopExistsFor(loopData.DataSet)
      then
        loopData.DataSet.Prior; // The position will be advanced by the parent loop
      Break;
    end;
    UpdateAggregatesForDataSet(loopData.DataSet)
  until (loopData.Field = nil)
    or (loopData.InitFieldValue <> loopData.Field.Value);
  if Result then begin
    // Activate the parent loop if it exists
    loopData.Active := False;
    iCurrLoop := loopData.ParentInd;
  end else if loopData.Field <> nil then
    loopData.InitFieldValue := loopData.Field.Value;
end;

function TDbReporter.FindDataSet(const Name: string): TDataSet;
var
  i: Integer;
begin
  for i := 0 to Datasets.Count - 1 do begin
    Result := TDataSet(Datasets[i]);
    if SameText(Result.Name, Name) then
      Exit;
  end;
  Result := nil;
end;

function TDbReporter.FormatFieldValue(Field: TField;
  Value: Variant): string;
var
  FmtStr: string;

  function IntegerFieldText(fld: TIntegerField): string;
  begin
    if (DocType = dtXmlSpreedsheet)
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
    if DocType = dtXmlSpreedsheet then begin
      format := ffGeneral;
      digits := 0;
      formatSettings.ThousandSeparator := #0;
      formatSettings.DecimalSeparator := '.';
      Result := FloatToStrF(Value, format, fld.Precision, digits, formatSettings);
    end else if FmtStr = '' then begin
      if fld.currency then begin
        format := ffCurrency;
        digits := FormatSettings.CurrencyDecimals;
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
    format: string;
  begin
    if DocType = dtXmlSpreedsheet then
      format := 'yyyy-mm-dd"T"hh:nn:ss.zzz'
    else if fld.DisplayFormat <> '' then
      format := fld.DisplayFormat
    else
      case fld.DataType of
        ftDate: format := FormatSettings.ShortDateFormat;
        ftTime: format := FormatSettings.LongTimeFormat;
      end;
    DateTimeToString(Result, format, Value);
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
  if DocType in [dtHtml, dtXML, dtXmlSpreedsheet] then
    Result := TNetEncoding.HTML.Encode(Result);
end;

function TDbReporter.ForTag(const Param: string): Boolean;
var
  ds: TDataSet;
  fld: TField;
  loopData: PLoopData;
  ags: PAggregates;
  i: Integer;
  ft: TFieldType;
  s: string;
begin
  i := Pos('.', Param);
  if i = 0 then
    s := Param
  else
    s := Copy(Param, 1, i-1);
  ds := FindDataSet(s);
  if ds = nil then
    raise TDbReporterException.CreateFmt(rsNoDataset, [s]);
  if not ds.Active then
    raise TDbReporterException.CreateFmt(rsClosedDataset, [s]);
  if i > 0 then begin
    s := Copy(Param, i+1, MaxInt);
    fld := ds.FindField(s);
    if fld = nil then
      raise TDbReporterException.CreateFmt(rsNoDatasetField, [Param]);
  end else
    fld := nil;
  if not LoopExistsFor(ds) then
    ds.First;
  Result := not ds.Eof;
  if Result then begin
    // Create a loop
    New(loopData);
    loopData^ := Default(TLoopData);
    loopData.Active := True;
    loopData.ParentInd := iCurrLoop;
    loopData.DataSet := ds;
    loopData.Field := fld;
    if Fld <> nil then
      loopData.InitFieldValue := Fld.Value;
    // Init aggregates
    loopData.iAgg1 := Aggregates.Count;
    for i := 0 to ds.FieldCount - 1 do begin
      ft := ds.Fields[i].DataType;
      if not (ft in ftNonTextTypes) then begin
        New(ags);
        ags^ := Default(TAggregates);
        ags.Field := ds.Fields[i];
        ags.Min := Null;
        ags.Max := Null;
        Aggregates.Add(ags);
      end;
    end;
    loopData.iAgg2 := Aggregates.Count-1;
    Loops.Add(loopData);
    iCurrLoop := Loops.Count-1;
    UpdateAggregatesFor(iCurrLoop);
  end;
end;

function TDbReporter.GetTagValue(const TagName: string): string;
type
  TAggregateFunc = (afNone, afSum, afAvg, afCount, afMin, afMax);

  function GetAggregateFunc(s: string): TAggregateFunc;
  const
    sAggregateFunc: array[TAggregateFunc] of string = (
      '', 'sum', 'avg', 'count', 'min', 'max'
    );
  begin
    s := LowerCase(s);
    Result := High(TAggregateFunc);
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

  function GetLoopRecordCountFor(DataSet: TDataSet): Integer;
  var
    i: Integer;
    cd: PLoopData;
  begin
    for i := Loops.Count-1 downto 0 do begin
      cd := PLoopData(Loops[i]);
      if cd.DataSet = DataSet then
        Exit(cd.RecordCount);
    end;
    raise TDbReporterException.CreateFmt(rsNoLoopForDataset, [TagName,
      DataSet.Name]);
  end;

var
  s, fldName: string;
  i: Integer;
  ds: TDataSet;
  fld: TField;
  AggregateFunc: TAggregateFunc;
begin
  if (TagName <> '') and (Datasets.Count > 0) then begin
    fldName := LowerCase(TagName);
    // Is this an aggregate?
    i := Pos('(', fldName);
    if (i > 0)
      and (Pos(')', fldName) = Length(fldName))
    then begin
      // for.ex. "sum(table.field)"
      s := Copy(fldName, 1, i-1);
      AggregateFunc := GetAggregateFunc(s);
      if AggregateFunc <> afNone then begin
        s := Copy(fldName, i+1, Length(fldName)-i-1);
        // "table.field"
        i := Pos('.', s);
        if i > 0 then begin
          ds := FindDataSet(Copy(s, 1, i-1));
          if ds <> nil then begin
            Delete(s, 1, i);
            if s = '*' then begin
              i := GetLoopRecordCountFor(ds);
              Exit(IntToStr(i));
            end else begin
              fld := ds.FindField(s);
              if fld <> nil then
                Exit(GetAggregateFuncValue(fld, AggregateFunc));
            end;
          end;
        end;
      end;
    end;
    // Is this a database field?
    i := Pos('.', TagName);
    if i > 0 then begin
      s := Copy(TagName, 1, i-1);
      ds := FindDataSet(s);
      if ds <> nil then begin
        if not ds.Active then
          raise TDbReporterException.CreateFmt(rsClosedDataset, [ds.Name]);
        s := Copy(TagName, i+1, MaxInt);
        fld := ds.FindField(s);
        if fld <> nil then
          Exit(FormatFieldValue(fld, fld.Value));
      end else
        raise TDbReporterException.CreateFmt(rsNoDataset, [s]); 
    end else begin
      // Is this a date/time function?
      if fldName = 'version' then
        Exit(Version)
      else if fldName = 'date' then
        Exit(DateToStr(Date))
      else if fldName = 'time' then
        Exit(TimeToStr(Time))
      else if fldName = 'now' then
        Exit(DateTimeToStr(Now));
    end;
  end;
  if not Assigned(OnGetCustomTagValue)
    or not OnGetCustomTagValue(TagName, Result)
  then
    raise TDbReporterException.CreateFmt(rsNoTag, [TagName]);
end;

procedure TDbReporter.UpdateAggregatesFor(iLoop: Integer);
var
  loopData: PLoopData;
  ags: PAggregates;
  i: Integer;
  val: Variant;
begin
  loopData := PLoopData(Loops[iLoop]);
  Inc(loopData.RecordCount);
  for i := loopData.iAgg1 to loopData.iAgg2 do begin
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
