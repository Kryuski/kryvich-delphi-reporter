(* Клас для стварэння справаздач у фарматах RTF, HTML і XML на базе шаблонаў.
  Class to create RTF, HTML & XML reports based on file templates.
  Copyright (C) 2007-2025 Kryvich
  https://sites.google.com/site/kryvich/reporter
  Licensed under LGPLv3 or later: http://www.gnu.org/licenses/lgpl.html
*)

unit KDR.Reporter;

interface

uses Classes;

const
  Version = '3.2.1'; // Kryvich's Delphi Reporter version

type
  // Fires when the parser reaches the {for} tag in the template.
  // Param - Dataset or field name to iterate over.
  // The event handler determines whether the reporter is allowed to loop.
  TForTagProc = function(const Param: string): Boolean of object;
  // Fires when the parser reaches the {end} tag in the template.
  // The event handler determines whether the reporter should stop iterating.
  TEndTagProc = function: Boolean of object;
  // Template & output document type
  // dtXmlSpreedsheet - Excel XML Spreedsheet 2003
  TDocType = (dtAutoDetect, dtTxt, dtRtf, dtHtml, dtXml, dtXmlSpreedsheet);
  // Template & report file encoding
  TCharEncoding = (ceAutoDetect, ceAnsi, ceUtf8);

  // Reporter that can iterate over RTF/HTML/XML file content
  TReporter = class
  private
    Template: string;
    Chunks: TList; // List of chunks from the template
    fDocType: TDocType;
    fCharEncoding: TCharEncoding;

    // Formats a string before placing it into a RTF file
    function FormatRtf(const s: string): RawByteString;
  protected
    // Returns the value for a template field
    function GetTagValue(const FieldName: string): string; virtual; abstract;
    // Loads a template from a file
    procedure LoadTemplate(const TemplateFileName: string);
{$IFDEF DEBUG}
    // Dump the chunks to file
    procedure DumpChunks(const FileName: string);
{$ENDIF}
  public
    OnForTag: TForTagProc; // Loop start event handler
    OnEndTag: TEndTagProc; // Loop end event handler

    constructor Create;
    destructor Destroy; override;
    // Create a report using a template
    procedure CreateReport(const TemplateFileName, OutputFileName: string);
    property DocType: TDocType read fDocType;
    property CharEncoding: TCharEncoding read fCharEncoding;
  end;

implementation

uses SysUtils, Variants, StrUtils, Character;

resourcestring
  rsNoClosingBracket = 'Closing bracket "}" not found. Check template!';
  rsEmptyTagFound = 'Empty tag found. Check template!';
  rsBadXmlTemplate = 'Bad formatted XML template. Check template!';
  rsForTagNotFound = 'No matching {for} tag found for {end} tag. '+
    'Check template!';
  rsNoClosingEnd = 'No closing {end} tag found for {for} tag. '+
    'Check template!';

type
  TReporterException = class(Exception);

  // Template chunk kind
  TChunkKind = (pkStaticText, pkFor, pkEnd, pkTag);

  // Template chunk
  PChunk = ^TChunk;
  TChunk = record
    Pos, Len: Integer;  // Position in template and Length
    ChunkKind: TChunkKind;
    Param: string;  // For {for} and other tags: tag parameter
    iPairTag: Integer;  // For {for} and {end} tag: index of the matching tag
  end;

// Removes line breaks and tabs, trims consecutive spaces
function NormalizeSpaces(const s: string): string;
var
  i: Integer;
  isSpace, wasSpace: Boolean;
begin
  Result := s;
  isSpace := False;
  for i := Length(Result) downto 1 do begin
    wasSpace := isSpace;
    isSpace := Result[i].IsWhiteSpace;
    if isSpace then
      if wasSpace
        or (i = 1)
        or (i = Length(Result))
      then
        Delete(Result, i, 1)
      else
        Result[i] := ' ';
  end;
end;

// Detects document type by file name
function DetectDocType(const FileName: string): TDocType;
var
  s: string;
begin
  s := LowerCase(Copy(FileName, Length(FileName)-3, 4));
  if (s = '.rtf') or (s = '.doc') then
    Result := dtRtf
  else if s = '.xml' then
    Result := dtXml
  else if s = '.txt' then
    Result := dtTxt
  else
    Result := dtHtml;
end;

// Detects file encoding based on document type
function DetectCharEncoding(DocType: TDocType): TCharEncoding;
begin
  Assert(DocType <> dtAutoDetect);
  if DocType in [dtTxt, dtRtf] then
    Result := ceAnsi
  else
    Result := ceUtf8;
end;

// Clarifies the type of a XML template based on its contents
procedure ClarifyXmlDocType(var DocType: TDocType; Buf: RawByteString);
begin
  if Pos(RawByteString('<?mso-application progid="Excel.Sheet"?>'),
    Copy(Buf, 1, 100)) > 0
  then
    DocType := dtXmlSpreedsheet;
end;

// Strips RTF formatting from a tag string
function StripRtfFormatting(const s: string): string;
var
  i: Integer;
  ch: AnsiChar;
begin
  Result := '';
  i := 1;
  while i <= Length(s) do begin
    case s[i] of
      #0..#31, '{', '}':; // Ignore RTF-formatting
      '\': begin
        Inc(i);
        if s[i] = '''' then begin // \'ee --> non-ASCII character
          Inc(i);
          HexToBin(PWideChar(@s[i]), @ch, 1);
          Result := Result + string(ch);
          Inc(i);
        end else begin
          while (i <= Length(s))
            and not CharInSet(s[i], [' ', '\', '{', '}'])
          do
            Inc(i);
          if (i <= Length(s)) and (s[i] = '\') then
            Continue;
        end;
      end;
      else Result := Result + s[i];
    end;
    Inc(i);
  end;
end;

// Strips XML formatting from a tag string
function StripXmlFormatting(const s: string): string;
var
  i, i1: Integer;
begin
  Result := s;
  i := 1;
  repeat
    i := PosEx('<', Result, i);
    if i <= 0 then
      Break;
    i1 := PosEx('>', Result, i);
    if i1 <= 0 then
      raise TReporterException.Create(rsBadXmlTemplate);
    Delete(Result, i, i1-i+1);
  until False;
end;

// Finds SubStr in Str between StrBefore and StrAfter, starting at Offset.
// If found, Returns the positions of Str, StrBefore and StrAfter.
function FindBetween(const SubStr, Str, StrBefore, StrAfter: string;
  Offset: Integer; out StrPos, StrBeforePos, StrAfterPos: Integer): Boolean;
var
  i: Integer;
begin
  Result := False;
  StrBeforePos := Pos(StrBefore, Str, Offset);
  if StrBeforePos <= 0 then
    Exit;
  i := StrBeforePos + Length(StrBefore);
  StrAfterPos := PosEx(StrAfter, Str, i);
  if StrAfterPos <= 0 then
    Exit;
  StrPos := PosEx(SubStr, Str, i);
  Result := (StrPos > 0)
    and (StrPos < StrAfterPos);
end;

{ TReporter }

constructor TReporter.Create;
begin
  inherited;
  Chunks := TList.Create;
end;

destructor TReporter.Destroy;
var
  i: Integer;
begin
  for i := 0 to Chunks.Count - 1 do
    Dispose(PChunk(Chunks[i]));
  Chunks.Free;
  inherited;
end;

function TReporter.FormatRtf(const s: string): RawByteString;
var
  i: Integer;
  ch: Char;
begin
  Result := '';
  for i := 1 to Length(s) do begin
    ch := s[i];
    if CharInSet(ch, ['\', '{', '}']) then
      Result := Result + '\' + AnsiChar(ch)
    else if Word(ch) <= 127 then
      Result := Result + AnsiChar(ch)
    else
      Result := Result + {\uc1}'\u' + RawByteString(Word(ch).ToString) + '?';
  end;
end;

procedure TReporter.CreateReport(const TemplateFileName, OutputFileName: string);
var
  fs: TFileStream;

  procedure ProcessChunks(iFrom, iTo: Integer);
  var
    i: Integer;
    s: string;
    buf: RawByteString;
    chunk: PChunk;
  begin
    i := iFrom;
    while i <= iTo do begin
      chunk := PChunk(Chunks[i]);
      case chunk.ChunkKind of
        pkStaticText: begin
          if chunk.Len > 0 then begin
            s := Copy(Template, chunk.Pos, chunk.Len);
            if CharEncoding = ceAnsi then
              buf := AnsiString(s)
            else
              buf := UTF8Encode(s);
            fs.Write(buf[1], Length(buf));
          end;
        end;
        pkFor: begin // Start of the "for" loop
          if Assigned(OnForTag)
            and OnForTag(chunk.Param) // Preparing for iteration
          then
            ProcessChunks(i+1, chunk.iPairTag);
          i := chunk.iPairTag;
        end;
        pkEnd: begin // End of "for" loop
          if not Assigned(OnEndTag)
            or OnEndTag
          then
            Break;
          i := iFrom;
          Continue;
        end;
        pkTag: begin // Other tag
          s := GetTagValue(chunk.Param);
          if Length(s) > 0 then begin
            if DocType = dtRtf then
              buf := FormatRtf(s)
            else if CharEncoding = ceAnsi then
              buf := AnsiString(s)
            else
              buf := UTF8Encode(s);
            fs.Write(buf[1], Length(buf));
          end;
        end;
      end;
      Inc(i);
    end;
  end;

begin
  LoadTemplate(TemplateFileName);
  fs := TFileStream.Create(OutputFileName, fmCreate);
  try
    ProcessChunks(0, Chunks.Count-1);
  finally
    fs.Free;
  end;
end;

procedure TReporter.LoadTemplate(const TemplateFileName: string);

  // Retrieves next reporter tag from template
  function GetNextTag(iStart: Integer; out iTagPos, iAfter: Integer;
    out ChunkKind: TChunkKind; out Param: string): Boolean;
  var
    tag, s: string;
  begin
    iTagPos := PosEx('\{', Template, iStart);
    Result := iTagPos > 0;
    if not Result then
      Exit;
    iStart := iTagPos + 2;
    iAfter := PosEx('\}', Template, iStart);
    if iAfter <= 0 then
      raise TReporterException.Create(rsNoClosingBracket);
    s := Copy(Template, iStart, iAfter-iStart);
    case DocType of
      dtRtf: tag := StripRtfFormatting(s);
      dtXml: tag := StripXmlFormatting(s);
      else tag := s;
    end;
    Inc(iAfter, 2);
    tag := NormalizeSpaces(tag);
    if tag = '' then
      raise TReporterException.Create(rsEmptyTagFound);
    s := LowerCase(tag);
    if (Length(s) > 4)
      and (Copy(s, 1, 4) = 'for ')
    then begin
      ChunkKind := pkFor;
      Param := Trim(Copy(tag, 5, MaxInt));
    end else if s = 'end' then begin
      ChunkKind := pkEnd;
      Param := '';
    end else begin
      ChunkKind := pkTag;
      Param := tag;
    end;
  end;

  procedure NewChunk(Pos: Integer; ChunkKind: TChunkKind; const Param: string);
  var
    chunk: PChunk;
  begin
    New(chunk);
    chunk^ := Default(TChunk);
    chunk.Pos := Pos;
    chunk.ChunkKind := ChunkKind;
    chunk.Param := Param;
    chunk.iPairTag := -1;
    Chunks.Add(chunk);
  end;

var
  fs: TFileStream;
  buf: RawByteString;
  i, i1, i2: Integer;
  s: string;
  chunkKind: TChunkKind;
  chunk, Chunk1: PChunk;
begin
  fs := TFileStream.Create(TemplateFileName, fmOpenRead or fmShareDenyNone);
  try
    SetLength(buf, fs.Size);
    fs.Read(buf[1], fs.Size);
  finally
    fs.Free;
  end;
  if fDocType = dtAutoDetect then begin
    fDocType := DetectDocType(TemplateFileName);
    if fDocType = dtXml then
      ClarifyXmlDocType(fDocType, buf);
  end;
  if fCharEncoding = ceAutoDetect then
    fCharEncoding := DetectCharEncoding(fDocType);
  if CharEncoding = ceAnsi then
    Template := string(buf)
  else
    Template := UTF8ToString(buf);
  NewChunk(1, pkStaticText, '');
  i := 1;
  while i <= Length(Template) do begin
    if not GetNextTag(i, i1, i2, chunkKind, s) then
      Break;
    NewChunk(i1, chunkKind, s);
    if i2 <= Length(Template) then
      NewChunk(i2, pkStaticText, '');
    i := i2;
  end;
  for i := 0 to Chunks.Count-1 do begin
    chunk := PChunk(Chunks[i]);
    if chunk.chunkKind = pkEnd then begin
      for i1 := i-1 downto 0 do begin
        Chunk1 := PChunk(Chunks[i1]);
        if (Chunk1.chunkKind = pkFor)
          and (Chunk1.iPairTag < 0)
        then begin
          chunk.iPairTag := i1;
          Chunk1.iPairTag := i;
          Break;
        end;
      end;
      if chunk.iPairTag < 0 then
        raise TReporterException.Create(rsForTagNotFound);
    end;
  end;
  for i := 0 to Chunks.Count-1 do begin
    chunk := PChunk(Chunks[i]);
    if (chunk.chunkKind = pkFor)
      and (chunk.iPairTag < 0)
    then
      raise TReporterException.Create(rsNoClosingEnd);
    if i < Chunks.Count-1 then
      i2 := PChunk(Chunks[i+1]).Pos
    else
      i2 := Length(Template) + 1;
    chunk.Len := i2 - chunk.Pos;
  end;
end;

{$IFDEF DEBUG}
procedure TReporter.DumpChunks(const FileName: string);
var
  fs: TFileStream;

  procedure Dump(const s: string);
  var
    s1: UTF8String;
  begin
    s1 := UTF8Encode(s+#13);
    fs.Write(s1[1], Length(s1)*SizeOf(AnsiChar));
  end;

var
  i: Integer;
  chunk: PChunk;
begin
  fs := TFileStream.Create(FileName, fmCreate);
  try
    Dump('Chunks.Count = ' + Chunks.Count.ToString);
    Dump('');
    Dump('Pos'#9'Len'#9'ChunkKind'#9'Param'#9'iPairTag');
    for i := 0 to Chunks.Count-1 do begin
      chunk := Chunks[i];
      Dump(Format('%d'#9'%d'#9'%d'#9'%s'#9'%d',
        [chunk.Pos, chunk.Len, Integer(chunk.ChunkKind), chunk.Param, chunk.iPairTag]));
    end;
  finally
    fs.Free;
  end;
end;
{$ENDIF}

end.
