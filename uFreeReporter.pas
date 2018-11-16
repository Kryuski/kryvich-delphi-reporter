(* Клас для стварэння справаздач у фарматах RTF і HTML на падставе шаблонаў.
  Class to create RTF & HTML reports based on file templates.
  Copyright (C) 2007-2018 Kryvich
  https://sites.google.com/site/kryvich/reporter
  Licensed under LGPLv3 or later: http://www.gnu.org/licenses/lgpl.html
*)

unit uFreeReporter;

interface

uses Classes;

const
  // Version of the reporter
  Version = '3.1';

type
  // Event calls when the parser reach the tag {for} in a template
  // Param - Dataset name or field name to iterate for
  // The event handler returns a value that determines whether
  // the reporter is allowed to perform the cycle
  TForTagProc = function(const Param: AnsiString): Boolean of object;
  // Event calls when the parser reach the tag {end} in a template
  // The event handler returns a value that determines whether
  // the reporter has to stop iterations
  TEndTagProc = function: Boolean of object;

  // Type of a template & an output document
  TDocType = (dtAutoDetect, dtTxt, dtRtf, dtHtml, dtXml);

  // Character Encoding of a template and a report
  TCharEncoding = (ceAuto, ceAnsi, ceUtf8);

  // Reporter which supports cycles in RTF/HTML files
  TFreeReporter = class
  private
    Template: AnsiString;
    // List of fragmets founded in a template
    Chunks: TList;  // list of TChunk
    fDocType: TDocType;
    fCharEncoding: TCharEncoding;

    // Format string before placing into a RTF file
    function FormatRtf(const s: AnsiString): AnsiString;
    function GetCharEncoding: TCharEncoding;
  protected
    // Return a value for the field of the template
    function GetTagValue(const FieldName: AnsiString): string; virtual; abstract;
    // Loading a template from the file
    procedure LoadTemplate(const TemplateName: string);
  public
    OnForTag: TForTagProc; // begin of a cycle
    OnEndTag: TEndTagProc; // end of a cycle

    constructor Create;
    destructor Destroy; override;
    // Create a report using a template
    procedure CreateReport(const TemplateName, OutputName: string);
    property DocType: TDocType read fDocType write fDocType;
    property CharEncoding: TCharEncoding read GetCharEncoding write fCharEncoding;
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
  rsNoClosingBracket = 'The close parenthesis "}" not found. Check template!';
  rsForTagNotFound = 'There is no matching tag {for} for the tag {end}.'#13+
    'Check template!';
  rsNoClosingEnd = 'No close tag {end} found for the {for} tag.'#13+
    'Check template!';

type
  TReporterException = class(Exception);

  // Kind of a chunk of a template
  TChunkKind = (pkNone, pkFor, pkEnd, pkTag);

  // Chunk of a template
  PChunk = ^TChunk;
  TChunk = record
    Pos, Len: Integer;  // Position in template and Length of a chunk
    ChunkKind: TChunkKind;
    Param: AnsiString;      // For {for} and other tags: parameter of the tag
    iPairTag: Integer;  // For {for} and {end} tag: index of a matching tag
  end;

{ TFreeReporter }

constructor TFreeReporter.Create;
begin
  inherited;
  Chunks := TList.Create;
end;

destructor TFreeReporter.Destroy;
var
  i: Integer;
begin
  for i := 0 to Chunks.Count - 1 do
    Dispose(PChunk(Chunks[i]));
  Chunks.Free;
  inherited;
end;

function TFreeReporter.FormatRtf(const s: AnsiString): AnsiString;
var
  i: Integer;
  ch: AnsiChar;
begin
  Result := '';
  for i := 1 to Length(s) do begin
    ch := s[i];
    if ch in ['\', '{', '}'] then
      Result := Result + '\' + s[i]
    else if ch in [' '..#127] then
      Result := Result + s[i]
    else begin // non-ASCII character --> \'ee
      Result := Result + '\''  ';
      BinToHex(PAnsiChar(@ch), PAnsiChar(@Result[Length(Result)-1]), 1);
    end;
  end;
end;

function TFreeReporter.GetCharEncoding: TCharEncoding;
begin
  if fCharEncoding = ceAuto then
    if DocType in [dtTxt, dtRtf] then
      Result := ceAnsi
    else
      Result := ceUtf8
  else
    Result := fCharEncoding;
end;

procedure TFreeReporter.CreateReport(const TemplateName, OutputName: string);
var
  fs: TFileStream;

  function DetectDocType(const TemplateName: string): TDocType;
  var
    s: string;
  begin
    s := LowerCase(Copy(TemplateName, Length(TemplateName)-3, 4));
    if (s = '.rtf') or (s = '.doc') then
      Result := dtRtf
    else if s = '.xml' then
      Result := dtXml
    else if s = '.txt' then
      Result := dtTxt
    else
      Result := dtHtml;
  end;

  procedure ProcessChunks(iFrom, iTo: Integer);
  var
    i: Integer;
    s: string;
    ansis: AnsiString;
    chunk: PChunk;
  begin
    i := iFrom;
    while i <= iTo do begin
      chunk := PChunk(Chunks[i]);
      case chunk.ChunkKind of
        pkNone: begin
          if chunk.Len > 0 then
            fs.Write(Template[chunk.Pos], chunk.Len);
        end;
        pkFor: begin // Begin of "for" cycle
          if Assigned(OnForTag)
            and OnForTag(chunk.Param) // Drop "for "
          then
            ProcessChunks(i+1, chunk.iPairTag);
          i := chunk.iPairTag;
        end;
        pkEnd: begin // End of "for" cycle
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
              ansis := FormatRtf(AnsiString(s))
            else if CharEncoding = ceAnsi then
              ansis := AnsiString(s)
            else
              ansis := UTF8Encode(s);
            fs.Write(ansis[1], Length(ansis));
          end;
        end;
      end;
      Inc(i);
    end;
  end;

begin
  if DocType = dtAutoDetect then
    DocType := DetectDocType(TemplateName);
  LoadTemplate(TemplateName);
  fs := TFileStream.Create(OutputName, fmCreate);
  try
    ProcessChunks(0, Chunks.Count-1);
  finally
    fs.Free;
  end;
end;

procedure TFreeReporter.LoadTemplate(const TemplateName: string);

  // Extract RTF Formatting from a template tag string
  procedure ExtractRtfFormatting(const s: AnsiString; var Tag,
    Formatting: AnsiString);
  var
    i, i1: Integer;
    ch: AnsiChar;
  begin
    Tag := '';
    Formatting := '';
    i := 1;
    while i <= Length(s) do begin
      case s[i] of
        #0..#31, '{', '}': Formatting := Formatting + s[i];
        '\': begin
          i1 := i;
          Inc(i);
          if s[i] = '''' then begin // \'ee --> non-ASCII character
            Inc(i);
            HexToBin(PAnsiChar(@s[i]), PAnsiChar(@ch), 1);
            Tag := Tag + ch;
            Inc(i);
          end else begin
            while (i <= Length(s)) and not (s[i] in [' ', '\', '{', '}']) do
              Inc(i);
            Formatting := Formatting + Copy(s, i1, i-i1);
            if (i <= Length(s)) and (s[i] = '\') then
              Continue;
            Formatting := Formatting + Copy(s, i, 1);
          end;
        end;
        else Tag := Tag + s[i];
      end;
      Inc(i);
    end;
    Assert(Length(Tag) + Length(Formatting) = Length(s),
      'FreeReporter: Internal error #01');
  end;

  // Get next reporter tag from a template
  function GetNextTag(iStart: Integer; var iTagStart, iAfterTag: Integer;
    var ChunkKind: TChunkKind; var Param: AnsiString): Boolean;
  var
    tag, formatting, s: AnsiString;
  begin
    iTagStart := {$IFDEF UNICODE}AnsiStrings.{$ENDIF}
      PosEx('\{', Template, iStart);
    Result := iTagStart > 0;
    if Result then begin
      iStart := iTagStart + 2;
      iAfterTag := PosEx('\}', Template, iStart);
      if iAfterTag <= 0 then
        raise TReporterException.Create(rsNoClosingBracket);
      if DocType = dtRtf then
        ExtractRtfFormatting(Copy(Template, iStart, iAfterTag-iStart), tag,
          formatting)
      else begin
        tag := Copy(Template, iStart, iAfterTag-iStart);
        formatting := '';
      end;
      if formatting <> '' then begin
        move(tag[1], Template[iStart], Length(tag));
        move(Template[iAfterTag], Template[iStart+Length(tag)], 2); // '\}'
        move(formatting[1], Template[iStart+Length(tag)+2], Length(formatting));
        iAfterTag := iStart+Length(tag);
      end;
      Inc(iAfterTag, 2);
      s := LowerCase(tag);
      if (Length(s) > 3)
        and (Copy(s, 1, 3) = 'for')
        and (s[4] <= ' ')
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
  end;

  procedure NewChunk(aPos: Integer; aChunkKind: TChunkKind;
    const aParam: AnsiString);
  var
    chunk: PChunk;
  begin
    New(chunk);
    FillChar(chunk^, SizeOf(TChunk), 0);
    chunk.Pos := aPos;
    chunk.ChunkKind := aChunkKind;
    chunk.Param := aParam;
    chunk.iPairTag := -1;
    Chunks.Add(chunk);
  end;

var
  fs: TFileStream;
  i, i1, i2: Integer;
  s: AnsiString;
  chunkKind: TChunkKind;
  chunk, Chunk1: PChunk;
begin
  fs := TFileStream.Create(TemplateName, fmOpenRead or fmShareDenyNone);
  try
    SetLength(Template, fs.Size);
    fs.Read(Template[1], fs.Size);
  finally
    fs.Free;
  end;
  NewChunk(1, pkNone, '');
  i := 1;
  while i <= Length(Template) do begin
    if not GetNextTag(i, i1, i2, chunkKind, s) then
      Break;
    NewChunk(i1, chunkKind, s);
    if i2 <= Length(Template) then
      NewChunk(i2, pkNone, '');
    i := i2;
  end;
  for i := 0 to Chunks.Count-1 do begin
    chunk := PChunk(Chunks[i]);
    if chunk.chunkKind = pkEnd then begin
      for i1 := i-1 downto 0 do begin
        Chunk1 := PChunk(Chunks[i1]);
        if (Chunk1.chunkKind = pkFor) and (Chunk1.iPairTag < 0) then begin
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
    if (chunk.chunkKind = pkFor) and (chunk.iPairTag < 0) then
      raise TReporterException.Create(rsNoClosingEnd);
    if i < Chunks.Count-1 then
      i2 := PChunk(Chunks[i+1]).Pos
     else
      i2 := Length(Template) + 1;
    chunk.Len := i2 - chunk.Pos;
  end;
end;

end.
