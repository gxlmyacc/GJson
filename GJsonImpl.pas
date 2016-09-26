unit GJsonImpl;

interface

{$DEFINE UNICODE}

uses
  SysUtils, Windows, Classes, ActiveX, GJsonIntf;

type
  TJsonBase = class;
  TJsonPair = class;
  TJsonObject = class;
  TJsonArray = class;
  TJsonValue = class;
  
  EJsonException = class(Exception);

  TJsonBase = class(TInterfacedObject, IJsonBase)
  private
    FIndent: Integer;
    FIndentOption: TIndentOption;
    FTag: Pointer;
    function GetIndent: Integer;
    function GetIndentOption: TIndentOption;
    function GetTag: Pointer;
    procedure SetIndent(const Value: Integer);
    procedure SetIndentOption(const Value: TIndentOption);
    procedure SetTag(const Value: Pointer);
  protected
    procedure RaiseError(const Msg: JsonString);
    procedure RaiseParseError(const JsonString: JsonString);
    procedure RaiseAssignError(Source: TObject);

    function  _IsJsonObject(const P: PJsonChar; const Len: Integer): Boolean;
    function  _IsJsonArray(const P: PJsonChar; const Len: Integer): Boolean;
    function  _IsJsonString(const P: PJsonChar; const Len: Integer): Boolean;
    function  _AnalyzeJsonValueType(const P: PJsonChar; const Len: Integer): TJsonValueType;

    function  _WriteStream(const AStream: IStream; const ABuffer: JsonString): Integer;
    function  _SaveToStream(const AStream: IStream): Integer; virtual; abstract;
    procedure _Parse(const JsonStr: PJsonChar; const Len: Integer); virtual; abstract;

    function  IsPairBegin(const C: JsonChar): Boolean;
    function  GetPairEnd(const C: JsonChar): JsonChar;
    function  MoveToPair(const PB, PE: PJsonChar): PJsonChar;
  public
    function Implementor: Pointer;

    procedure Parse(const JsonStr: JsonString); virtual;
    procedure ParseStream(AStream: IStream; const Utf8: Boolean = False); virtual;
    procedure ParseFile(const AFileName: JsonString; const Utf8: Boolean = False); virtual;

    function  Stringify: JsonString; virtual;
    function  SaveToStream(AStream: IStream): Integer;
    procedure SaveToFile(const AFileName: JsonString; const AEncode: JsonString = ''); virtual;

    procedure Assign(const Source: IJsonBase); virtual; abstract;
    function  CalcSize(): Integer;

    function Encode(const S: JsonString): JsonString; overload;
    function Encode(const P: PJsonChar; const Len: Integer): JsonString; overload;
    function Decode(const S: JsonString): JsonString; overload;
    function Decode(const P: PJsonChar; const Len: Integer): JsonString; overload;

    procedure Split(const S: JsonString; const Delimiter: JsonChar; const OnFind: TFindSplitStrEvent);
    procedure SplitP(const PB, PE: PJsonChar; const Delimiter: JsonChar; const OnFind: TFindSplitStrEvent);

    function IsJsonObject(const S: JsonString): Boolean;
    function IsJsonArray(const S: JsonString): Boolean;
    function IsJsonString(const S: JsonString): Boolean;
    function IsJsonNumber(const S: JsonString): Boolean;
    function IsJsonBoolean(const S: JsonString): Boolean;
    function IsJsonNull(const S: JsonString): Boolean;

    function AnalyzeJsonValueType(const S: JsonString): TJsonValueType;
  public
    property IndentOption: TIndentOption read GetIndentOption write SetIndentOption;
    property Indent: Integer read GetIndent write SetIndent;
    property Tag: Pointer read GetTag write SetTag;
  end;

  TJsonArray = class(TJsonBase, IJsonArray)
  private
    FList: IInterfaceList;
    FParseStr: JsonString;
    FParseLen: Integer;
    FParsed: Boolean;
    function GetCapacity: Integer;
    function GetValues(const KeyName, KeyValue: JsonString): IJsonValue;
    function GetValueAt(Index: Integer): IJsonValue;
    function GetLength: Integer;
    function GetA(const Index: Integer): IJsonArray;
    function GetB(const Index: Integer): Boolean;
    function GetD(const Index: Integer): Double;
    function GetI(const Index: Integer): Int64;
    function GetO(const Index: Integer): IJsonObject;
    function GetS(const Index: Integer): JsonString;
    function GetR(const Index: Integer): JsonString;
    function GetM(const Index: Integer): TJsonMethod;
    procedure SetCapacity(const Value: Integer);
    procedure SetValueAt(Index: Integer; const Value: IJsonValue);
    procedure SetValues(const KeyName, KeyValue: JsonString;const Value: IJsonValue);
    procedure PutA(const Index: Integer; const Value: IJsonArray);
    procedure PutB(const Index: Integer; const Value: Boolean);
    procedure PutD(const Index: Integer; const Value: Double);
    procedure PutI(const Index: Integer; const Value: Int64);
    procedure PutO(const Index: Integer; const Value: IJsonObject);
    procedure PutS(const Index: Integer; const Value: JsonString);
    procedure PutR(const Index: Integer; const Value: JsonString);
    procedure PutM(const Index: Integer; const Value: TJsonMethod);
  protected
    procedure ParseJosnIfNone;
    procedure _ParseFind(const PaserStr: PJsonChar; const AIndex: Integer; const ASplitStr: JsonString; const AParsing: Boolean);
    procedure _Parse(const JsonStr: PJsonChar; const Len: Integer); override;
    function  _SaveToStream(const AStream: IStream): Integer; override;
    procedure _QuickSort(L, R: Integer; SCompare: TJsonArraySortCompare; const Key: JsonString; CaseSensitive: Boolean);
  public
    constructor Create(const AParseStr: PJsonChar; const AParseLen: Integer; const ACapacity: Integer = 1024);
    destructor Destroy; override;

    procedure Assign(const Source: IJsonBase); override;
    function  Add(): IJsonValue; overload;
    function  Add(const Value: Boolean): Integer; overload;
    function  Add(const Value: Int64): Integer; overload;
    function  Add(const Value: Double): Integer; overload;
    function  Add(const Value: JsonString): Integer; overload;
    function  Add(const Value: TJsonMethod): Integer; overload;
    function  Add(const Value: IJsonArray): Integer; overload;
    function  Add(const Value: IJsonObject): Integer; overload;
    function  Add(const Value: IJsonValue): Integer; overload;
    function  AddSO(const Value: JsonString): IJsonValue;
    function  AddSA(const Value: JsonString): IJsonValue;
    function  First: IJsonValue;
    function  Last: IJsonValue;
    procedure Insert(const Index: Integer; const Value: IJsonValue);
    procedure Delete(const Index: Integer); overload;
    function  Delete(const Value: JsonString; const Key: JsonString = ''; CaseSensitive: Boolean = True): Integer; overload;
    procedure Clear;
    procedure Merge(const Addition: IJsonArray);
    procedure Exchange(const Index1, Index2: Integer);
    function  Join(const separator: JsonString; const Key: JsonString = ''): JsonString;
    function  IndexOf(const Value: JsonString; const Key: JsonString = ''; CaseSensitive: Boolean = True): Integer;
    function  LastIndexOf(const Value: JsonString; const Key: JsonString = ''; CaseSensitive: Boolean = True): Integer;
    procedure Sort(Compare: TJsonArraySortCompare = nil; const Key: JsonString = ''; CaseSensitive: Boolean = True);
    function  Slice(startPos: Integer; endPos: Integer = 0): IJsonArray;
    procedure Reverse;

    function  IsEmpty: Boolean;

    function Call(const Index: Integer; const Param: IJsonValue = nil; const This: IJsonValue = nil): IJsonValue; overload;
    function Call(const Index: Integer; const Param: JsonString; const This: IJsonValue = nil): IJsonValue; overload;    
  public
    property Length: Integer read GetLength;
    property Values[const KeyName, KeyValue: JsonString]: IJsonValue read GetValues write SetValues;
    property ValueAt[Index: Integer]: IJsonValue read GetValueAt write SetValueAt; default;
    property Capacity: Integer read GetCapacity write SetCapacity;
    property A[const Index: Integer]: IJsonArray read GetA write PutA;
    property O[const Index: Integer]: IJsonObject read GetO write PutO;
    property B[const Index: Integer]: Boolean read GetB write PutB;
    property I[const Index: Integer]: Int64 read GetI write PutI;
    property D[const Index: Integer]: Double read GetD write PutD;
    property S[const Index: Integer]: JsonString read GetS write PutS;
    property R[const Index: Integer]: JsonString read GetR write PutR;
    property M[const Index: Integer]: TJsonMethod read GetM write PutM;
  end;

  TJsonHash = class(TInterfacedObject, IJsonHash)
  private
    Buckets: array of PJsonHashItem;
  protected
    function Find(const Key: JsonString): PPJsonHashItem;
    function HashOf(const Key: PJsonChar; const Len: Cardinal): Cardinal; 
  public
    constructor Create(Size: Cardinal = 1024);
    destructor Destroy; override;
    procedure Add(const Key: JsonString; Value: Integer);
    procedure Clear;
    procedure Remove(const Key: JsonString);
    function Modify(const Key: JsonString; Value: Integer): Boolean;
    function ValueOf(const Key: JsonString): Integer;
  end;

  TJsonPair = class(TJsonBase)
  private
    FOwner: TJsonObject;
    FParseStr: JsonString;
    FParseLen: Integer;
    FParsed: Boolean;
    FName: JsonString;
    FValue: IJsonValue;
    function  GetName: JsonString;
    function  GetValue: IJsonValue;
    procedure SetName(const Value: JsonString);
  protected
    procedure ParseJosnIfNone;
    procedure _ParseFind(const PaserStr: PJsonChar; const AIndex: Integer; const ASplitStr: JsonString; const AParsing: Boolean);
    procedure _Parse(const JsonStr: PJsonChar; const Len: Integer); override;
    function  _SaveToStream(const AStream: IStream): Integer; override;
  public
    constructor Create(AOwner: TJsonObject; const AName: JsonString);
    constructor CreateForParse(AOwner: TJsonObject; const AParseStr: PJsonChar; const AParseLen: Integer);
    destructor Destroy; override;

    procedure Assign(const Source: IJsonBase); override;
  public
    property Name: JsonString read GetName write SetName;
    property Value: IJsonValue read GetValue;
  end;

  TJsonPairList = class(TList)
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
  end;
  
  TJsonObject = class(TJsonBase, IJsonObject)
  private
    FList: TJsonPairList;
    FCapacity: Integer;
    FListHash: IJsonHash;
    FAutoAdd: Boolean;
    FHashValid: Boolean;
    FParseStr: JsonString;
    FParseLen: Integer;
    FParsed: Boolean;
    FOnParseObjectProp: TParseObjectProp;
    function GetLength: Integer;
    function GetCapacity: Integer;
    function GetItems(Index: Integer): TJsonPair;
    function GetValues(Name: JsonString): IJsonValue;
    function GetNameAt(const Index: Integer): JsonString;
    function GetValueAt(const Index: Integer): IJsonValue;
    function GetAutoAdd: Boolean;
    function GetA(const Key: JsonString): IJsonArray;
    function GetB(const Key: JsonString): Boolean;
    function GetD(const Key: JsonString): Double;
    function GetI(const Key: JsonString): Int64;
    function GetO(const Key: JsonString): IJsonObject;
    function GetS(const Key: JsonString): JsonString;
    function GetR(const Key: JsonString): JsonString;
    function GetM(const Key: JsonString): TJsonMethod;
    function GetPath(Path: JsonString): IJsonValue;
    function GetOnParseObjectProp: TParseObjectProp;
    procedure SetCapacity(const Value: Integer);
    procedure SetAutoAdd(const Value: Boolean);
    procedure SetValueAt(const Index: Integer; const Value: IJsonValue);
    procedure SetValues(Key: JsonString; const Value: IJsonValue);
    procedure PutA(const Key: JsonString; const Value: IJsonArray);
    procedure PutB(const Key: JsonString; const Value: Boolean);
    procedure PutD(const Key: JsonString; const Value: Double);
    procedure PutI(const Key: JsonString; const Value: Int64);
    procedure PutO(const Key: JsonString; const Value: IJsonObject);
    procedure PutS(const Key, Value: JsonString);
    procedure PutR(const Key, Value: JsonString);
    procedure PutM(const Key: JsonString; const Value: TJsonMethod);
    procedure PutPath(Path: JsonString; const Value: IJsonValue);
    procedure SetOnParseObjectProp(const Value: TParseObjectProp);
  protected
    procedure UpdateHash;
    procedure ParseJosnIfNone;
    procedure _ParseFind(const PaserStr: PJsonChar; const AIndex: Integer; const ASplitStr: JsonString; const AParsing: Boolean);
    procedure _Parse(const JsonStr: PJsonChar; const Len: Integer); override;
    function  _SaveToStream(const AStream: IStream): Integer; override;

    property Items[Index: Integer]: TJsonPair read GetItems;
  public
    constructor Create(AOwner: TJsonBase; const AParseStr: PJsonChar; const AParseLen: Integer; const ACapacity: Integer = 1024);
    destructor Destroy; override;
    
    function  Format(const AStr: JsonString; BeginSep: JsonChar = '{'; EndSep: JsonChar = '}'): JsonString;

    procedure Assign(const Source: IJsonBase); override;
    procedure Merge(const Addition: IJsonObject);

    function  Add(const Key: JsonString): IJsonValue; overload;
    function  Add(const Key: JsonString; const Value: Boolean): Integer; overload;
    function  Add(const Key: JsonString; const Value: Int64): Integer; overload;
    function  Add(const Key: JsonString; const Value: Double): Integer; overload;
    function  Add(const Key: JsonString; const Value: JsonString): Integer; overload;
    function  Add(const Key: JsonString; const Value: TJsonMethod): Integer; overload;
    function  Add(const Key: JsonString; const Value: IJsonArray): Integer; overload;
    function  Add(const Key: JsonString; const Value: IJsonObject): Integer; overload;
    function  Add(const Key: JsonString; const Value: IJsonValue): Integer; overload;
    function  AddSO(const Key: JsonString; const Value: JsonString): IJsonValue;
    function  AddSA(const Key: JsonString; const Value: JsonString): IJsonValue;
    function  AddPair(const Key: JsonString): TJsonPair; overload;
    function  AddPair(const APair: TJsonPair): Integer; overload;
    function  Insert(const Index: Integer; const Key: JsonString): IJsonValue;
    function  IndexOf(const Key: JsonString): Integer;
    procedure Delete(const Index: Integer); overload;
    function  Delete(const Key: JsonString): Integer; overload;
    procedure Clear;

    function Call(const Key: JsonString; const Param: IJsonValue = nil; const This: IJsonValue = nil): IJsonValue; overload;
    function Call(const Key: JsonString; const Param: JsonString; const This: IJsonValue = nil): IJsonValue; overload;
  public
    property Length: Integer read GetLength;
    property Values[Key: JsonString]: IJsonValue read GetValues write SetValues; default;
    property NameAt[const Index: Integer]: JsonString read GetNameAt;
    property ValueAt[const Index: Integer]: IJsonValue read GetValueAt write SetValueAt;
    property AutoAdd: Boolean read GetAutoAdd write SetAutoAdd;
    property Capacity: Integer read GetCapacity write SetCapacity;
    property O[const Key: JsonString]: IJsonObject read GetO write PutO;
    property B[const Key: JsonString]: Boolean read GetB write PutB;
    property I[const Key: JsonString]: Int64 read GetI write PutI;
    property D[const Key: JsonString]: Double read GetD write PutD;
    property S[const Key: JsonString]: JsonString read GetS write PutS;
    property R[const Key: JsonString]: JsonString read GetR write PutR;
    property A[const Key: JsonString]: IJsonArray read GetA write PutA;
    property M[const Key: JsonString]: TJsonMethod read GetM write PutM;
    property Path[Path: JsonString]: IJsonValue read GetPath write PutPath;
    
    property OnParseObjectProp: TParseObjectProp read GetOnParseObjectProp write SetOnParseObjectProp;
  end;

  TJsonValue = class(TJsonBase, IJsonValue)
  private
    FParseStr: JsonString;
    FParseLen: Integer;
    FParsed: Boolean;
    
    FValueType: TJsonValueType;
    FStringValue: JsonString;
    FIntegerValue: Int64;
    FDoubleValue: Double;
    FBooleanValue: Boolean;
    FObjectValue: IJsonObject;
    FArrayValue: IJsonArray;
    FMethodValue: TJsonMethod;
    function GetValueType: TJsonValueType;
    function GetA: IJsonArray;
    function GetB: Boolean;
    function GetI: Int64;
    function GetD: Double;
    function GetO: IJsonObject;
    function GetS: JsonString;
    function GetR: JsonString;
    function GetM: TJsonMethod;
    function GetIsNull: Boolean;
    procedure SetA(const Value: IJsonArray);
    procedure SetB(const Value: Boolean);
    procedure SetI(const Value: Int64);
    procedure SetD(const Value: Double);
    procedure SetO(const Value: IJsonObject);
    procedure SetS(const Value: JsonString);
    procedure SetR(const Value: JsonString);
    procedure SetM(const Value: TJsonMethod);
    procedure SetIsNull(const Value: Boolean);
  protected
    procedure ParseJosnIfNone;
    procedure RaiseValueTypeError(const AsValueType: TJsonValueType);
    
    procedure _Parse(const JsonStr: PJsonChar; const Len: Integer); override;
    function  _SaveToStream(const AStream: IStream): Integer; override;
  public
    constructor Create(const AParseStr: PJsonChar; const AParseLen: Integer);
    destructor Destroy; override;

    procedure Assign(const Source: IJsonBase); override;
    function  Clone: IJsonValue; 
    procedure Clear;

    function IsType(const AType: TJsonValueType): Boolean;
  public
    property ValueType: TJsonValueType read GetValueType;
    property IsNull: Boolean read GetIsNull write SetIsNull;
    property S: JsonString read GetS write SetS;
    property R: JsonString read GetR write SetR;
    property D: Double read GetD write SetD;
    property I: Int64 read GetI write SetI;
    property B: Boolean read GetB write SetB;
    property O: IJsonObject read GetO write SetO;
    property A: IJsonArray read GetA write SetA;
    property M: TJsonMethod read GetM write SetM;
  end;

function _SV(const AStr: JsonString = '{}'): IJsonValue;
function _SO(const AStr: JsonString = '{}'): IJsonObject;
function _SA(const AStr: JsonString = '[]'): IJsonArray;

function _SVFile(const AFileName: JsonString): IJsonValue;
function _SASplit(Str: JsonString; const separator: JsonChar; howmany: Integer = 0): IJsonArray;

function _StrToUnicode(const str: WideString): WideString;

implementation

function _SV(const AStr: JsonString = '{}'): IJsonValue;
begin
  Result := TJsonValue.Create(nil, 0);
  Result.Parse(AStr);
end;

function _SO(const AStr: JsonString = '{}'): IJsonObject;
begin
  Result := _SV(AStr).O;
end;

function _SA(const AStr: JsonString = '[]'): IJsonArray;
begin
  Result := _SV(AStr).A;
end;

function _SVFile(const AFileName: JsonString): IJsonValue;
begin
  Result := _SV();
  Result.ParseFile(AFileName);
end;

function _SASplit(Str: JsonString; const separator: JsonChar; howmany: Integer = 0): IJsonArray;
var
  i: Integer;
begin
  Result := _SA();
  i := Pos(separator, Str);
  while i > 0 do
  begin
    Result.Add(Copy(Str, 1, i-1));
    if (howmany > 0) and (Result.Length >= howmany) then Exit;
    Delete(Str, 1, i);
    i := Pos(separator, Str);
  end;
  if (howmany > 0) and (Result.Length >= howmany) then Exit;
  if (Str <> '') then Result.Add(Str);
end;

function _StrToUnicode(const str: WideString): WideString;
var
  pos, start_offset, len: Integer;
  pstr: PWideChar;
  c: WideChar;
  buf: array[0..5] of WideChar;
type
  TByteChar = record
  case integer of
    0: (a, b: Byte);
    1: (c: WideChar);
  end;
const
  super_hex_chars: PWideChar = '0123456789abcdef';
begin
  if str = '' then Exit;
  pstr := PWideChar(str);
  len := Length(str);
  pos := 0; start_offset := 0;
  while pos < len do
  begin
    c := pstr[pos];
    case c of
      #8,#9,#10,#12,#13,'"','\','/':
      begin
        if(pos - start_offset > 0) then
          Result := Result + Copy(pstr, start_offset+1, pos - start_offset);

        if(c = #8) then Result := Result + '\b'
        else if (c = #9) then Result := Result + '\t'
        else if (c = #10) then Result := Result + '\n'
        else if (c = #12) then Result := Result + '\f'
        else if (c = #13) then Result := Result + '\r'
        else if (c = '"') then Result := Result + '\"'
        else if (c = '\') then Result := Result + '\\'
        else if (c = '/') then Result := Result + '\/';
        inc(pos);
        start_offset := pos;
      end;
    else
      if (Word(c) > 255) then
      begin
        if(pos - start_offset > 0) then
          Result := Result + Copy(pstr, start_offset+1, pos - start_offset);
        buf[0] := '\';
        buf[1] := 'u';
        buf[2] := super_hex_chars[TByteChar(c).b shr 4];
        buf[3] := super_hex_chars[TByteChar(c).b and $f];
        buf[4] := super_hex_chars[TByteChar(c).a shr 4];
        buf[5] := super_hex_chars[TByteChar(c).a and $f];
        Result := Result + buf;
        inc(pos);
        start_offset := pos;
      end
      else
      if (c < #32) or (c > #127) then
      begin
        if(pos - start_offset > 0) then
          Result := Result + Copy(pstr, start_offset+1, pos - start_offset);
        buf[0] := '\';
        buf[1] := 'u';
        buf[2] := '0';
        buf[3] := '0';
        buf[4] := super_hex_chars[ord(c) shr 4];
        buf[5] := super_hex_chars[ord(c) and $f];
        Result := Result + buf;
        inc(pos);
        start_offset := pos;
      end
      else
        inc(pos);
    end;
  end;
  if(pos - start_offset > 0) then
    Result := Result + Copy(pstr, start_offset+1, pos - start_offset);
end;

function _IndentSpace(const AIndent: Integer): JsonString;
{$IFDEF UNICODE}
var
  i: Integer;
{$ENDIF}
begin
  SetLength(Result, AIndent);
  {$IFDEF UNICODE}
  for i := 1 to AIndent do
    Result[i] := ' ';
  {$ELSE}
  FillChar(Result[1], AIndent * JsonCharSize, ' '); 
  {$ENDIF}
end;

function _JsonArrayCompare(const Value1, Value2: IJsonValue; const Key: JsonString; CaseSensitive: Boolean): Integer;
var
  s1, s2: JsonString;
  flags: Cardinal;
begin
  if Key = '' then
  begin
    s1 := Value1.S;
    s2 := Value2.S;
  end
  else
  begin
    s1 := Value1.O.S[Key];
    s2 := Value2.O.S[Key];
  end;
  if CaseSensitive then
    flags := 0
  else
    flags := NORM_IGNORECASE;
  Result := CompareStringW(LOCALE_USER_DEFAULT, flags, PJsonChar(s1), Length(s1), PJsonChar(s2), Length(s2)) - 2;
end;

function _JsonArrayIndexOfCompare(const Value: IJsonValue; const Value2: WideString; const Key: JsonString; CaseSensitive: Boolean): Integer;
var
  s1: JsonString;
  flags: Cardinal;
begin
  if Key = '' then
    s1 := Value.S
  else
    s1 := Value.O.S[Key];
  if CaseSensitive then
    flags := 0
  else
    flags := NORM_IGNORECASE;
  Result := CompareStringW(LOCALE_USER_DEFAULT, flags, PJsonChar(s1), Length(s1), PJsonChar(Value2), Length(Value2)) - 2;
end;

{ TJsonBase }

function TJsonBase.AnalyzeJsonValueType(const S: JsonString): TJsonValueType;
begin
  Result := _AnalyzeJsonValueType(PJsonChar(S), System.Length(S));
end;

function TJsonBase.CalcSize: Integer;
begin
  Result := Length(Self.Stringify)
end;

function TJsonBase.Decode(const S: JsonString): JsonString;
begin
  Result := Decode(PJsonChar(S), Length(S));
end;

function TJsonBase.Decode(const P: PJsonChar; const Len: Integer): JsonString;
  function HexValue(const C: JsonChar): Word;
  begin
    if (C >= '0') and (C <='9') then
      Result := Word(C) - Word('0')
    else
    if (C >= 'a') and (C <='f') then
      Result := (Word(C) - Word('a')) + 10
    else
    if (C >= 'A') and (C <='F') then
      Result :=(Word(C) - Word('A')) + 10
    else
      raise EJsonException.Create('Illegal hexadecimal characters "' + C + '"');
  end;

  function HexToUnicode(const Hex: PJsonChar): JsonString;
  begin
    try
      {$IFDEF UNICODE}
      Result := JsonChar((HexValue(Hex[0]) shl 12) + (HexValue(Hex[1]) shl 8) +
        (HexValue(Hex[2]) shl 4) + HexValue(Hex[3]));
      {$ELSE}
      Result := WideChar((HexValue(Hex[0]) shl 12) + (HexValue(Hex[1]) shl 8) +
        (HexValue(Hex[2]) shl 4) + (HexValue(Hex[3]) shl 0));
      {$ENDIF}    
    except
      raise EJsonException.Create('Illegal four-hex-digits "' + Hex + '"');
    end;
  end;

var
  I: Integer;
  C: JsonChar;
begin
  Result := '';
  I := 0;
  while I < Len do
  begin
    C := P[I];
    Inc(I);
    if C = '\' then
    begin
      C := P[I];
      Inc(I);
      case C of
        'b': Result := Result + #8;
        't': Result := Result + #9;
        'n': Result := Result + #10;
        'f': Result := Result + #12;
        'r': Result := Result + #13;
        'u':
          begin
            if I + 4 > Len then
              RaiseParseError('Json is not correct!');
            Result := Result + HexToUnicode(@P[I]);
            Inc(I, 4);
          end;
      else
        Result := Result + C;
      end;
    end
    else Result := Result + C;
  end;
end;

function TJsonBase.Encode(const S: JsonString): JsonString;
begin
  Result := Encode(PJsonChar(S), Length(S));
end;

function TJsonBase.Encode(const P: PJsonChar; const Len: Integer): JsonString;
var
  I: Integer;
  C: JsonChar;
begin
  Result := '';
  for I := 0 to Len-1 do
  begin
    C := P[I];
    case C of
      '"', '\', '/': Result := Result + '\' + C;
      #8: Result := Result + '\b';
      #9: Result := Result + '\t';
      #10: Result := Result + '\n';
      #12: Result := Result + '\f';
      #13: Result := Result + '\r';
    else
      Result := Result + C;
    end;
  end;
end;

function TJsonBase.GetIndentOption: TIndentOption;
begin
  Result := FIndentOption;
end;

function TJsonBase.GetIndent: Integer;
begin
  Result := FIndent;
end;

function TJsonBase.GetPairEnd(const C: JsonChar): JsonChar;
begin
  case C of
    '{': Result := '}';
    '[': Result := ']';
    '"': Result := '"';
  else
    Result := #0;
  end;
end;

function TJsonBase.GetTag: Pointer;
begin
  Result := FTag;
end;

function TJsonBase.Implementor: Pointer;
begin
  Result := Self;
end;

function TJsonBase.IsJsonArray(const S: JsonString): Boolean;
begin
  Result := _IsJsonArray(PJsonChar(S), Length(S));
end;

function TJsonBase.IsJsonBoolean(const S: JsonString): Boolean;
begin
  Result := SameText(S, 'true') or SameText(S, 'false');
end;

function TJsonBase.IsJsonNull(const S: JsonString): Boolean;
begin
  Result := SameText(S, 'null');
end;

function TJsonBase.IsJsonNumber(const S: JsonString): Boolean;
var
  Number: Extended;
begin
  Result := TryStrToFloat(S, Number);
end;

function TJsonBase.IsJsonObject(const S: JsonString): Boolean;
begin
  Result := _IsJsonObject(PJsonChar(S), Length(S));
end;

function TJsonBase.IsJsonString(const S: JsonString): Boolean;
begin
  Result := _IsJsonString(PJsonChar(S), Length(S));
end;

function TJsonBase.IsPairBegin(const C: JsonChar): Boolean;
begin
  Result := (C = '{') or (C = '[') or (C = '"');
end;

function TJsonBase.MoveToPair(const PB, PE: PJsonChar): PJsonChar;
var
  PairBegin, PairEnd: JsonChar;
  C: JsonChar;
begin
  PairBegin := PB^;
  PairEnd := GetPairEnd(PairBegin);
  Result := PB;
  while Result < PE do
  begin
    Result := Result + 1;
    C := Result^;
    if C = PairEnd then
      Break
    else
    if (PairBegin = '"') and (C = '\') then
      Result := Result + 1
    else
    if (PairBegin <> '"') and IsPairBegin(C) then
      Result := MoveToPair(Result, PE);
  end;
end;

procedure TJsonBase.ParseFile(const AFileName: JsonString; const Utf8: Boolean);
var
  fs: TFileStream;
begin
  fs := TFileStream.Create(AFileName, fmOpenRead);
  try
    ParseStream(TStreamAdapter.Create(fs), Utf8);
  finally
    fs.Free;
  end;
end;

{$IFDEF UNICODE}
function GetFileStr(const AStream: IStream; const IsUtf8: Boolean): WideString;
var
  bom: array[0..2] of Byte;
  buffer: array of AnsiChar;
  iSize: Integer;
  newPos: Int64;
  codePage: Cardinal;
  statstg: TStatStg;
begin
  AStream.Stat(statstg, 0);
  if statstg.cbSize >= 3 then
    AStream.Read(@bom, sizeof(bom), nil)
  else
    ZeroMemory(@bom[0], SizeOf(bom));
  if (bom[0] = $FF) and (bom[1] = $FE) then  //unicode
  begin
    AStream.Seek(2, STREAM_SEEK_SET, newPos);
    SetLength(Result, (statstg.cbSize-newPos) div 2);
    AStream.Read(PWideChar(Result), statstg.cbSize-newPos, nil);
  end
  else
  begin  //Ansi or utf8
    if (bom[0] = $EF) and (bom[1] = $BB) and (bom[2] = $BF) then  //utf-8
    begin
      AStream.Seek(3, STREAM_SEEK_SET, newPos);
      codePage := CP_UTF8;
    end
    else
    begin
      AStream.Seek(0, STREAM_SEEK_SET, newPos);
      if IsUtf8 then
        codePage := CP_UTF8
      else
        codePage := CP_ACP;
    end;
    SetLength(buffer, (statstg.cbSize-newPos));
    AStream.Read(@buffer[0], statstg.cbSize-newPos, nil);
    iSize := MultiByteToWideChar(codePage, 0, @buffer[0], Length(buffer), nil, 0);
    if iSize <= 0 then
      Exit;
    SetLength(Result, iSize);
    MultiByteToWideChar(codePage, 0, @buffer[0], Length(buffer), PWideChar(Result), iSize);
  end;
end;
{$ELSE}
function GetFileStr(const AStream: IStream; const IsUtf8: Boolean): AnsiString;
var
  bom: array[0..2] of Byte;
  bufferw: array of WideChar;
  iSize: Integer;
  newPos: Int64;
  statstg: TStatStg;
begin
  AStream.Stat(statstg, 0);
  if statstg.cbSize >= 3 then
    AStream.Read(@bom, sizeof(bom), nil)
  else
    ZeroMemory(@bom[0], SizeOf(bom));
  if (bom[0] = $FF) and (bom[1] = $FE) then  //unicode
  begin
    AStream.Seek(2, STREAM_SEEK_SET, newPos);
    SetLength(bufferw, (statstg.cbSize-newPos) div 2);
    AStream.Read(@bufferw[0], statstg.cbSize-newPos, nil);
    iSize := WideCharToMultiByte(CP_ACP, 0, @bufferw[0], Length(bufferw), nil, 0, nil, nil);
    if iSize <= 0 then
      Exit;
    SetLength(Result, iSize);
    WideCharToMultiByte(CP_ACP, 0, @bufferw[0], Length(bufferw), PAnsiChar(Result), iSize, nil, nil);
  end
  else
  begin  //Ansi or utf8
    if (bom[0] = $EF) and (bom[1] = $BB) and (bom[2] = $BF) then  //utf-8
      AStream.Seek(3, STREAM_SEEK_SET, newPos)
    else
      AStream.Seek(0, STREAM_SEEK_SET, newPos);
    SetLength(Result, (statstg.cbSize-newPos) div 2);
    AStream.Read(PAnsiChar(Result), statstg.cbSize-newPos, nil);
    if IsUtf8 then
      Result := Utf8ToAnsi(Result);
  end;
end;
{$ENDIF}

procedure TJsonBase.ParseStream(AStream: IStream; const Utf8: Boolean);
begin
  Parse(GetFileStr(AStream, Utf8));
end;

procedure TJsonBase.RaiseAssignError(Source: TObject);
var
  SourceClassName: JsonString;
begin
  if Source is TObject then
    SourceClassName := Source.ClassName
  else
    SourceClassName := 'nil';
  RaiseError({$IFDEF UNICODE}WideFormat{$ELSE}Format{$ENDIF}('assign error: %s to %s', [SourceClassName, ClassName]));
end;

procedure TJsonBase.RaiseError(const Msg: JsonString);
begin
  raise EJsonException.CreateFmt('<%s>%s', [ClassName, Msg]);
end;

procedure TJsonBase.RaiseParseError(const JsonString: JsonString);
begin
  RaiseError({$IFDEF UNICODE}WideFormat{$ELSE}Format{$ENDIF}('parse error: %s', [JsonString]))
end;

procedure TJsonBase.SaveToFile(const AFileName: JsonString; const AEncode: JsonString);
var
  ms: TMemoryStream;
  fs: TFileStream;
  sHeaderStr: {$IF CompilerVersion > 18.5}RawByteString{$ELSE}AnsiString{$IFEND};
  sRawStr: {$IF CompilerVersion > 18.5}RawByteString{$ELSE}AnsiString{$IFEND};
begin
  if (AEncode = '') or (AEncode = 'utf-8') or (AEncode = 'utf8') then
  begin
    if (AEncode = 'utf-8') or (AEncode = 'utf8') then
      sRawStr := #$EF#$BB#$BF + {$IFDEF UNICODE}UTF8Encode{$ELSE}AnsiToUtf8{$ENDIF}(Self.Stringify)
    else
      sRawStr := {$IFDEF UNICODE}AnsiString(Self.Stringify){$ELSE}Self.Stringify{$ENDIF};
    fs := TFileStream.Create(AFileName, fmCreate or fmOpenWrite or fmShareExclusive);
    try
      fs.Write(PAnsiChar(sRawStr)^, Length(sRawStr))
    finally
      fs.Free;
    end;   
  end
  else
  if AEncode = 'unicode' then
  begin
    ms := TMemoryStream.Create();
    try
      sHeaderStr := #$FF#$FE;
      ms.Write(sHeaderStr[1], 2);
      {$IFDEF UNICODE}
      SaveToStream(TStreamAdapter.Create(ms));
      {$ELSE}
      sRawStr := Self.Stringify;
      ms.Size := ms.Position + Length(sRawStr) * 2;
      StringToWideChar(sRawStr, PWideChar(PAnsiChar(ms.Memory) + ms.Position), ms.Size-ms.Position);
      {$ENDIF}
      ms.SaveToFile(AFileName);
    finally
      ms.Free;
    end;  
  end
  else
    RaiseError('Can not support ' + AEncode + ' encode!');
end;

procedure TJsonBase.SetIndentOption(const Value: TIndentOption);
begin
  FIndentOption := Value;
end;

procedure TJsonBase.SetIndent(const Value: Integer);
begin
  FIndent := Value;
end;

procedure TJsonBase.SetTag(const Value: Pointer);
begin
  FTag := Value;
end;

procedure TJsonBase.Split(const S: JsonString; const Delimiter: JsonChar;
  const OnFind: TFindSplitStrEvent);
var
  PB, PE: PJsonChar;
begin
  PB := PJsonChar(S);
  PE := @PB[Length(S)];
  SplitP(PB, PE, Delimiter, OnFind);
end;

function TJsonBase.Stringify: JsonString;
var
  ms: TMemoryStream;
begin
  ms := TMemoryStream.Create();
  try
    SaveToStream(TStreamAdapter.Create(ms));
    SetLength(Result, ms.Size div JsonCharSize);
    MoveMemory(@Result[1], ms.Memory, ms.Size);
  finally
    ms.Free;
  end;
end;

function TJsonBase._WriteStream(const AStream: IStream; const ABuffer: JsonString): Integer;
var
  iWriteSize: Integer;
  statstg: TStatStg;
  dwPos: Int64;
begin
  iWriteSize := Length(ABuffer) * JsonCharSize;
  AStream.Stat(statstg, 0);
  if statstg.cbSize < 64 then
    AStream.SetSize(64);
  AStream.Stat(statstg, 0);
  AStream.Seek(0, STREAM_SEEK_CUR, dwPos);
  if dwPos + iWriteSize > statstg.cbSize then
    AStream.SetSize(statstg.cbSize * 2);
  AStream.Write(PJsonChar(ABuffer), iWriteSize, @Result);
end;

function TJsonBase.SaveToStream(AStream: IStream): Integer;
var
  OldPos: Int64;
  statstg: TStatStg;
begin
  AStream.Seek(0, STREAM_SEEK_CUR, OldPos);
  Result := _SaveToStream(AStream);
  AStream.Stat(statstg, 0);
  if (statstg.cbSize-OldPos) > Result then
    AStream.SetSize(Result + OldPos);
end;

procedure TJsonBase.SplitP(const PB, PE: PJsonChar;
  const Delimiter: JsonChar; const OnFind: TFindSplitStrEvent);

  procedure _TrimCharsFromBegin(var PB: PJsonChar; const PE: PJsonChar);
  begin
    while PB < PE  do
    begin
      if PB >= PE then
        break;
      if PB^ <= #32 then
      begin
        while (PB^ <= #32) and (PB < PE) do
          PB := PB + 1;
      end
      else
      if (PB^ = '/') and ((PB+1)^ = '/') then
      begin
        PB := PB + 2;
        while ((PB-1)^<>#10) and (PB < PE) do
         PB := PB + 1;
      end
      else
      if (PB^ = '/') and ((PB+1)^ = '*') then
      begin
        PB := PB + 2;
        while (((PB-1)^<>'/') or ((PB-2)^<>'*')) and (PB < PE) do
         PB := PB + 1;
      end
      else
        break;
    end;
  end;
  procedure _TrimCharsFromEnd(const PB: PJsonChar; var PE: PJsonChar);
  var
    PtrEnd: PJsonChar;
  begin
    PtrEnd := PE - 1;
    while PB < PtrEnd  do
    begin
      if PB >= PtrEnd then
        break;
      if PtrEnd^ <= #32 then
      begin
        while (PtrEnd^ <= #32) and (PB < PE) do
          PtrEnd := PtrEnd - 1;
      end
      else
      if (PtrEnd^ = '/') and ((PtrEnd-1)^ = '*') then
      begin
        PtrEnd := PtrEnd - 2;
        while (((PtrEnd+1)^<>'/') or ((PtrEnd+2)^<>'*'))  and (PB < PtrEnd) do
         PtrEnd := PtrEnd - 1;
      end
      else
        break;
    end;
    PE := PtrEnd + 1;
  end;

var
  PtrBegin, PtrEnd, PtrTrimEnd: PJsonChar;
  C: JsonChar;
  StrItem: JsonString;
  iIndex: Integer;
begin
  PtrBegin := PB;
  _TrimCharsFromBegin(PtrBegin, PE);
  PtrEnd := PtrBegin;

  if PtrEnd >= PE then
    Exit;
  iIndex := 0;
  while PtrEnd < PE do
  begin
    C := PtrEnd^;
    if C = Delimiter then
    begin
      PtrTrimEnd := PtrEnd;
      _TrimCharsFromEnd(PtrBegin, PtrTrimEnd);
      if PtrTrimEnd <= PtrBegin then
        continue;

      SetLength(StrItem, PtrTrimEnd - PtrBegin);
      MoveMemory(@StrItem[1], PtrBegin, (PtrTrimEnd - PtrBegin) * JsonCharSize);

      OnFind(PB, iIndex, StrItem, True);  
      iIndex := iIndex + 1;
      
      PtrEnd := PtrEnd + 1;
      _TrimCharsFromBegin(PtrEnd, PE);
      PtrBegin := PtrEnd;
      Continue;
    end
    else
    if IsPairBegin(C) then
      PtrEnd := MoveToPair(PtrEnd, PE);

    PtrEnd := PtrEnd + 1;
    _TrimCharsFromBegin(PtrEnd, PE);
  end;

  PtrTrimEnd := PtrEnd;
  _TrimCharsFromEnd(PtrBegin, PtrTrimEnd);

  SetLength(StrItem, PtrTrimEnd - PtrBegin);
  MoveMemory(@StrItem[1], PtrBegin, (PtrTrimEnd - PtrBegin) * JsonCharSize);
  OnFind(PB, iIndex, StrItem, True);
  //notify parse end
  OnFind(PB, iIndex+1, StrItem, False);
end;

function TrimJson(var S: PJsonChar; var iLen: Integer): Boolean;
var
  E: PJsonChar;
begin
  E := S + iLen - 1;
  while (S <= E) and (S^ <= ' ') do S := S + 1;
  if S < E then
  begin
    while E^ <= ' ' do
    begin
      E^ := #0;
      E := E - 1;
    end;
    iLen := E - S + 1;
  end;
  Result := S < E;
end;

procedure TJsonBase.Parse(const JsonStr: JsonString);
var
  pJson: PJsonChar;
  iLen: Integer;
begin
  pJson := PJsonChar(JsonStr);
  iLen := System.Length(JsonStr);
  if not TrimJson(pJson, iLen) then Exit;
  _Parse(pJson, iLen);
end;

function TJsonBase._IsJsonArray(const P: PJsonChar;
  const Len: Integer): Boolean;
begin
  Result := (Len >= 2) and (P[0] = '[') and (P[Len-1] = ']');
end;

function TJsonBase._IsJsonObject(const P: PJsonChar; const Len: Integer): Boolean;
begin
  Result := (Len >= 2) and (P[0] = '{') and (P[Len-1] = '}');
end;

function TJsonBase._IsJsonString(const P: PJsonChar; const Len: Integer): Boolean;
begin
  Result := (Len >= 2) and (P[0] = '"') and (P[Len-1] = '"');
end;

function TJsonBase._AnalyzeJsonValueType(const P: PJsonChar;
  const Len: Integer): TJsonValueType;
var
  IntValue: Int64;
  Number: Extended;
  S: string;
begin
  Result := jvNull;
  if Len >= 2 then
  begin
    if (P[0] = '{') and (P[Len-1] = '}') then
      Result := jvObject
    else
    if (P[0] = '[') and (P[Len-1] = ']') then
      Result := jvArray
    else
    if (P[0] = '"') and (P[Len-1] = '"') then
      Result := jvString
    else
    begin
      S := LowerCase(P);
      if S = 'null' then
        Result := jvNull
      else
      if (S = 'true') or (S = 'false') then
        Result := jvBoolean
      else
      begin
        if (P^ = '0') and ((P+1)^ = 'x') then
          S := '$' + Copy(S, 3, MaxInt);
        if TryStrToInt64(S, IntValue) then
          Result := jvInteger
        else
        if TryStrToFloat(S, Number) then
          Result := jvDouble;
      end;
    end;
  end
  else
    if TryStrToInt64(P, IntValue) then
      Result := jvInteger;
end;

{ TJsonPair }

procedure TJsonPair.Assign(const Source: IJsonBase);
var
  Src: TJsonPair;
begin
  if not(TObject(Source.Implementor) is TJsonPair) then
    RaiseAssignError(TObject(Source.Implementor));
  Src := TJsonPair(Source.Implementor);
  FName := Src.FName;
  if FValue = nil then
    FValue := TJsonValue.Create(nil, 0);
  FValue.Assign(Src.Value);
end;

constructor TJsonPair.Create(AOwner: TJsonObject; const AName: JsonString);
begin
  inherited Create();
  FOwner := AOwner;
  FName := AName;
  FParseStr := '';
  FParsed := True;
end;

constructor TJsonPair.CreateForParse(AOwner: TJsonObject; const AParseStr: PJsonChar; const AParseLen: Integer);
begin
  inherited Create();
  FOwner := AOwner;
  FParseStr := AParseStr;
  FParseLen := AParseLen;
  FParsed := False;
end;

destructor TJsonPair.Destroy;
begin
  FValue := nil;
  inherited Destroy;
end;

procedure TJsonPair._ParseFind(const PaserStr: PJsonChar; const AIndex: Integer; const ASplitStr: JsonString; const AParsing: Boolean);
begin
  if AParsing then
  begin
    case AIndex of
      0:
      begin
        if IsJsonString(ASplitStr) then
          FName := Decode(Copy(ASplitStr, 2, Length(ASplitStr) - 2))
        else
          FName := ASplitStr;
        if FName = EmptyStr then
          RaiseParseError(ASplitStr);
      end;
      1:
      begin
        FValue := TJsonValue.Create(PJsonChar(ASplitStr), Length(ASplitStr));
      end;
    else
      RaiseParseError(PaserStr);
    end;
  end
  else
  begin
    if AIndex <> 2 then RaiseParseError(PaserStr);
    if @FOwner.FOnParseObjectProp <> nil then
      FOwner.FOnParseObjectProp(FOwner, FName, FValue);
  end;
end;

procedure TJsonPair._Parse(const JsonStr: PJsonChar; const Len: Integer);
begin
  try
    SplitP(JsonStr, @JsonStr[Len], ':', _ParseFind);
  finally
    FParsed := True;
    FParseStr := '';
  end;
end;

function TJsonPair._SaveToStream(const AStream: IStream): Integer;
begin
  Result := 0;
  if (FValue = nil) or (FValue.ValueType = jvMethod) then
    Exit;
  if (FIndentOption = ioIndent) and (FIndent > 0) then
    Result := Result + _WriteStream(AStream, _IndentSpace(FIndent));
  Result := Result + _WriteStream(AStream, '"' +Encode(FName) + '":');
  Result := Result + TJsonBase(FValue.Implementor)._SaveToStream(AStream);
end;

procedure TJsonPair.SetName(const Value: JsonString);
begin
  ParseJosnIfNone;
  FName := Value;
end;

function TJsonPair.GetValue: IJsonValue;
begin
  ParseJosnIfNone;
  if FValue = nil then
    FValue := TJsonValue.Create(nil, 0);
  Result := FValue;
end;

procedure TJsonPair.ParseJosnIfNone;
begin
  if (not FParsed) and (FParseStr <> '') then
   _Parse(PJsonChar(FParseStr), FParseLen);
end;

function TJsonPair.GetName: JsonString;
begin
  ParseJosnIfNone;
  Result := FName;
end;

{ TJsonValue }

procedure TJsonValue.Assign(const Source: IJsonBase);
var
  Src: TJsonValue;
begin
  Clear;
  if Source = nil then Exit;
  if TObject(Source.Implementor) is TJsonObject then
  begin
    FValueType := jvObject;
    FObjectValue := TJsonObject.Create(Self, nil, 0);
    FObjectValue.Assign(Source);
  end
  else
  if TObject(Source.Implementor) is TJsonArray then
  begin
    FValueType := jvArray;
    FArrayValue := TJsonArray.Create(nil, 0);
    FArrayValue.Assign(Source);
  end
  else
  if TObject(Source.Implementor) is TJsonValue then
  begin
    Src := TJsonValue(Source.Implementor);
    FValueType := Src.ValueType;
    case FValueType of
      jvNull: ;
      jvString, jvRawString:  FStringValue  := Src.FStringValue;
      jvInteger: FIntegerValue := Src.FIntegerValue;
      jvDouble:  FDoubleValue  := Src.FDoubleValue;
      jvBoolean: FBooleanValue := Src.FBooleanValue;
      jvObject:
        begin
          FObjectValue := TJsonObject.Create(Self, nil, 0);
          FObjectValue.Assign(Src.FObjectValue);
        end;
      jvArray:
        begin
          FArrayValue := TJsonArray.Create(nil, 0);
          FArrayValue.Assign(Src.FArrayValue);
        end;
      jvMethod:  FMethodValue  := Src.FMethodValue;
    end;
  end
  else
    RaiseAssignError(TObject(Source.Implementor));
end;

procedure TJsonValue.Clear;
begin
  case FValueType of
    jvNull: ;
    jvString, jvRawString:  FStringValue := '';
    jvInteger: FIntegerValue := 0;
    jvDouble:  FDoubleValue := 0.0;
    jvBoolean: FBooleanValue := False;
    jvObject:  FObjectValue := nil;
    jvArray:   FArrayValue := nil;
    jvMethod:  FMethodValue := nil;
  end;
  FValueType := jvNull;
end;

function TJsonValue.Clone: IJsonValue;
begin
  ParseJosnIfNone;
  Result := TJsonValue.Create(nil, 0);
  case FValueType of
    jvNull: ;
    jvString:  Result.S := Self.S;
    jvRawString: Result.R := Self.R;
    jvInteger: Result.I := Self.I;
    jvDouble:  Result.D := Self.D;
    jvBoolean: Result.B := Self.B;
    jvObject:  Result.O := Self.O;
    jvArray:   Result.A := Self.A;
    jvMethod:  Result.M := Self.M;
  end;
end;

constructor TJsonValue.Create(const AParseStr: PJsonChar; const AParseLen: Integer);
begin
  inherited Create();
  FParseStr := AParseStr;
  FParseLen := AParseLen;
  FParsed   := AParseLen = 0;
end;

destructor TJsonValue.Destroy;
begin
  Clear;
  inherited Destroy;
end;

function TJsonValue.GetA: IJsonArray;
begin
  ParseJosnIfNone;
  if IsNull then
  begin
    FValueType := jvArray;
    FArrayValue := TJsonArray.Create(nil, 0);
  end;
  if FValueType <> jvArray then RaiseValueTypeError(jvArray);
  Result := FArrayValue;
end;

function TJsonValue.GetB: Boolean;
begin
  ParseJosnIfNone;
  Result := False;
  case FValueType of
    jvNull: Result := False;
    jvString, jvRawString: Result := SameText(FStringValue, 'true');
    jvInteger: Result := (FIntegerValue <> 0);
    jvDouble:  Result := (FDoubleValue <> 0);
    jvBoolean: Result := FBooleanValue;
    jvObject:  Result := FObjectValue <> nil;
    jvArray:  Result := FArrayValue <> nil;
    jvMethod: Result := @FMethodValue <> nil;
  end;
end;

function TJsonValue.GetI: Int64;
begin
  ParseJosnIfNone;
  Result := 0;
  case FValueType of
    jvNull: Result := 0;
    jvString, jvRawString:  Result := Trunc(StrToFloat(FStringValue));
    jvInteger: Result := FIntegerValue;
    jvDouble:  Result := Trunc(FDoubleValue);
    jvBoolean: Result := Ord(FBooleanValue);
    jvObject, jvArray, jvMethod: RaiseValueTypeError(jvInteger);
  end;
end;

function TJsonValue.GetM: TJsonMethod;
begin
  ParseJosnIfNone;
  Result := FMethodValue;
end;

function TJsonValue.GetD: Double;
begin
  ParseJosnIfNone;
  Result := 0;
  case FValueType of
    jvNull:    Result := 0;
    jvString, jvRawString:  Result := StrToFloat(FStringValue);
    jvInteger: Result := FIntegerValue;
    jvDouble:  Result := FDoubleValue;
    jvBoolean: Result := Ord(FBooleanValue);
    jvObject, jvArray, jvMethod: RaiseValueTypeError(jvDouble);
  end;
end;

function TJsonValue.GetO: IJsonObject;
begin
  ParseJosnIfNone;
  if IsNull then
  begin
    FValueType := jvObject;
    FObjectValue := TJsonObject.Create(Self, nil, 0);
  end;
  if FValueType <> jvObject then RaiseValueTypeError(jvObject);
  Result := FObjectValue;
end;

function TJsonValue.GetS: JsonString;
const
  BooleanStr: array[Boolean] of JsonString = ('false', 'true');
begin
  ParseJosnIfNone;
  Result := '';
  case FValueType of
    jvNull: Result := '';
    jvString: Result := FStringValue;
    jvRawString: Result := Decode(FStringValue);
    jvInteger: Result := IntToStr(FIntegerValue);
    jvDouble:  Result := FloatToStr(FDoubleValue);
    jvBoolean: Result := BooleanStr[FBooleanValue];
    jvObject:
      if FObjectValue = nil then
        Result := ''
      else
        Result := FObjectValue.Stringify;
    jvArray:
      if FArrayValue = nil then
        Result := ''
      else
        Result := FArrayValue.Stringify;
    jvMethod: RaiseValueTypeError(jvString);
  end;
end;

function TJsonValue.GetR: JsonString;
const
  BooleanStr: array[Boolean] of JsonString = ('false', 'true');
begin
  ParseJosnIfNone;
  Result := '';
  case FValueType of
    jvNull: Result := '';
    jvString: Result := Decode(FStringValue);
    jvRawString: Result := FStringValue;
    jvInteger: Result := IntToStr(FIntegerValue);
    jvDouble:  Result := FloatToStr(FDoubleValue);
    jvBoolean: Result := BooleanStr[FBooleanValue];
    jvObject:
      if FObjectValue = nil then
        Result := ''
      else
        Result := FObjectValue.Stringify;
    jvArray:
      if FArrayValue = nil then
        Result := ''
      else
        Result := FArrayValue.Stringify;
    jvMethod: RaiseValueTypeError(jvString);
  end;
end;

function TJsonValue.GetIsNull: Boolean;
begin
  ParseJosnIfNone;
  Result := (FValueType = jvNull);
end;

function TJsonValue.GetValueType: TJsonValueType;
begin
  ParseJosnIfNone;
  Result := FValueType;
end;

procedure TJsonValue._Parse(const JsonStr: PJsonChar; const Len: Integer);
begin
  try
    Clear;
    FValueType := _AnalyzeJsonValueType(JsonStr, Len);
    case FValueType of
      jvNull:    ;//RaiseParseError(JsonStr);
      jvRawString: RaiseParseError(JsonStr);
      jvString:  FStringValue := Decode(JsonStr+1, Len - 2);
      jvInteger: FIntegerValue := StrToInt64(JsonStr);
      jvDouble:  FDoubleValue  := StrToFloat(JsonStr);
      jvBoolean: FBooleanValue := SameText(JsonStr, 'true');
      jvObject:
        FObjectValue := TJsonObject.Create(Self, JsonStr, Len);
      jvArray:
        FArrayValue := TJsonArray.Create(JsonStr, Len);
    end;
  finally
    FParsed := True;
    FParseStr := '';
  end;
end;

procedure TJsonValue.RaiseValueTypeError(
  const AsValueType: TJsonValueType);
const
  StrJsonValueType: array[TJsonValueType] of JsonString =
    ('jvNull', 'jvString', 'jsRawString', 'jvInteger', 'jvFloat', 'jvBoolean', 'jvObject', 'jvArray', 'jvMethod');
var
  S: JsonString;
begin
  S := Format('value type error: %s to %s', [StrJsonValueType[FValueType], StrJsonValueType[AsValueType]]);
  RaiseError(S);
end;

function TJsonValue._SaveToStream(const AStream: IStream): Integer;
const
  StrBoolean: array[Boolean] of JsonString = ('false', 'true');
var
  str: JsonString;
begin
  Result := 0;
  if (FIndentOption = ioNone) and (not FParsed) and (FParseStr <> EmptyStr) then
  begin
    Result := Result + _WriteStream(AStream, FParseStr);
    Exit;
  end;
  ParseJosnIfNone;
  case FValueType of
    jvNull:    str := 'null';
    jvString:  str := '"' + Encode(PJsonChar(FStringValue), System.Length(FStringValue)) + '"';
    jvRawString: str := '"' + FStringValue + '"';
    jvInteger: str := IntToStr(FIntegerValue);
    jvDouble:  str := FloatToStr(FDoubleValue);
    jvBoolean: str := StrBoolean[FBooleanValue];
    jvObject:
    begin
      FObjectValue.IndentOption := FIndentOption;
      Result := Result + TJsonBase(FObjectValue.Implementor)._SaveToStream(AStream);
    end;
    jvArray:
    begin
      FArrayValue.IndentOption := FIndentOption;
      Result := Result + TJsonBase(FArrayValue.Implementor)._SaveToStream(AStream);
    end;
  end;
  if str <> '' then
    Result := Result + _WriteStream(AStream, str);
end;

procedure TJsonValue.SetA(const Value: IJsonArray);
begin
  ParseJosnIfNone;
  if FValueType <> jvArray then
  begin
    Clear;
    FValueType := jvArray;
  end;
  if Value = nil then
    FArrayValue := TJsonArray.Create(nil, 0)
  else
    FArrayValue := Value;
end;

procedure TJsonValue.SetB(const Value: Boolean);
begin
  ParseJosnIfNone;
  if FValueType <> jvBoolean then
  begin
    Clear;
    FValueType := jvBoolean;
  end;
  FBooleanValue := Value;
end;

procedure TJsonValue.SetI(const Value: Int64);
begin
  ParseJosnIfNone;
  if FValueType <> jvInteger then
  begin
    Clear;
    FValueType := jvInteger;
  end;
  FIntegerValue := Value;
end;

procedure TJsonValue.SetM(const Value: TJsonMethod);
begin
  ParseJosnIfNone;
  if FValueType <> jvMethod then
  begin
    Clear;
    FValueType := jvMethod;
  end;
  FMethodValue := Value;
end;

procedure TJsonValue.SetD(const Value: Double);
begin
  ParseJosnIfNone;
  if FValueType <> jvDouble then
  begin
    Clear;
    FValueType := jvDouble;
  end;
  FDoubleValue := Value;
end;

procedure TJsonValue.SetO(const Value: IJsonObject);
begin
  ParseJosnIfNone;
  if FValueType <> jvObject then
  begin
    Clear;
    FValueType := jvObject;
  end;
  if Value = nil then
    FObjectValue := TJsonObject.Create(Self, nil, 0)
  else
    FObjectValue := Value;
end;

procedure TJsonValue.SetS(const Value: JsonString);
begin
  ParseJosnIfNone;
  if FValueType <> jvString then
  begin
    Clear;
    FValueType := jvString;
  end;
  FStringValue := Value;
end;

procedure TJsonValue.SetR(const Value: JsonString);
begin
  ParseJosnIfNone;
  if FValueType <> jvRawString then
  begin
    Clear;
    FValueType := jvRawString;
  end;
  FStringValue := Value;
end;

procedure TJsonValue.SetIsNull(const Value: Boolean);
begin
  ParseJosnIfNone;
  if FValueType <> jvNull then
  begin
    Clear;
    FValueType := jvNull;
  end;
end;

procedure TJsonValue.ParseJosnIfNone;
begin
  if (not FParsed) and (FParseStr <> EmptyStr) then
    _Parse(PJsonChar(FParseStr), FParseLen);
end;

function TJsonValue.IsType(const AType: TJsonValueType): Boolean;
begin
  ParseJosnIfNone;
  Result := FValueType = AType;
end;

{ TJsonArray }

function TJsonArray.Add(const Value: Boolean): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.B := Value;
end;

function TJsonArray.Add(const Value: Int64): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.I := Value;
end;

function TJsonArray.Add(const Value: Double): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.D := Value;
end;

function TJsonArray.Add: IJsonValue;
begin
  ParseJosnIfNone;
  Result := TJsonValue.Create(nil, 0);
  FList.Add(Result);
end;

function TJsonArray.Add(const Value: IJsonObject): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.O := Value;
end;

function TJsonArray.Add(const Value: IJsonValue): Integer;
begin
  ParseJosnIfNone;
  Result := FList.Add(Value);
end;

function TJsonArray.Add(const Value: JsonString): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.S := Value;
end;

function TJsonArray.Add(const Value: IJsonArray): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.A := Value;
end;

function TJsonArray.Add(const Value: TJsonMethod): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  Self.Add.M := Value;
end;

function TJsonArray.AddSA(const Value: JsonString): IJsonValue;
begin
  ParseJosnIfNone;
  Result := _SV(Value);
  Add(Result);
end;

function TJsonArray.AddSO(const Value: JsonString): IJsonValue;
begin
  ParseJosnIfNone;
  Result := _SV(Value);
  Add(Result);
end;

procedure TJsonArray.Assign(const Source: IJsonBase);
var
  Src: TJsonArray;
  I: Integer;
begin
  Clear;
  if not (TObject(Source.Implementor) is TJsonArray) then
    RaiseAssignError(TObject(Source.Implementor));
  Src := TJsonArray(Source.Implementor);
  for I := 0 to Src.Length - 1 do Add.Assign(Src[I]);
end;

function TJsonArray.Call(const Index: Integer; const Param, This: IJsonValue): IJsonValue;
var
  LMethod: TJsonMethod;
begin
  ParseJosnIfNone;
  LMethod := Self.M[index];
  if @LMethod = nil then
    RaiseError(Format('Method [%d] not exist in JsonArray!', [index]));
  LMethod(This, Param, Result);
end;

function TJsonArray.Call(const Index: Integer;
  const param: JsonString; const This: IJsonValue): IJsonValue;
begin
  ParseJosnIfNone;
  Result := Call(Index, _SV(param), This)
end;

procedure TJsonArray.Clear;
begin
  FList.Clear;
  //FParseStr := EmptyStr;
  FParsed := True;
end;

constructor TJsonArray.Create(const AParseStr: PJsonChar; const AParseLen: Integer; const ACapacity: Integer);
begin
  inherited Create();
  FParseStr := AParseStr;
  FParseLen := AParseLen;
  FParsed   := AParseLen = 0;
  FList := TInterfaceList.Create;
  FList.Capacity := ACapacity;
end;

procedure TJsonArray.Delete(const Index: Integer);
begin
  ParseJosnIfNone;
  FList.Delete(Index);
end;

destructor TJsonArray.Destroy;
begin
  Clear;
  FList := nil;
  inherited;
end;

function TJsonArray.IndexOf(const Value, Key: JsonString; CaseSensitive: Boolean): Integer;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  for Result := 0 to GetLength - 1 do
  begin
    LValue := GetValueAt(Result);
    if _JsonArrayIndexOfCompare(LValue, Value, Key, CaseSensitive) = 0 then
      Exit;
  end;
  Result := -1;
end;

function TJsonArray.LastIndexOf(const Value, Key: JsonString; CaseSensitive: Boolean): Integer;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  for Result := GetLength - 1 downto 0 do
  begin
    LValue := GetValueAt(Result);
    if _JsonArrayIndexOfCompare(LValue, Value, Key, CaseSensitive) = 0 then
      Exit;
  end;
  Result := -1;
end;

function TJsonArray.GetA(const Index: Integer): IJsonArray;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).A;
end;

function TJsonArray.GetB(const Index: Integer): Boolean;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).B;
end;

function TJsonArray.GetCapacity: Integer;
begin
  Result := FList.Capacity;
end;

function TJsonArray.GetD(const Index: Integer): Double;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).D;
end;

function TJsonArray.GetI(const Index: Integer): Int64;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).I;
end;

function TJsonArray.GetLength: Integer;
begin
  ParseJosnIfNone;
  Result := FList.Count;
end;

function TJsonArray.GetM(const Index: Integer): TJsonMethod;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).M;
end;

function TJsonArray.GetO(const Index: Integer): IJsonObject;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).O;
end;

function TJsonArray.GetS(const Index: Integer): JsonString;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).S;
end;

function TJsonArray.GetR(const Index: Integer): JsonString;
begin
  ParseJosnIfNone;
  Result := GetValueAt(Index).R;
end;

function TJsonArray.GetValueAt(Index: Integer): IJsonValue;
begin
  ParseJosnIfNone;
  Result := FList[Index] as IJsonValue;
end;

function TJsonArray.GetValues(const KeyName, KeyValue: JsonString): IJsonValue;
var
  i: Integer;
begin
  ParseJosnIfNone;
  i := IndexOf(KeyValue, KeyName);
  if i < 0 then
    Result := nil
  else
    Result := GetValueAt(i)
end;

procedure TJsonArray.Insert(const Index: Integer; const Value: IJsonValue);
begin
  ParseJosnIfNone;
  FList.Insert(Index, Value);
end;

procedure TJsonArray.Merge(const Addition: IJsonArray);
var
  I: Integer;
begin
  ParseJosnIfNone;
  if Addition = nil then Exit;
  for I := 0 to Addition.Length - 1 do
    Add(Addition[I]);
end;


procedure TJsonArray._ParseFind(const PaserStr: PJsonChar;
  const AIndex: Integer; const ASplitStr: JsonString;
  const AParsing: Boolean);
var
  LItem: IJsonValue;
begin
  if AParsing then
  begin
    LItem := TJsonValue.Create(PJsonChar(ASplitStr), System.Length(ASplitStr));
    FList.Add(LItem);
  end;
end;

procedure TJsonArray._Parse(const JsonStr: PJsonChar; const Len: Integer);
var
  PB, PE: PJsonChar;
begin
  try
    Clear;
    if not _IsJsonArray(JsonStr, Len) then RaiseParseError(JsonStr);
    PB := @JsonStr[1];
    PE := @JsonStr[Len - 1];
    if PB >= PE then
      Exit;
    SplitP(PB, PE, ',', _ParseFind);
  finally
    FParsed   := True;
    FParseStr := '';
  end;
end;

procedure TJsonArray.ParseJosnIfNone;
begin
  if (not FParsed) and (FParseStr <> EmptyStr) then
    _Parse(PJsonChar(FParseStr), FParseLen);
end;

procedure TJsonArray.PutA(const Index: Integer; const Value: IJsonArray);
begin
  ParseJosnIfNone;
  GetValueAt(Index).A := Value;
end;

procedure TJsonArray.PutB(const Index: Integer; const Value: Boolean);
begin
  ParseJosnIfNone;
  GetValueAt(Index).B := Value;
end;

procedure TJsonArray.PutD(const Index: Integer; const Value: Double);
begin
  ParseJosnIfNone;
  GetValueAt(Index).D := Value;
end;

procedure TJsonArray.PutI(const Index: Integer; const Value: Int64);
begin
  ParseJosnIfNone;
  GetValueAt(Index).I := Value;
end;

procedure TJsonArray.PutM(const Index: Integer; const Value: TJsonMethod);
begin
  ParseJosnIfNone;
  GetValueAt(Index).M := Value;
end;

procedure TJsonArray.PutO(const Index: Integer; const Value: IJsonObject);
begin
  ParseJosnIfNone;
  GetValueAt(Index).O := Value;
end;

procedure TJsonArray.PutS(const Index: Integer; const Value: JsonString);
begin
  ParseJosnIfNone;
  GetValueAt(Index).S := Value;
end;

procedure TJsonArray.PutR(const Index: Integer; const Value: JsonString);
begin
  ParseJosnIfNone;
  GetValueAt(Index).R := Value;
end;

function TJsonArray._SaveToStream(const AStream: IStream): Integer;
var
  I: Integer;
  Item: TJsonValue;
begin
  Result := 0;
  if (FIndentOption = ioNone) and (not FParsed) and (FParseStr <> EmptyStr) then
  begin
    Result := Result + _WriteStream(AStream, FParseStr);
    Exit;
  end;
  ParseJosnIfNone;
  if (FIndentOption = ioIndent) and (FIndent > 0) then
    Result := Result + _WriteStream(AStream, _IndentSpace(FIndent));
  Result := Result + _WriteStream(AStream, '[');

  for I := 0 to FList.Count - 1 do
  begin
    if I > 0 then
      Result := Result + _WriteStream(AStream, ',');
    if FIndentOption = ioIndent then
      Result := Result + _WriteStream(AStream, sLineBreak);
    Item := TJsonValue((FList[I] as IJsonValue).Implementor);
    Item.IndentOption := FIndentOption;
    Item.Indent := FIndent + JsonIndentSize;
    Result := Result + Item._SaveToStream(AStream);
    Item.Indent := 0;
  end;
  
  if (FIndentOption = ioIndent) and (FList.Count > 0) then
    Result := Result + _WriteStream(AStream, sLineBreak);
  if (FIndentOption = ioIndent) and (FIndent > 0) and (FList.Count > 0) then
    Result := Result + _WriteStream(AStream, _IndentSpace(FIndent));
  Result := Result + _WriteStream(AStream, ']');
end;

procedure TJsonArray.SetCapacity(const Value: Integer);
begin
  FList.Capacity := Value;
end;

procedure TJsonArray.SetValueAt(Index: Integer; const Value: IJsonValue);
begin
  ParseJosnIfNone;
  GetValueAt(Index).Assign(Value);
end;

procedure TJsonArray.SetValues(const KeyName, KeyValue: JsonString; const Value: IJsonValue);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(KeyValue, KeyName);
  if idx >= 0 then
    SetValueAt(idx, Value);
end;

function TJsonArray.IsEmpty: Boolean;
begin
  Result := GetLength <= 0;
end;

function TJsonArray.Delete(const Value, Key: JsonString; CaseSensitive: Boolean): Integer;
begin
  Result := IndexOf(Value, Key, CaseSensitive);
  if Result >= 0 then
    Delete(Result);
end;

function TJsonArray.Join(const separator, key: JsonString): JsonString;
var
  i: Integer;
begin
  for i := 0 to Self.Length - 1 do
  begin
    if Result <> '' then Result := Result + separator; 
    if key = '' then
      Result := Result + GetS(i)
    else
      Result := Result + getO(i).Path[key].S
  end;
end;

procedure TJsonArray.Exchange(const Index1, Index2: Integer);
begin
  ParseJosnIfNone;
  FList.Exchange(Index1, Index2);
end;

procedure TJsonArray._QuickSort(L, R: Integer; SCompare: TJsonArraySortCompare; const Key: JsonString; CaseSensitive: Boolean);
var
  I, J, P: Integer;
begin
  repeat
    I := L;
    J := R;
    P := (L + R) shr 1;
    repeat
      while SCompare(ValueAt[I], ValueAt[P], Key, CaseSensitive) < 0 do Inc(I);
      while SCompare(ValueAt[J], ValueAt[P], Key, CaseSensitive) > 0 do Dec(J);
      if I <= J then
      begin
        Exchange(I, J);
        if P = I then
          P := J
        else if P = J then
          P := I;
        Inc(I);
        Dec(J);
      end;
    until I > J;
    if L < J then _QuickSort(L, J, SCompare, Key, CaseSensitive);
    L := I;
  until I >= R;
end;

procedure TJsonArray.Sort(Compare: TJsonArraySortCompare; const Key: JsonString; CaseSensitive: Boolean);
begin
  if Self.Length <= 0 then Exit;
  if @Compare = nil then
    _QuickSort(0, Self.Length - 1, _JsonArrayCompare, Key, CaseSensitive)
  else
    _QuickSort(0, Self.Length - 1, Compare, Key, CaseSensitive)
end;

procedure TJsonArray.Reverse;
var
  i, i1, i2, iHigh: Integer;
begin
  iHigh := Self.Length - 1;
  for i := 0 to iHigh do
  begin
    i1 := i;
    i2 := iHigh-i;
    if i2 <= i1 then Break;
    Exchange(i1, i2);
  end;
end;

function TJsonArray.Slice(startPos, endPos: Integer): IJsonArray;
var
  i: Integer;
begin
  Result := _SA();
  if Self.Length <= 0 then Exit;
  if startPos = 0 then Exit;
  if endPos = 0 then endPos := Self.Length;
  if startPos < 0 then startPos := Self.Length + startPos + 1;
  if endPos < 0 then endPos := Self.Length + endPos + 1;
  if startPos > Self.Length then Exit;
  if endPos > Self.Length then endPos := Self.Length;
  for i := startPos-1 to endPos-1 do
    Result.Add(Self.ValueAt[i]);
end;

function TJsonArray.First: IJsonValue;
begin
  if Self.Length > 0 then
    Result := Self.ValueAt[0]
  else
    Result := nil;
end;

function TJsonArray.Last: IJsonValue;
begin
  if Self.Length > 0 then
    Result := Self.ValueAt[Self.Length-1]
  else
    Result := nil;
end;

{ TJsonPairList }

procedure TJsonPairList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if (Action = lnDeleted) and (Ptr<>nil) then
    TJsonPair(Ptr).Free;
end;


{ TJsonObject }

function TJsonObject.Add(const Key: JsonString): IJsonValue;
begin
  ParseJosnIfNone;
  Result := AddPair(Key).Value;
end;

function TJsonObject.AddPair(const Key: JsonString): TJsonPair;
begin
  ParseJosnIfNone;
  Result := TJsonPair.Create(Self, Key);
  FListHash.Add(Key, FList.Add(Result));
end;

function TJsonObject.AddPair(const APair: TJsonPair): Integer;
begin
  ParseJosnIfNone;
  Result := FList.Add(APair);
  FListHash.Add(APair.Name, Result);
end;

procedure TJsonObject.Assign(const Source: IJsonBase);
var
  Src: TJsonObject;
  I: Integer;
begin
  Clear;
  if not (TObject(Source.Implementor) is TJsonObject) then
    RaiseAssignError(TObject(Source.Implementor));

  Src := TJsonObject(Source.Implementor);
  for I := 0 to Src.Length - 1 do
    AddPair(Src.Items[I].Name).Assign(Src.Items[I]);
end;

procedure TJsonObject.Clear;
begin
  if FList = nil then
    Exit;
  FList.Clear;
  //FParseStr := EmptyStr;
  FParsed := True;
end;

constructor TJsonObject.Create(AOwner: TJsonBase; const AParseStr: PJsonChar; const AParseLen: Integer;
  const ACapacity: Integer);
begin
  inherited Create();
  FParseStr := AParseStr;
  FParseLen := AParseLen;
  FParsed   := AParseLen = 0;
  FCapacity := ACapacity;
  FListHash := TJsonHash.Create(FCapacity);
  FHashValid := True;
  
  FList := TJsonPairList.Create;
  FAutoAdd := True;
end;

procedure TJsonObject.Delete(const Index: Integer);
begin
  ParseJosnIfNone;
  FList.Delete(Index);
  FHashValid := False;
end;

function TJsonObject.Delete(const Key: JsonString): Integer;
begin
  ParseJosnIfNone;
  Result := IndexOf(Key);
  if Result >= 0 then
    Delete(Result);
end;

destructor TJsonObject.Destroy;
begin
  Clear;
  FListHash := nil;
  if FList <> nil then
    FreeAndNil(FList);
  inherited Destroy;
end;

function TJsonObject.IndexOf(const Key: JsonString): Integer;
begin
  ParseJosnIfNone;
  UpdateHash;
  Result := FListHash.ValueOf(Key);
end;

function TJsonObject.GetA(const Key: JsonString): IJsonArray;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.A
  else
  if AutoAdd then
    Result := AddPair(Key).Value.A
  else
    Result := nil;
end;

function TJsonObject.GetAutoAdd: Boolean;
begin
  Result := FAutoAdd;
end;

function TJsonObject.GetB(const Key: JsonString): Boolean;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.B
  else
    Result := False;
end;

function TJsonObject.GetCapacity: Integer;
begin
  Result := FCapacity;
end;

function TJsonObject.GetD(const Key: JsonString): Double;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.D
  else
    Result := 0.0;
end;

function TJsonObject.GetI(const Key: JsonString): Int64;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.I
  else
    Result := 0;
end;

function TJsonObject.GetItems(Index: Integer): TJsonPair;
begin
  ParseJosnIfNone;
  Result := TJsonPair(FList[Index]);
end;

function TJsonObject.GetLength: Integer;
begin
  ParseJosnIfNone;
  Result := FList.Count;
end;

function TJsonObject.GetO(const Key: JsonString): IJsonObject;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.O
  else
  if AutoAdd then
    Result := AddPair(Key).Value.O
  else
    Result := nil;
end;

function TJsonObject.GetS(const Key: JsonString): JsonString;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.S
  else
    Result := '';
end;

function TJsonObject.GetR(const Key: JsonString): JsonString;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.R
  else
    Result := '';
end;

function TJsonObject.GetValueAt(const Index: Integer): IJsonValue;
var
  Pair: TJsonPair;
begin
  ParseJosnIfNone;
  Pair := TJsonPair(FList[Index]);
  Result := Pair.Value;
end;

function TJsonObject.GetValues(Name: JsonString): IJsonValue;
var
  Index: Integer;
  Pair: TJsonPair;
begin
  ParseJosnIfNone;
  Index := IndexOf(Name);
  if Index < 0 then
  begin
    Result := nil;
    Exit;
  end;
  Pair := TJsonPair(FList[Index]);
  Result := Pair.Value;
end;

function TJsonObject.Insert(const Index: Integer; const Key: JsonString): IJsonValue;
var
  Pair: TJsonPair;
begin
  ParseJosnIfNone;
  Pair := TJsonPair.Create(Self, Key);
  FList.Insert(Index, Pair);
  Result := Pair.Value;
  FHashValid := False;
end;

procedure TJsonObject.Merge(const Addition: IJsonObject);
var
  I, idx: Integer;
  o: TJsonObject;
  v: IJsonValue;
begin
  ParseJosnIfNone;
  if Addition = nil then Exit;
  o := TJsonObject(Addition.Implementor);
  for I := 0 to Addition.Length - 1 do
  begin
    idx := IndexOf(o.Items[I].Name);
    if idx >= 0 then
    begin
      v := Items[idx].Value;
      case v.ValueType of
        jvObject:
        begin
          if o.Items[I].Value.ValueType = jvObject then
            v.O.Merge(o.Items[I].Value.O)
          else
            v.Assign(o.Items[I].Value);
        end;
        jvArray:
        begin
          if o.Items[I].Value.ValueType = jvArray then
            v.A.Merge(o.Items[I].Value.A)
          else
            v.Assign(o.Items[I].Value);
        end;
      else
        v.Assign(o.Items[I].Value);
      end;
    end
    else
      AddPair(o.Items[I].Name).Assign(o.Items[I]);
  end;
end;

procedure TJsonObject._ParseFind(const PaserStr: PJsonChar; const AIndex: Integer; 
  const ASplitStr: JsonString; const AParsing: Boolean);
begin
  if AParsing then
    AddPair(TJsonPair.CreateForParse(Self, PJsonChar(ASplitStr), System.Length(ASplitStr)));
end;

procedure TJsonObject._Parse(const JsonStr: PJsonChar; const Len: Integer);
var
  PB, PE: PJsonChar;
begin
  try
    Clear;
    if not _IsJsonObject(JsonStr, Len) then
      RaiseParseError(JsonStr);
    PB := @JsonStr[1];
    PE := @JsonStr[Len - 1];
    if PB >= PE then
      Exit;
    SplitP(PB, PE, ',', _ParseFind);
  finally
    FParsed := True;
    FParseStr := '';
  end;
end;

function TJsonObject.Add(const Key: JsonString; const Value: IJsonObject): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.O := Value;
end;

function TJsonObject.Add(const Key: JsonString; const Value: IJsonArray): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.A := Value;
end;

function TJsonObject.Add(const Key: JsonString; const Value: IJsonValue): Integer;
var
  LPair: TJsonPair;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  LPair := AddPair(Key);
  if Value <> nil then
    LPair.FValue := Value;
end;

function TJsonObject.AddSA(const Key, Value: JsonString): IJsonValue;
begin
  ParseJosnIfNone;
  Result := AddPair(Key).Value;
  Result.A.Parse(Value);
end;

function TJsonObject.AddSO(const Key, Value: JsonString): IJsonValue;
begin
  ParseJosnIfNone;
  Result := AddPair(Key).Value;
  Result.O.Parse(Value);
end;

function TJsonObject.Add(const Key, Value: JsonString): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.S := Value;
end;

function TJsonObject.Add(const Key: JsonString; const Value: Int64): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.I := Value;
end;

function TJsonObject.Add(const Key: JsonString; const Value: Boolean): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.B := Value;
end;

function TJsonObject.Add(const Key: JsonString; const Value: Double): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.D := Value;
end;

procedure TJsonObject.PutA(const Key: JsonString; const Value: IJsonArray);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).A := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.PutB(const Key: JsonString; const Value: Boolean);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).B := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.PutD(const Key: JsonString; const Value: Double);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).D := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.PutI(const Key: JsonString; const Value: Int64);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).I := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.PutO(const Key: JsonString; const Value: IJsonObject);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).O := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.PutS(const Key, Value: JsonString);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).S := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.PutR(const Key, Value: JsonString);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).R := Value
  else
  if FAutoAdd then
    AddPair(Key).Value.R := Value;
end;

function TJsonObject._SaveToStream(const AStream: IStream): Integer;
var
  I: Integer;
  Item: TJsonPair;
begin
  Result := 0;
  if (FIndentOption = ioNone) and (not FParsed) and (FParseStr <> EmptyStr) then
  begin
    Result := Result + _WriteStream(AStream, FParseStr);
    Exit;
  end;
  ParseJosnIfNone;
  if (FIndentOption = ioIndent) and (FIndent > 0) then
    Result := Result + _WriteStream(AStream, _IndentSpace(FIndent));

  Result := Result + _WriteStream(AStream, '{');
  for I := 0 to FList.Count - 1 do
  begin
    if I > 0 then
      Result := Result + _WriteStream(AStream, ',');
    if (FIndentOption = ioIndent) then
      Result := Result + _WriteStream(AStream, sLineBreak);

    Item := TJsonPair(FList[I]);
    Item.IndentOption := FIndentOption;
    Item.Indent := FIndent + JsonIndentSize;

    Result := Result + Item._SaveToStream(AStream);
    Item.Indent := 0;
  end;

  if (FIndentOption = ioIndent) and (FList.Count > 0) then
    Result := Result + _WriteStream(AStream, sLineBreak);
  if (FIndentOption = ioIndent) and (FIndent > 0) and (FList.Count > 0) then
    Result := Result + _WriteStream(AStream, _IndentSpace(FIndent));

  Result := Result + _WriteStream(AStream, '}');
end;

procedure TJsonObject.SetAutoAdd(const Value: Boolean);
begin
  FAutoAdd := Value;
end;

procedure TJsonObject.SetCapacity(const Value: Integer);
begin
  if Value = FCapacity then
    Exit;
  if Value <= 64 then
    Exit;
  FListHash := TJsonHash.Create(FCapacity);
  FList.Capacity := Value;
end;

procedure TJsonObject.SetValueAt(const Index: Integer; const Value: IJsonValue);
begin
  ParseJosnIfNone;
  GetValueAt(Index).Assign(Value);
end;

procedure TJsonObject.SetValues(Key: JsonString; const Value: IJsonValue);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).Assign(Value)
  else
  if FAutoAdd then
    Add(Key, Value);
end;

procedure TJsonObject.UpdateHash;
var
  I: Integer;
begin
  if FHashValid then Exit;
  FListHash.Clear;
  for I := 0 to GetLength - 1 do
    FListHash.Add(TJsonPair(FList[I]).Name, I);
  FHashValid := True;
end;

function TJsonObject.GetM(const Key: JsonString): TJsonMethod;
var
  LValue: IJsonValue;
begin
  ParseJosnIfNone;
  LValue := GetValues(Key);
  if LValue <> nil then
    Result := LValue.M
  else
    Result := nil;
end;

procedure TJsonObject.PutM(const Key: JsonString; const Value: TJsonMethod);
var
  idx: Integer;
begin
  ParseJosnIfNone;
  idx := IndexOf(Key);
  if idx >= 0 then
    GetValueAt(idx).M := Value
  else
  if FAutoAdd then
    Add(Key, Value);
end;

function TJsonObject.Add(const Key: JsonString; const Value: TJsonMethod): Integer;
begin
  ParseJosnIfNone;
  Result := Self.Length;
  AddPair(Key).Value.M := Value;
end;

function TJsonObject.Call(const Key: JsonString; const Param, This: IJsonValue): IJsonValue;
var
  LMethod: TJsonMethod;   
begin
  ParseJosnIfNone;
  LMethod := Self.M[Key];
  if @LMethod = nil then
    RaiseError(SysUtils.Format('Method [%s] not exist in JsonObject!', [Key]));
  LMethod(This, Param, Result);
end;

function TJsonObject.Call(const Key, Param: JsonString; const This: IJsonValue): IJsonValue;
begin
  ParseJosnIfNone;
  Result := Call(Key, _SV(Param), This)
end;

procedure TJsonObject.ParseJosnIfNone;
begin
  if (not FParsed) and (FParseStr <> EmptyStr) then
   _Parse(PJsonChar(FParseStr), FParseLen);
end;

function TJsonObject.GetOnParseObjectProp: TParseObjectProp;
begin
  Result := FOnParseObjectProp;
end;

procedure TJsonObject.SetOnParseObjectProp(const Value: TParseObjectProp);
begin
  FOnParseObjectProp := Value;
end;

function TJsonObject.GetNameAt(const Index: Integer): JsonString;
var
  Pair: TJsonPair;
begin
  ParseJosnIfNone;
  Pair := TJsonPair(FList[Index]);
  Result := Pair.Name;
end;

function TJsonObject.Format(const AStr: JsonString; BeginSep, EndSep: JsonChar): JsonString;
var
  p1, p2: PJsonChar;
begin
  Result := '';
  p2 := PJsonChar(AStr);
  p1 := p2;
  while true do
  begin
    if p2^ = BeginSep then
    begin
      if p2 > p1 then
        Result := Result + Copy(p1, 0, p2-p1);
      inc(p2);
      p1 := p2;
      while true do
        if p2^ = EndSep then
          Break
        else
        if p2^ = #0 then
          Exit
        else
          inc(p2);
      Result := Result + GetS(copy(p1, 0, p2-p1));
      inc(p2);
      p1 := p2;
    end
    else
    if p2^ = #0 then
    begin
      if p2 > p1 then
        Result := Result + Copy(p1, 0, p2-p1);
      Break;
    end
    else
      inc(p2);  
  end;
end;

function TJsonObject.GetPath(Path: JsonString): IJsonValue;
var
  i, iLen: Integer;
  jObj: IJsonObject;
  List: array of JsonString;
begin
  ParseJosnIfNone;
  i := Pos('.', Path);
  iLen := 0;
  while i > 0 do
  begin
    iLen := iLen + 1;
    SetLength(List, iLen);
    List[iLen-1] := Copy(Path, 1, i-1);
    System.Delete(Path, 1, i);
    i := Pos('.', Path);
  end;
  jObj := Self;
  for i := 0 to iLen-1 do
  begin
    jObj := jObj.O[List[i]];
    if jObj = nil then
      jObj := jObj.Add(List[i]).O;
  end;
  if Path <> '' then
  begin
    Result := jObj.Values[Path];
    if Result = nil then
      Result := jObj.Add(Path);
  end;
end;

procedure TJsonObject.PutPath(Path: JsonString; const Value: IJsonValue);
var
  i, iLen: Integer;
  jObj: IJsonObject;
  List: array of JsonString;
begin
  ParseJosnIfNone;
  i := Pos('.', Path);
  iLen := 0;
  while i > 0 do
  begin
    iLen := iLen + 1;
    SetLength(List, iLen);
    List[iLen-1] := Copy(Path, 1, i-1);
    System.Delete(Path, 1, i);
    i := Pos('.', Path);
  end;
  jObj := Self;
  for i := 0 to iLen-1 do
  begin
    jObj := jObj.O[List[i]];
    if jObj = nil then
      jObj := jObj.Add(List[i]).O;
  end;
  if Path <> '' then
   jObj.Values[Path] := Value;
end;

{ TJsonHash }

procedure TJsonHash.Add(const Key: JsonString; Value: Integer);
var
  Hash: Integer;
  Bucket: PJsonHashItem;
begin
  Hash := HashOf(PJsonChar(Key), Length(Key)) mod Cardinal(Length(Buckets));
  New(Bucket);
  Bucket^.Key := Key;
  Bucket^.Value := Value;
  Bucket^.Next := Buckets[Hash];
  Buckets[Hash] := Bucket;
end;

procedure TJsonHash.Clear;
var
  I: Integer;
  P, N: PJsonHashItem;
begin
  for I := 0 to Length(Buckets) - 1 do
  begin
    P := Buckets[I];
    while P <> nil do
    begin
      N := P^.Next;
      Dispose(P);
      P := N;
    end;
    Buckets[I] := nil;
  end;
end;

constructor TJsonHash.Create(Size: Cardinal);
begin
  inherited Create;
  SetLength(Buckets, Size);
end;

destructor TJsonHash.Destroy;
begin
  Clear;
  inherited;
end;

function TJsonHash.Find(const Key: JsonString): PPJsonHashItem;
var
  Hash: Integer;
begin
  Hash := HashOf(PJsonChar(Key), Length(Key)) mod Cardinal(Length(Buckets));
  Result := @Buckets[Hash];
  while Result^ <> nil do
  begin
    if Result^.Key = Key then
      Exit;
    Result := @Result^.Next;
  end;
end;

function TJsonHash.HashOf(const Key: PJsonChar; const Len: Cardinal): Cardinal;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to Len-1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(Key[I]);
end;

function TJsonHash.Modify(const Key: JsonString; Value: Integer): Boolean;
var
  P: PJsonHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
  begin
    Result := True;
    P^.Value := Value;
  end
  else
    Result := False;
end;

procedure TJsonHash.Remove(const Key: JsonString);
var
  P: PJsonHashItem;
  Prev: PPJsonHashItem;
begin
  Prev := Find(Key);
  P := Prev^;
  if P <> nil then
  begin
    Prev^ := P^.Next;
    Dispose(P);
  end;
end;

function TJsonHash.ValueOf(const Key: JsonString): Integer;
var
  P: PJsonHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
    Result := P^.Value
  else
    Result := -1;
end;

end.

