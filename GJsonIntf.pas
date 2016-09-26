unit GJsonIntf;

interface

{$DEFINE UNICODE}

uses
  ActiveX;

type
  TJsonValueType = (jvNull, jvString, jvRawString, jvInteger, jvDouble, jvBoolean,
    jvObject, jvArray, jvMethod);
    
  {$IFDEF UNICODE}
  JsonString = WideString;
  JsonChar   = WideChar;
  PJsonChar  = PWideChar;
  {$ELSE}
  JsonString = AnsiString;
  JsonChar   = AnsiChar;
  PJsonChar  = PAnsiChar;
  {$ENDIF}

  PPJsonHashItem = ^PJsonHashItem;
  PJsonHashItem = ^TJsonHashItem;
  TJsonHashItem = record
    Next: PJsonHashItem;
    Key: JsonString;
    Value: Integer;
  end;

const
  JsonCharSize   = SizeOf(JsonChar);
  JsonIndentSize =  2;

type
  IJsonBase = interface;
  IJsonArray = interface;
  IJsonObject = interface;
  IJsonValue = interface;
  
  TJsonMethod = procedure(const This, Params: IJsonValue; var Result: IJsonValue) of object;

  IJsonHash = interface
  ['{AE758EC1-6806-4F0C-9300-5C021E67F577}']
    function Find(const Key: JsonString): PPJsonHashItem;
    function HashOf(const Key: PJsonChar; const Len: Cardinal): Cardinal; 

    procedure Add(const Key: JsonString; Value: Integer);
    procedure Clear;
    procedure Remove(const Key: JsonString);
    function Modify(const Key: JsonString; Value: Integer): Boolean;
    function ValueOf(const Key: JsonString): Integer;
  end;

  TIndentOption = (ioNone, ioIndent, ioCompact);
  TFindSplitStrEvent = procedure (const PaserStr: PJsonChar; const AIndex: Integer;
    const ASplitStr: JsonString; const AParsing: Boolean) of object;

  IJsonBase = interface
  ['{B2BCABB5-5D80-47F9-9103-1C058C11777A}']
    function GetIndent: Integer;
    function GetIndentOption: TIndentOption;
    function GetTag: Pointer;
    procedure SetIndent(const Value: Integer);
    procedure SetIndentOption(const Value: TIndentOption);
    procedure SetTag(const Value: Pointer);

    function Implementor: Pointer;

    procedure Parse(const JsonString: JsonString);
    procedure ParseStream(AStream: IStream; const Utf8: Boolean = False);
    procedure ParseFile(const AFileName: JsonString; const Utf8: Boolean = False);

    function  Stringify: JsonString;
    function  SaveToStream(AStream: IStream): Integer;
    procedure SaveToFile(const AFileName: JsonString; const AEncode: JsonString = '');

    procedure Assign(const Source: IJsonBase);
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

    property IndentOption: TIndentOption read GetIndentOption write SetIndentOption;
    property Indent: Integer read GetIndent write SetIndent;
    property Tag: Pointer read GetTag write SetTag;
  end;

  TJsonArraySortCompare = function(const Value1, Value2: IJsonValue; const Key: JsonString; CaseSensitive: Boolean): Integer;

  IJsonArray = interface(IJsonBase)
  ['{BB1D9B8A-A282-4605-ACB5-38B5180B12FA}']
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
    property R[const Index: Integer]: JsonString read GetS write PutR;
    property M[const Index: Integer]: TJsonMethod read GetM write PutM;
  end;

  TParseObjectProp = procedure (const AObject: IJsonObject;
    const APropName: JsonString; const APropValue: IJsonValue) of object;

  IJsonObject = interface(IJsonBase)
  ['{D0D87EC4-3FBE-4DD8-AD0B-E373BC424E84}']
    function GetLength: Integer;
    function GetCapacity: Integer;
    function GetValues(Key: JsonString): IJsonValue;
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

    function  Format(const AStr: JsonString; BeginSep: JsonChar = '{'; EndSep: JsonChar = '}'): JsonString;
    
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
    function  Insert(const Index: Integer; const Key: JsonString): IJsonValue;
    procedure Merge(const Addition: IJsonObject);
    function  IndexOf(const Key: JsonString): Integer;
    procedure Delete(const Index: Integer); overload;
    function  Delete(const Key: JsonString): Integer; overload;
    procedure Clear;

    function Call(const Key: JsonString; const Param: IJsonValue = nil; const This: IJsonValue = nil): IJsonValue; overload;
    function Call(const Key: JsonString; const Param: JsonString; const This: IJsonValue = nil): IJsonValue; overload;

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
    property Path[Path: JsonString]: IJsonValue read GetPath write PutPath;

    property OnParseObjectProp: TParseObjectProp read GetOnParseObjectProp write SetOnParseObjectProp;
  end;

  IJsonValue = interface(IJsonBase)
  ['{D435408F-C5DE-4188-AD8C-D28C542B78D5}']
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
    
    procedure Clear;
    function  Clone: IJsonValue;

    function IsType(const AType: TJsonValueType): Boolean;

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

function SV(const AStr: JsonString = '{}'): IJsonValue;
function SO(const AStr: JsonString = '{}'): IJsonObject;
function SA(const AStr: JsonString = '[]'): IJsonArray;
function SVFile(const AFileName: JsonString): IJsonValue;
function SOFile(const AFileName: JsonString): IJsonObject;
function SAFile(const AFileName: JsonString): IJsonArray;
function SASplit(const Str: JsonString; const separator: JsonChar; howmany: Integer = 0): IJsonArray;

function StrToUnicode(const str: WideString): WideString;

implementation

uses
  GJsonImpl;

function SV(const AStr: JsonString): IJsonValue;
begin
  Result := _SV(AStr)
end;

function SO(const AStr: JsonString): IJsonObject;
begin
  Result := _SO(AStr)
end;

function SA(const AStr: JsonString): IJsonArray;
begin
  Result := _SA(AStr)
end;

function SVFile(const AFileName: JsonString): IJsonValue;
begin
  Result := _SVFile(AFileName)
end;

function SOFile(const AFileName: JsonString): IJsonObject;
begin
  Result := _SVFile(AFileName).O;
end;

function SAFile(const AFileName: JsonString): IJsonArray;
begin
  Result := _SVFile(AFileName).A;
end;

function SASplit(const Str: JsonString; const separator: JsonChar; howmany: Integer = 0): IJsonArray;
begin
  Result := _SASplit(Str, separator, howmany);
end;

function StrToUnicode(const str: WideString): WideString;
begin
  Result := _StrToUnicode(str);
end;

end.
