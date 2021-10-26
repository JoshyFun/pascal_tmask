unit umask;

(*

  Fixed autoreverse compilation.
  Fixed empty negated group [!]
  Fixed escapechar in groups [\x], range and optional char
  TMaskException renamed to EMaskException to match Delphi's exception name.
  EMaskException.cCode renamed to FCode (It is a property)
  TMaskBase.cCaseSensitive renamed to FCaseSensitive (It is a property)
  TMaskBase.cMaskEscapeChar renamed to FMaskEscapeChar (property)
  TMaskBase.cMaskOpcodesAllowed renamed to FMaskOpcodesAllowed (property)
  cOriginalMask (3 classes) renamed to FOriginalMask (property)
  cWindowsMask (3 classes) renamed to FWindowsMask (property)
  cMaskWindowsQuirkAllowed (3 classes) renamed to FMaskWindowsQuirkAllowed (property)
  Implemented UTF8 branchless CodePoint length.

  Added:
    TMaskDelphiUnicode=class(TMaskUnicode)
    TMaskDelphiUTF8=class(TMaskUTF8)
  As sample of initialization.

  ---

  Fixed a bug when OpcodeRanges is enabled but OptionalChar is NOT, the char
  will be always added as an OptionalChar. Now a InvalidCharMask exception is raised.

  Made TMaskOpCode, TMaskOpcodesEnum and TMaskOpcodesSet class public.

*)

{$mode objfpc}{$H+}

// RANGES_AUTOREVERSE
// It reverse ranges if needed, so range "[z-a]"
// is interpreted as "[a-z]"
{$DEFINE RANGES_AUTOREVERSE}

{$DEFINE USE_INLINE}
{$DEFINE USE_BRANCHLESS_UTF8LEN}

interface

uses
  Classes, SysUtils;

type

  { EMaskException }

  EMaskException=class(Exception)
  public
    type
      TMaskExceptionCode=(eMaskException_InternalError,
                          eMaskException_InvalidCharMask,
                          eMaskException_MissingClose,
                          eMaskException_IncompleteMask,
                          eMaskException_InvalidEscapeChar,
                          eMaskException_InvalidUTF8Sequence
                          );
  protected
    FCode: TMaskExceptionCode;
  public
    constructor Create(const msg: string; const aCode: TMaskExceptionCode);
    constructor CreateFmt(const msg: string; const args: array of const; const aCode: TMaskExceptionCode);
    property Code: TMaskExceptionCode read FCode;
  end;

  { TMaskBase }

  TMaskBase = class
  private
    procedure SetMaskEscapeChar(AValue: Char);
  public
    type
      // Literal = It must match
      // Range = Match any char in the range
      // Negate = Negate match in a group
      // AnyChar = It matches any char, but one must match
      // AnyCharOrNone = Matches one or none char (only in a group)
      // AnyCharToNext = Matches any chars amount, if fail, restart in the
      //                 next position up to finish the mask or the matched string
      // OptionalChar = Optional char
      // CharsGroupBegin = Begin optional chars or ranges "["
      // CharsGroupEnd = End optional chars or ranges "]"
      TMaskOpCode = (
        Literal=0,
        Range=1,
        Negate=2,
        AnyChar=3,
        AnyCharOrNone=4,
        AnyCharToNext=5,
        OptionalChar=6,
        CharsGroupBegin=10,
        CharsGroupEnd=11
      );
      TMaskOpcodesEnum=(eMaskOpcodeAnyChar,
                        eMaskOpcodeAnyCharOrNone,
                        eMaskOpcodeAnyText,
                        eMaskOpcodeRange,
                        eMaskOpcodeOptionalChar,
                        eMaskOpcodeNegateGroup,
                        eMaskOpcodeEscapeChar);
      TMaskOpcodesSet=set of TMaskOpcodesEnum;
  protected
    type
      TMaskFailCause = (
        Success = 0,
        MatchStringExhausted = 1,
        MaskExhausted = 2,
        MaskNotMatch = 3,
        UnexpectedEnd = 4
      );
      (*
        Windows mask works in a different mode than regular mask, it has too many
        quirks and corner cases inherited from CP/M, then adapted to DOS (8.3) file
        names and adapted again for long file names.

        Anyth?ng.abc    = "?" matches exactly 1 char
        Anyth*ng.abc    = "*" matches 0 or more of chars

        ------- Quirks -------

        --eWindowsQuirk_AnyExtension
          Anything*.*     = ".*" is removed.

        --eWindowsQuirk_FilenameEnd
          Anything??.abc  = "?" matches 1 or 0 chars (except '.')
                            (Not the same as "Anything*.abc", but the same
                            as regex "Anything.{0,2}\.abc")
                            Internally converted to "Anything[??].abc"

        --eWindowsQuirk_Extension3More
          Anything.abc    = Matches "Anything.abc" but also "Anything.abc*" (3 char extension)
          Anything.ab     = Matches "Anything.ab" and never "anything.abcd"

        --eWindowsQuirk_EmptyIsAny
          ""              = Empty string matches anything "*"

        --eWindowsQuirk_AllByExtension (Not in use anymore)
          .abc            = Runs as "*.abc"

        --eWindowsQuirk_NoExtension
          Anything*.      = Matches "Anything*" without extension
      *)
      TWindowsQuirks=(eWindowsQuirk_AnyExtension,eWindowsQuirk_FilenameEnd,eWindowsQuirk_Extension3More,eWindowsQuirk_EmptyIsAny,eWindowsQuirk_AllByExtension,eWindowsQuirk_NoExtension);
      TWindowsQuirkSet=set of TWindowsQuirks;
    const GROW_BY=100;
          TWindowsQuirksAllAllowed=[eWindowsQuirk_AnyExtension,
                                    eWindowsQuirk_FilenameEnd,
                                    eWindowsQuirk_Extension3More,
                                    eWindowsQuirk_EmptyIsAny,
                                    eWindowsQuirk_AllByExtension,
                                    eWindowsQuirk_NoExtension];
          TWindowsQuirksDefaultAllowed=[eWindowsQuirk_AnyExtension,
                                        eWindowsQuirk_FilenameEnd,
                                        eWindowsQuirk_Extension3More,
                                        eWindowsQuirk_EmptyIsAny,
                                        {eWindowsQuirk_AllByExtension,} // Not in use anymore
                                        eWindowsQuirk_NoExtension];
          TMaskOpCodesAllAllowed=[eMaskOpcodeAnyChar,
                                  eMaskOpcodeAnyCharOrNone,
                                  eMaskOpcodeAnyText,
                                  eMaskOpcodeRange,
                                  eMaskOpcodeOptionalChar,
                                  eMaskOpcodeNegateGroup,
                                  eMaskOpcodeEscapeChar];
          TMaskOpCodesDefaultAllowed=TMaskOpCodesAllAllowed;
  protected
    procedure Add(const aLength: integer; const aData: PBYTE);
    procedure Add(const aValue: integer);{$IFDEF USE_INLINE}inline;{$ENDIF}
    procedure Add(const aValue: TMaskOpCode);{$IFDEF USE_INLINE}inline;{$ENDIF}
    procedure IncrementLastCounterBy(const aOpcode: TMaskOpCode; const aValue: integer);
  protected
    FCaseSensitive: Boolean;
    cMaskIsCompiled: Boolean;
    cMaskCompiled: TBytes;
    cMaskCompiledIndex: integer;
    cMaskCompiledAllocated: integer;
    cMaskCompiledLimit: integer;
    cMaskLimit: integer;
    cMatchStringLimit: integer;
    cMatchMinimumLiteralBytes: SizeInt;
    cMatchMaximumLiteralBytes: SizeInt;
    FMaskOpcodesAllowed: TMaskOpcodesSet;
    // EscapeChar forces next char to be a literal one, not a wildcard.
    FMaskEscapeChar: Char;
    procedure Compile; virtual;
    class procedure Exception_InvalidCharMask(const aMaskChar: string; const aOffset: integer=-1); static;
    class procedure Exception_MissingCloseChar(const aMaskChar: string; const aOffset: integer=-1); static;
    class procedure Exception_IncompleteMask(); static;
    class procedure Exception_InvalidEscapeChar(); static;
    procedure Exception_InternalError();
    function intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause; virtual; abstract;
    property OPCodesAllowed: TMaskOpcodesSet read FMaskOpcodesAllowed write FMaskOpcodesAllowed;
  public
    constructor Create(const aCaseSensitive: Boolean=false);
    constructor CreateAdvanced(const aCaseSensitive: Boolean=false; const aOpcodesAllowed: TMaskOpcodesSet=TMaskOpCodesAllAllowed);
    property CaseSensitive: Boolean read FCaseSensitive;
    property EscapeChar: Char read FMaskEscapeChar write SetMaskEscapeChar;
  end;

  { TMaskANSI }

  TMaskANSI = class (TMaskBase)
  private
    cMatchString: RawByteString;
  protected
    FOriginalMask: RawByteString;
    function intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause; override;
  public
    constructor Create(const aMask: RawByteString; const aCaseSensitive: Boolean = False);
    constructor CreateAdvanced(const aMask: RawByteString; const aCaseSensitive: Boolean=false; const aOpcodesAllowed: TMaskOpcodesSet=TMaskOpCodesAllAllowed);
    procedure Compile; override;
    function AnsiOEMToUTF8(const aAnsiChar: char): RawByteString;
    function UTF8ToAnsiOEM(const aUTF8String: RawByteString): RawByteString;
    function Matches(const aStringToMatch: RawByteString): Boolean; virtual;
    property Mask: RawByteString read FOriginalMask write FOriginalMask;
    property OPCodesAllowed;
  end;

  { TMaskUTF8 }

  TMaskUTF8 = class (TMaskBase)
  private
    cMatchString: RawByteString;
  protected
    FOriginalMask: RawByteString;
    class function UTF8Length(const P: PBYTE): integer; static; {$IFDEF USE_INLINE}inline;{$ENDIF}
    class function CompareUTF8Sequences(const P1,P2: PBYTE): integer; static;{$IFDEF USE_INLINE}inline;{$ENDIF}
    function intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause; override;
  public
    constructor Create(const aMask: RawByteString; const aCaseSensitive: Boolean = False);
    constructor CreateAdvanced(const aMask: RawByteString; const aCaseSensitive: Boolean=false; const aOpcodesAllowed: TMaskOpcodesSet=TMaskOpCodesAllAllowed);
    procedure Compile; override;
    function Matches(const aStringToMatch: RawByteString): Boolean; virtual;
    property Mask: RawByteString read FOriginalMask write FOriginalMask;
    property OPCodesAllowed;
  end;

  { TMaskUnicode }

  TMaskUnicode = class (TMaskBase)
  private
    cMatchString: UnicodeString;
    const
      UTF16_CP_BYTES=2;
  protected
    FOriginalMask: UnicodeString;
    class function UTF16Length(const P: PWORD): integer; static; {$IFDEF USE_INLINE}inline;{$ENDIF}
    class function CompareUTF16Sequences(const P1,P2: PWORD): integer; static;{$IFDEF USE_INLINE}inline;{$ENDIF}
    class function InUnicodeChars(aUnicodeChar: UnicodeChar; aCharSet: array of UnicodeChar): Boolean; static;
    function intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause; override;
  public
    constructor Create(const aMask: UnicodeString; const aCaseSensitive: Boolean = False);
    constructor CreateAdvanced(const aMask: UnicodeString; const aCaseSensitive: Boolean=false; const aOpcodesAllowed: TMaskOpcodesSet=TMaskOpCodesAllAllowed);
    procedure Compile; override;
    function Matches(const aStringToMatch: UnicodeString): Boolean; virtual;
    property Mask: UnicodeString read FOriginalMask write FOriginalMask;
    property OPCodesAllowed;
  end;

  { TMaskAnsiWindows }

  TMaskAnsiWindows=class(TMaskANSI)
  protected
    FMaskWindowsQuirkAllowed: TWindowsQuirkSet;
    cMaskWindowsQuirkInUse: TWindowsQuirkSet;
    FWindowsMask: RawByteString;
    class procedure SplitFileNameExtension(const aSourceFileName: RawByteString; out aFileName: RawByteString; out aExtension: RawByteString; const aIsMask: Boolean=false);static;
  public
    constructor Create(const aMask: RawByteString; const aCaseSensitive: Boolean = False);
    constructor CreateAdvanced(const aMask: RawByteString; const aCaseSensitive: Boolean=false; const aWindowsQuirksAllowed: TWindowsQuirkSet=TWindowsQuirksAllAllowed);
    procedure Compile; override;
    function Matches(const aFileName: RawByteString): Boolean; override;
    property Mask: RawByteString read FWindowsMask write FWindowsMask;
    property Quirks: TWindowsQuirkSet read FMaskWindowsQuirkAllowed write FMaskWindowsQuirkAllowed;
  end;

  { TMaskUTF8Windows }

  TMaskUTF8Windows=class(TMaskUTF8)
  protected
    FMaskWindowsQuirkAllowed: TWindowsQuirkSet;
    cMaskWindowsQuirkInUse: TWindowsQuirkSet;
    FWindowsMask: RawByteString;
    class procedure SplitFileNameExtension(const aSourceFileName: RawByteString; out aFileName: RawByteString; out aExtension: RawByteString; const aIsMask: Boolean=false);static;
  public
    constructor Create(const aMask: RawByteString; const aCaseSensitive: Boolean = False);
    constructor CreateAdvanced(const aMask: RawByteString; const aCaseSensitive: Boolean=false; const aWindowsQuirksAllowed: TWindowsQuirkSet=TWindowsQuirksAllAllowed);
    procedure Compile; override;
    function Matches(const aFileName: RawByteString): Boolean; override;
    property Mask: RawByteString read FWindowsMask write FWindowsMask;
    property Quirks: TWindowsQuirkSet read FMaskWindowsQuirkAllowed write FMaskWindowsQuirkAllowed;
  end;

  { TMaskUnicodeWindows }

  TMaskUnicodeWindows=class(TMaskUnicode)
  protected
    FMaskWindowsQuirkAllowed: TWindowsQuirkSet;
    cMaskWindowsQuirkInUse: TWindowsQuirkSet;
    FWindowsMask: UnicodeString;
    class procedure SplitFileNameExtension(const aSourceFileName: UnicodeString; out aFileName: UnicodeString; out aExtension: UnicodeString; const aIsMask: Boolean=false); static;
  public
    constructor Create(const aMask: UnicodeString; const aCaseSensitive: Boolean = False);
    constructor CreateAdvanced(const aMask: UnicodeString; const aCaseSensitive: Boolean=false; const aWindowsQuirksAllowed: TWindowsQuirkSet=TWindowsQuirksAllAllowed);
    procedure Compile; override;
    function Matches(const aFileName: UnicodeString): Boolean; override;
    property Mask: UnicodeString read FWindowsMask write FWindowsMask;
    property Quirks: TWindowsQuirkSet read FMaskWindowsQuirkAllowed write FMaskWindowsQuirkAllowed;
  end;

  { TMaskDelphiUnicode }

  TMaskDelphiUnicode=class(TMaskUnicode)
  public
    constructor Create(const aMask: UnicodeString);
  end;

  { TMaskDelphiUTF8 }

  TMaskDelphiUTF8=class(TMaskUTF8)
  public
    constructor Create(const aMask: RawByteString);
  end;

implementation

resourcestring
  rsInvalidCharMaskAt = 'Invalid char mask "%s" at %d';
  rsInvalidCharMask = 'Invalid char mask "%s"';
  rsMissingCloseCharMaskAt = 'Missing close char mask "%s" at %d';
  rsMissingCloseCharMask = 'Missing close char mask "%s"';
  rsIncompleteMask = 'Reached end of mask, but missing close/escape '
    +'sequence.';
  rsInvalidEscapeChar = 'Escape character must be ASCII <= 127';
  rsInternalError = 'Internal %s error.';
  {$IFNDEF USE_BRANCHLESS_UTF8LEN}
  rsUTF8WrongEncoding = 'UTF8 wrong encoding detected. %.2x (%.2x) %.2x';
  {$ENDIF}

{ TMaskDelphiUnicode }

constructor TMaskDelphiUnicode.Create(const aMask: UnicodeString);
begin
  inherited CreateAdvanced(aMask,false,TMaskOpCodesAllAllowed-[eMaskOpcodeAnyCharOrNone,eMaskOpcodeEscapeChar]);
  Compile;
end;

{ TMaskDelphiUTF8 }

constructor TMaskDelphiUTF8.Create(const aMask: RawByteString);
begin
  inherited CreateAdvanced(aMask,false,TMaskOpCodesAllAllowed-[eMaskOpcodeAnyCharOrNone,eMaskOpcodeEscapeChar]);
  Compile;
end;

{ EMaskException }

constructor EMaskException.Create(const msg: string;
  const aCode: TMaskExceptionCode);
begin
  CreateFmt(msg,[],aCode);
end;

constructor EMaskException.CreateFmt(const msg: string;
  const args: array of const; const aCode: TMaskExceptionCode);
begin
  FCode:=aCode;
  Inherited CreateFmt(msg,args);
end;

{ TMaskUnicodeWindows }

class procedure TMaskUnicodeWindows.SplitFileNameExtension(
  const aSourceFileName: UnicodeString; out aFileName: UnicodeString; out
  aExtension: UnicodeString; const aIsMask: Boolean);
var
  j: Integer;
  lLowLimit: integer;
begin
  // Default values
  aFileName:=aSourceFileName;
  aExtension:='';

  // This is because .foo is considered a file name ".foo" as one.
  if aIsMask then begin
    lLowLimit:=0;
  end else begin
    lLowLimit:=1;
  end;

  j:=Length(aSourceFileName);
  while j>lLowLimit do begin
    if aSourceFileName[j]='.' then begin
      aFileName:=copy(aSourceFileName,1,j-1);
      aExtension:=copy(aSourceFileName,j);
      break;
    end;
    dec(j);
  end;
end;

constructor TMaskUnicodeWindows.Create(const aMask: UnicodeString;
  const aCaseSensitive: Boolean);
begin
  CreateAdvanced(aMask,aCaseSensitive,TWindowsQuirksDefaultAllowed);
  Compile;
end;

constructor TMaskUnicodeWindows.CreateAdvanced(const aMask: UnicodeString;
  const aCaseSensitive: Boolean; const aWindowsQuirksAllowed: TWindowsQuirkSet);
begin
  FMaskWindowsQuirkAllowed:=aWindowsQuirksAllowed;
  FWindowsMask:=aMask;
  inherited CreateAdvanced(aMask,aCaseSensitive,TMaskOpCodesAllAllowed);
end;

procedure TMaskUnicodeWindows.Compile;

  function RightStr(const aStr: UnicodeString; const aCounter: integer): UnicodeString;
  begin
    Result:=copy(aStr,Length(aStr)-aCounter,aCounter);
  end;

  function StringOfUnicodeChar(const aUnicodeString: UnicodeString; const aCounter: integer): UnicodeString;
  var
    lBlockBytes: integer;
    lBlockChars: integer;
    lPosition, j: integer;
  begin
    lBlockChars:=Length(aUnicodeString);
    lBlockBytes:=lBlockChars*sizeof(UnicodeChar);
    SetLength(Result,lBlockChars*aCounter);
    lPosition:=1;
    for j := 0 to Pred(aCounter) do begin
      move(aUnicodeString[1],Result[lPosition],lBlockBytes);
      inc(lPosition,lBlockChars);
    end;
  end;

  function OptionalQMarksAtEnd(aMask: UnicodeString): UnicodeString;
  var
    lCounter: integer;
    k: integer;
  begin
    lCounter:=0;
    for k := Length(aMask) downto 1 do begin
      if aMask[k]='?' then begin
        inc(lCounter);
      end else begin
        break;
      end;
    end;
    if lCounter>0 then begin
      aMask:=copy(aMask,1,Length(aMask)-lCounter)+'['+StringOfUnicodeChar('?',lCounter)+']';
    end;
    Result:=aMask;
  end;

  function EscapeSpecialChars(const aString: UnicodeString): UnicodeString;
  var
    j: integer;
  begin
    Result:=aString;
    for j := Length(Result) downto 1 do begin
      if InUnicodeChars(Result[j],['[',']',FMaskEscapeChar]) then begin
        // Escape the []\ chars as in Windows mask mode they are plain chars.
        insert(WideChar(FMaskEscapeChar),Result,j);
      end;
    end;
  end;

var
  lFileNameMask: UnicodeString;
  lExtensionMask: UnicodeString;
  lModifiedMask: UnicodeString;

begin
  lModifiedMask:=FWindowsMask;

  // Quirk "blah.*" = "blah*"
  if eWindowsQuirk_AnyExtension in FMaskWindowsQuirkAllowed then begin
    if RightStr(lModifiedMask,3)='*.*' then begin
      lModifiedMask:=copy(lModifiedMask,1,Length(lModifiedMask)-2);
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_AnyExtension];
    end;
  end;

  SplitFileNameExtension(lModifiedMask,lFileNameMask,lExtensionMask,true);

  // Quirk "blah.abc" = "blah.abc*"
  if eWindowsQuirk_Extension3More in FMaskWindowsQuirkAllowed then begin
    if (Length(lExtensionMask)=4) and (Length(lFileNameMask)>0) then begin
      lExtensionMask:=lExtensionMask+'*';
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_Extension3More];
    end;
  end;

  // Quirk "" = "*"
  if (Length(lFileNameMask)=0) and (Length(lExtensionMask)=0) then begin
    if eWindowsQuirk_EmptyIsAny in FMaskWindowsQuirkAllowed then begin
      lFileNameMask:='*';
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_EmptyIsAny];
    end;
  end else begin
  // Quirk ".abc"
    if eWindowsQuirk_AllByExtension in FMaskWindowsQuirkAllowed then begin
      if (Length(lFileNameMask)=0) and (length(lExtensionMask)>0) then begin
        if lExtensionMask[1]='.' then begin
          lFileNameMask:='*';
          cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_AllByExtension];
        end;
      end;
    end;
  end;

  lFileNameMask:=EscapeSpecialChars(lFileNameMask);
  lExtensionMask:=EscapeSpecialChars(lExtensionMask);

  // Quirk "file???.ab?" matches "file1.ab1" and "file123.ab"
  if eWindowsQuirk_FilenameEnd in FMaskWindowsQuirkAllowed then begin
    lFileNameMask:=OptionalQMarksAtEnd(lFileNameMask);
    lExtensionMask:=OptionalQMarksAtEnd(lExtensionMask);
  end;

  if eWindowsQuirk_NoExtension in FMaskWindowsQuirkAllowed then begin
    if Length(lExtensionMask)=1 then begin
      cMaskWindowsQuirkInUse:=[eWindowsQuirk_NoExtension];
      lExtensionMask:='';
    end;
  end;

  inherited Mask:=lFileNameMask+lExtensionMask;
  inherited Compile;
end;

function TMaskUnicodeWindows.Matches(const aFileName: UnicodeString): Boolean;
var
  lFileName, lExtension: UnicodeString;
begin
  if eWindowsQuirk_NoExtension in cMaskWindowsQuirkInUse then begin
    SplitFileNameExtension(aFileName,lFileName,lExtension,false);
    // eWindowsQuirk_NoExtension = Empty extension
    if lExtension<>'' then exit(false);
  end;
  Result:=inherited Matches(aFileName);
end;

{ TMaskUTF8Windows }

class procedure TMaskUTF8Windows.SplitFileNameExtension(
  const aSourceFileName: RawByteString; out aFileName: RawByteString; out
  aExtension: RawByteString; const aIsMask: Boolean);
var
  j: Integer;
  lLowLimit: integer;
begin
  // Default values
  aFileName:=aSourceFileName;
  aExtension:='';

  // This is because .foo is considered a file name ".foo" as one.
  if aIsMask then begin
    lLowLimit:=0;
  end else begin
    lLowLimit:=1;
  end;

  j:=Length(aSourceFileName);
  while j>lLowLimit do begin
    if aSourceFileName[j]='.' then begin
      aFileName:=copy(aSourceFileName,1,j-1);
      aExtension:=copy(aSourceFileName,j);
      break;
    end;
    dec(j);
  end;
end;

constructor TMaskUTF8Windows.Create(const aMask: RawByteString;
  const aCaseSensitive: Boolean);
begin
  CreateAdvanced(aMask,aCaseSensitive,TWindowsQuirksDefaultAllowed);
  Compile;
end;

constructor TMaskUTF8Windows.CreateAdvanced(const aMask: RawByteString;
  const aCaseSensitive: Boolean; const aWindowsQuirksAllowed: TWindowsQuirkSet);
begin
  FMaskWindowsQuirkAllowed:=aWindowsQuirksAllowed;
  FWindowsMask:=aMask;
  inherited CreateAdvanced(aMask,aCaseSensitive,TMaskOpCodesAllAllowed);
end;

procedure TMaskUTF8Windows.Compile;

  function OptionalQMarksAtEnd(aMask: RawByteString): RawByteString;
  var
    lCounter: integer;
    k: integer;
  begin
    lCounter:=0;
    for k := Length(aMask) downto 1 do begin
      if aMask[k]='?' then begin
        inc(lCounter);
      end else begin
        break;
      end;
    end;
    if lCounter>0 then begin
      aMask:=copy(aMask,1,Length(aMask)-lCounter)+'['+StringOfChar('?',lCounter)+']';
    end;
    Result:=aMask;
  end;

  function EscapeSpecialChars(const aString: RawByteString): RawByteString;
  var
    j: integer;
  begin
    Result:=aString;
    for j := Length(Result) downto 1 do begin
      if Result[j] in ['[',']',FMaskEscapeChar] then begin
        // Escape the []\ chars as in Windows mask mode they are plain chars.
        insert(FMaskEscapeChar,Result,j);
      end;
    end;
  end;

var
  lFileNameMask: RawByteString;
  lExtensionMask: RawByteString;
  lModifiedMask: RawByteString;

begin
  lModifiedMask:=FWindowsMask;

  // Quirk "blah.*" = "blah*"
  if eWindowsQuirk_AnyExtension in FMaskWindowsQuirkAllowed then begin
    if RightStr(lModifiedMask,3)='*.*' then begin
      lModifiedMask:=copy(lModifiedMask,1,Length(lModifiedMask)-2);
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_AnyExtension];
    end;
  end;

  SplitFileNameExtension(lModifiedMask,lFileNameMask,lExtensionMask,true);

  // Quirk "blah.abc" = "blah.abc*"
  if eWindowsQuirk_Extension3More in FMaskWindowsQuirkAllowed then begin
    if (Length(lExtensionMask)=4) and (Length(lFileNameMask)>0) then begin
      lExtensionMask:=lExtensionMask+'*';
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_Extension3More];
    end;
  end;

  // Quirk "" = "*"
  if (Length(lFileNameMask)=0) and (Length(lExtensionMask)=0) then begin
    if eWindowsQuirk_EmptyIsAny in FMaskWindowsQuirkAllowed then begin
      lFileNameMask:='*';
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_EmptyIsAny];
    end;
  end else begin
  // Quirk ".abc"
    if eWindowsQuirk_AllByExtension in FMaskWindowsQuirkAllowed then begin
      if (Length(lFileNameMask)=0) and (length(lExtensionMask)>0) then begin
        if lExtensionMask[1]='.' then begin
          lFileNameMask:='*';
          cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_AllByExtension];
        end;
      end;
    end;
  end;

  lFileNameMask:=EscapeSpecialChars(lFileNameMask);
  lExtensionMask:=EscapeSpecialChars(lExtensionMask);

  // Quirk "file???.ab?" matches "file1.ab1" and "file123.ab"
  if eWindowsQuirk_FilenameEnd in FMaskWindowsQuirkAllowed then begin
    lFileNameMask:=OptionalQMarksAtEnd(lFileNameMask);
    lExtensionMask:=OptionalQMarksAtEnd(lExtensionMask);
  end;

  if eWindowsQuirk_NoExtension in FMaskWindowsQuirkAllowed then begin
    if Length(lExtensionMask)=1 then begin
      cMaskWindowsQuirkInUse:=[eWindowsQuirk_NoExtension];
      lExtensionMask:='';
    end;
  end;

  inherited Mask:=lFileNameMask+lExtensionMask;
  inherited Compile;
end;

function TMaskUTF8Windows.Matches(const aFileName: RawByteString): Boolean;
var
  lFileName, lExtension: RawByteString;

begin
  if eWindowsQuirk_NoExtension in cMaskWindowsQuirkInUse then begin
    SplitFileNameExtension(aFileName,lFileName,lExtension,false);
    // eWindowsQuirk_NoExtension = Empty extension
    if lExtension<>'' then exit(false);
  end;
  Result:=Inherited Matches(aFileName);
end;

{ TMaskAnsiWindows }

class procedure TMaskAnsiWindows.SplitFileNameExtension(
  const aSourceFileName: RawByteString; out aFileName: RawByteString; out
  aExtension: RawByteString; const aIsMask: Boolean);
var
  j: Integer;
  lLowLimit: integer;
begin
  // Default values
  aFileName:=aSourceFileName;
  aExtension:='';

  // This is because .foo is considered a file name ".foo" as one.
  if aIsMask then begin
    lLowLimit:=0;
  end else begin
    lLowLimit:=1;
  end;

  j:=Length(aSourceFileName);
  while j>lLowLimit do begin
    if aSourceFileName[j]='.' then begin
      aFileName:=copy(aSourceFileName,1,j-1);
      aExtension:=copy(aSourceFileName,j);
      break;
    end;
    dec(j);
  end;
end;

constructor TMaskAnsiWindows.Create(const aMask: RawByteString;
  const aCaseSensitive: Boolean);
begin
  CreateAdvanced(aMask,aCaseSensitive,TWindowsQuirksDefaultAllowed);
  Compile;
end;

constructor TMaskAnsiWindows.CreateAdvanced(const aMask: RawByteString;
  const aCaseSensitive: Boolean; const aWindowsQuirksAllowed: TWindowsQuirkSet);
begin
  FMaskWindowsQuirkAllowed:=aWindowsQuirksAllowed;
  FWindowsMask:=aMask;
  inherited CreateAdvanced(aMask,aCaseSensitive,TMaskOpCodesAllAllowed);
end;

procedure TMaskAnsiWindows.Compile;

  function OptionalQMarksAtEnd(aMask: RawByteString): RawByteString;
  var
    lCounter: integer;
    k: integer;
  begin
    lCounter:=0;
    for k := Length(aMask) downto 1 do begin
      if aMask[k]='?' then begin
        inc(lCounter);
      end else begin
        break;
      end;
    end;
    if lCounter>0 then begin
      aMask:=copy(aMask,1,Length(aMask)-lCounter)+'['+StringOfChar('?',lCounter)+']';
    end;
    Result:=aMask;
  end;

  function EscapeSpecialChars(const aString: RawByteString): RawByteString;
  var
    j: integer;
  begin
    Result:=aString;
    for j := Length(Result) downto 1 do begin
      if Result[j] in ['[',']',FMaskEscapeChar] then begin
        // Escape the []\ chars as in Windows mask mode they are plain chars.
        insert(FMaskEscapeChar,Result,j);
      end;
    end;
  end;

var
  lFileNameMask: RawByteString;
  lExtensionMask: RawByteString;
  lModifiedMask: RawByteString;

begin
  lModifiedMask:=FWindowsMask;

  // Quirk "blah.*" = "blah*"
  if eWindowsQuirk_AnyExtension in FMaskWindowsQuirkAllowed then begin
    if RightStr(lModifiedMask,3)='*.*' then begin
      lModifiedMask:=copy(lModifiedMask,1,Length(lModifiedMask)-2);
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_AnyExtension];
    end;
  end;

  SplitFileNameExtension(lModifiedMask,lFileNameMask,lExtensionMask,true);

  // Quirk "blah.abc" = "blah.abc*"
  if eWindowsQuirk_Extension3More in FMaskWindowsQuirkAllowed then begin
    if (Length(lExtensionMask)=4) and (Length(lFileNameMask)>0) then begin
      lExtensionMask:=lExtensionMask+'*';
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_Extension3More];
    end;
  end;

  // Quirk "" = "*"
  if (Length(lFileNameMask)=0) and (Length(lExtensionMask)=0) then begin
    if eWindowsQuirk_EmptyIsAny in FMaskWindowsQuirkAllowed then begin
      lFileNameMask:='*';
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_EmptyIsAny];
    end;
  end else begin
  // Quirk ".abc"
    if eWindowsQuirk_AllByExtension in FMaskWindowsQuirkAllowed then begin
      if (Length(lFileNameMask)=0) and (length(lExtensionMask)>0) then begin
        if lExtensionMask[1]='.' then begin
          lFileNameMask:='*';
          cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_AllByExtension];
        end;
      end;
    end;
  end;

  lFileNameMask:=EscapeSpecialChars(lFileNameMask);
  lExtensionMask:=EscapeSpecialChars(lExtensionMask);

  // Quirk "file???.ab?" matches "file1.ab1" and "file123.ab"
  if eWindowsQuirk_FilenameEnd in FMaskWindowsQuirkAllowed then begin
    lFileNameMask:=OptionalQMarksAtEnd(lFileNameMask);
    lExtensionMask:=OptionalQMarksAtEnd(lExtensionMask);
  end;

  if eWindowsQuirk_NoExtension in FMaskWindowsQuirkAllowed then begin
    if Length(lExtensionMask)=1 then begin
      cMaskWindowsQuirkInUse:=cMaskWindowsQuirkInUse+[eWindowsQuirk_NoExtension];
      lExtensionMask:='';
    end;
  end;

  inherited Mask:=lFileNameMask+lExtensionMask;
  inherited Compile;
end;

function TMaskAnsiWindows.Matches(const aFileName: RawByteString): Boolean;
var
  lFileName: RawByteString;
  lExtension: RawByteString;
begin
  if eWindowsQuirk_NoExtension in cMaskWindowsQuirkInUse then begin
    SplitFileNameExtension(aFileName,lFileName,lExtension,false);
    // eWindowsQuirk_NoExtension = Empty extension
    if lExtension<>'' then exit(false);
  end;
  Result:=inherited Matches(aFileName);
end;

{ TMaskBase }

procedure TMaskBase.SetMaskEscapeChar(AValue: Char);
begin
  if FMaskEscapeChar=AValue then Exit;
  if FMaskEscapeChar>#127 then begin
    Exception_InvalidEscapeChar();
  end;
  FMaskEscapeChar:=AValue;
end;

procedure TMaskBase.Add(const aLength: integer; const aData: PBYTE);
var
  lCounter: integer;
begin
  if cMaskCompiledIndex+aLength>=cMaskCompiledAllocated then begin
    cMaskCompiledAllocated:=cMaskCompiledAllocated+aLength+GROW_BY;
    SetLength(cMaskCompiled,cMaskCompiledAllocated);
  end;
  for lCounter := 0 to Pred(aLength) do begin
    cMaskCompiled[cMaskCompiledIndex]:=(aData+lCounter)^;
    inc(cMaskCompiledIndex);
  end;
end;

procedure TMaskBase.Add(const aValue: integer);
begin
  Add(sizeof(aValue),@aValue);
end;

procedure TMaskBase.Add(const aValue: TMaskOpCode);
var
  v: BYTE;
begin
  v:=BYTE(aValue);
  Add(1,@v);
end;

procedure TMaskBase.IncrementLastCounterBy(const aOpcode: TMaskOpCode;
  const aValue: integer);
var
  p: PINTEGER;
begin
  cMaskCompiledIndex:=cMaskCompiledIndex-sizeof(aValue);
  if TMaskOpCode(cMaskCompiled[cMaskCompiledIndex-1])<>aOpcode then begin
    Exception_InternalError();
  end;
  P:=@cMaskCompiled[cMaskCompiledIndex];
  Add(P^+aValue);
end;

procedure TMaskBase.Compile;
begin
  cMaskIsCompiled:=true;
end;

class procedure TMaskBase.Exception_InvalidCharMask(const aMaskChar: string;
  const aOffset: integer);
begin
  if aOffset>=0 then begin
    raise EMaskException.CreateFmt(rsInvalidCharMaskAt, [aMaskChar, aOffset], eMaskException_InvalidCharMask);
  end else begin
    raise EMaskException.CreateFmt(rsInvalidCharMask, [aMaskChar], eMaskException_InvalidCharMask);
  end;
end;

class procedure TMaskBase.Exception_MissingCloseChar(const aMaskChar: string;
  const aOffset: integer);
begin
  if aOffset>=0 then begin
    raise EMaskException.CreateFmt(rsMissingCloseCharMaskAt, [aMaskChar, aOffset], eMaskException_MissingClose);
  end else begin
    raise EMaskException.CreateFmt(rsMissingCloseCharMask, [aMaskChar], eMaskException_MissingClose);
  end;
end;

class procedure TMaskBase.Exception_IncompleteMask();
begin
  raise EMaskException.CreateFmt(rsIncompleteMask, [], eMaskException_IncompleteMask);
end;

class procedure TMaskBase.Exception_InvalidEscapeChar();
begin
  raise EMaskException.Create(rsInvalidEscapeChar, eMaskException_InvalidEscapeChar);
end;

procedure TMaskBase.Exception_InternalError();
begin
  raise EMaskException.CreateFmt(rsInternalError, [self.ClassName], eMaskException_InternalError);
end;

constructor TMaskBase.CreateAdvanced(const aCaseSensitive: Boolean;
  const aOpcodesAllowed: TMaskOpcodesSet);
begin
  FMaskOpcodesAllowed:=aOpcodesAllowed;
  FCaseSensitive:=aCaseSensitive;
  FMaskEscapeChar:='\';
end;

constructor TMaskBase.Create(const aCaseSensitive: Boolean);
begin
  CreateAdvanced(aCaseSensitive,TMaskOpCodesDefaultAllowed);
end;

{ TMaskANSI }

function TMaskANSI.AnsiOEMToUTF8(const aAnsiChar: char): RawByteString;
type
  LocalCPString=type AnsiString(CP_OEMCP);
var
  lString: LocalCPString;
begin
  lString:=aAnsiChar;
  Result:=lString;
end;

function TMaskANSI.UTF8ToAnsiOEM(const aUTF8String: RawByteString
  ): RawByteString;
type
  LocalCPString=type AnsiString(CP_OEMCP);
var
  lString: LocalCPString;
begin
  lString:=UTF8ToAnsi(aUTF8String);
  Result:=lString;
end;

procedure TMaskANSI.Compile;
var
  j: Integer;
  lCharsGroupInsertSize: integer;
  lLast: TMaskOpCode;
  lMask: RawByteString;
  lFirstRange, lSecondRange: integer;

  function IsARange(aPosition: integer; out aFirstSequence: integer; out aSecondSequence: integer): Boolean; {$IFDEF USE_INLINE}inline;{$ENDIF}
  begin
    Result:=false;
    aFirstSequence:=0;
    aSecondSequence:=0;
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[aPosition]=FMaskEscapeChar) then begin
      if aPosition+1>cMaskLimit then exit; // ==>
      inc(aPosition);
    end;
    aFirstSequence:=aPosition;
    inc(aPosition);
    if aPosition+1>cMaskLimit then exit; // ==>
    // It is not a range, return false
    if lMask[aPosition]<>'-' then exit;

    inc(aPosition,1);
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[aPosition]=FMaskEscapeChar) then begin
      if aPosition+1>cMaskLimit then exit; // ==>
      inc(aPosition);
    end;
    aSecondSequence:=aPosition;
    Result:=true;
  end;

begin
  inherited Compile;
  if not FCaseSensitive then begin
    lMask:=LowerCase(FOriginalMask);
  end else begin
    lMask:=FOriginalMask;
  end;
  cMaskLimit:=Length(lMask);
  lLast:=TMaskOpCode.Literal;
  SetLength(cMaskCompiled,0);
  j:=1;
  while j<=cMaskLimit do begin
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[j]=FMaskEscapeChar) then begin
      // next is Literal
      inc(j);
      if j<=cMaskLimit then begin
        Add(TMaskOpCode.Literal);
        Add(1,@lMask[j]);
        inc(cMatchMinimumLiteralBytes);
        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,1);
        lLast:=TMaskOpCode.Literal;
        inc(j);
      end else begin
        Exception_IncompleteMask();
      end;
    end else begin
      if lMask[j] in ['*','?','['] then begin
        case lMask[j] of
          '*':
            begin
              if eMaskOpcodeAnyText in FMaskOpcodesAllowed then begin
                if lLast<>TMaskOpCode.AnyCharToNext then begin
                  Add(TMaskOpCode.AnyCharToNext);
                  lLast:=TMaskOpCode.AnyCharToNext;
                  // * = No limit
                  cMatchMaximumLiteralBytes:=High(cMatchMaximumLiteralBytes);
                end;
              end else begin
                Add(TMaskOpCode.Literal);
                Add(1,@lMask[j]);
                inc(cMatchMinimumLiteralBytes);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
          '?':
            begin
              if eMaskOpcodeAnyChar in FMaskOpcodesAllowed then begin
                Add(TMaskOpCode.AnyChar);
                lLast:=TMaskOpCode.AnyChar;
                inc(cMatchMinimumLiteralBytes,1);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,1);
              end else begin
                Add(TMaskOpCode.Literal);
                Add(1,@lMask[j]);
                inc(cMatchMinimumLiteralBytes);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,1);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
          '[':
            begin
              if (eMaskOpcodeOptionalChar in FMaskOpcodesAllowed) or
                 (eMaskOpcodeRange in FMaskOpcodesAllowed) or
                 (eMaskOpcodeAnyCharOrNone in FMaskOpcodesAllowed)
                 then begin
                lLast:=TMaskOpCode.CharsGroupBegin;
                Add(TMaskOpCode.CharsGroupBegin);
                inc(cMatchMinimumLiteralBytes);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes);
                lCharsGroupInsertSize:=cMaskCompiledIndex;
                Add(0);
                inc(j); // CP length is 1 because it is "["
                if j<cMaskLimit then begin
                  if (lMask[j]='!') and (eMaskOpcodeNegateGroup in FMaskOpcodesAllowed) then begin
                    Add(TMaskOpCode.Negate);
                    inc(j); // CP length is 1 because it is "!"
                    lLast:=TMaskOpCode.Negate;
                  end;
                end;

                while j<=cMaskLimit do begin
                  if (lMask[j]='?') and (eMaskOpcodeAnyCharOrNone in FMaskOpcodesAllowed) then begin
                    // This syntax is permitted [??] but not this one [?a] or [a?]
                    if (lLast=TMaskOpCode.CharsGroupBegin) or (lLast=TMaskOpCode.AnyCharOrNone) then begin
                      if lLast=TMaskOpCode.AnyCharOrNone then begin
                        // Increment counter
                        IncrementLastCounterBy(TMaskOpCode.AnyCharOrNone,1);
                      end else begin
                        Add(TMaskOpCode.AnyCharOrNone);
                        Add(1); // Counter
                        // Discount minimun bytes added at the "CharGroupBegin"
                        // because [?] could be 1 or 0 chars, so minimum is zero
                        // but the CharsGroupBegin assumes 1 char as all other
                        // masks replace the group by 1 char position.
                        // This code will run 1 time per group at maximun.
                        dec(cMatchMinimumLiteralBytes);
                        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then dec(cMatchMaximumLiteralBytes);
                      end;
                      if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes);
                      lLast:=TMaskOpCode.AnyCharOrNone;
                    end else begin
                      Exception_InvalidCharMask(lMask[j],j);
                    end;

                  end else if (lLast=TMaskOpCode.AnyCharOrNone) and (lMask[j]<>']') then begin
                    //lMask[j] is not '?', but previous mask was '?' and it is an invalid sequence.
                    // "[??] = Valid" // "[a?] or [?a] = Invalid"
                    Exception_InvalidCharMask(lMask[j],j);

                  end else if (eMaskOpcodeRange in FMaskOpcodesAllowed) and IsARange(j,lFirstRange,lSecondRange) then begin
                    Add(TMaskOpCode.Range);
                    // Check if reverse range is needed
                    if lMask[lFirstRange]>#127 then begin
                      Exception_InvalidCharMask(lMask[lFirstRange],j);
                    end;
                    if lMask[lSecondRange]>#127 then begin
                      Exception_InvalidCharMask(lMask[lSecondRange],j);
                    end;
                    {$IFDEF RANGES_AUTOREVERSE}
                    if lMask[lFirstRange]<lMask[lSecondRange] then begin
                      Add(1,@lMask[lFirstRange]);
                      Add(1,@lMask[lSecondRange]);
                      j:=lSecondRange;
                    end else begin
                      Add(1,@lMask[lSecondRange]);
                      Add(1,@lMask[lFirstRange]);
                      j:=lSecondRange;
                    end;
                    {$ELSE}
                    Add(1,@lMask[lFirstRange]);
                    Add(1,@lMask[lSecondRange]);
                    j:=lSecondRange;
                    {$ENDIF}
                    lLast:=TMaskOpCode.Range;

                  end else if lMask[j]=']' then begin
                    if (lLast=TMaskOpCode.CharsGroupBegin) or (lLast=TMaskOpCode.Negate) then begin
                      //Error empty match
                      Exception_InvalidCharMask(lMask[j],j);
                    end;
                    // Insert the new offset in case of a positive match in CharsGroup
                    PInteger(@cMaskCompiled[lCharsGroupInsertSize])^:=cMaskCompiledIndex;
                    Add(TMaskOpCode.CharsGroupEnd);
                    lLast:=TMaskOpCode.CharsGroupEnd;
                    break;
                  end else begin
                    if (lMask[j]=FMaskEscapeChar) and (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) then begin
                      inc(j);
                      if j>cMaskLimit then begin
                        Exception_IncompleteMask();
                      end;
                    end;
                    if eMaskOpcodeOptionalChar in FMaskOpcodesAllowed then begin
                      Add(TMaskOpCode.OptionalChar);
                      Add(1,@lMask[j]);
                      lLast:=TMaskOpCode.OptionalChar;
                    end else begin
                      Exception_InvalidCharMask(lMask[j],j);
                    end;
                  end;
                  inc(j);
                end;
                if j>cMaskLimit then begin
                  Exception_MissingCloseChar(']',cMaskLimit);
                end;
              end else begin
                if (lMask[j]=FMaskEscapeChar) and (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) then begin
                  // next is Literal
                  inc(j);
                  if j>cMaskLimit then begin
                    Exception_IncompleteMask();
                  end;
                end;
                Add(TMaskOpCode.Literal);
                Add(1,@lMask[j]);
                inc(cMatchMinimumLiteralBytes);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,1);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
        end;
      end else begin
        // Literal
        Add(TMaskOpCode.Literal);
        Add(1,@lMask[j]);
        inc(cMatchMinimumLiteralBytes);
        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,1);
        lLast:=TMaskOpCode.Literal;
      end;
      inc(j);
    end;
  end;
  SetLength(cMaskCompiled,cMaskCompiledIndex);
  cMaskCompiledLimit:=cMaskCompiledIndex-1;
end;

function TMaskANSI.intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause;
var
  c,c2: char;
  lFailCause: TMaskFailCause;
  lNegateCharGroup: Boolean;
  lSkipOnSuccessGroup: integer;
  t1: Boolean;
  lTryCounter: integer;
  j: integer;
begin
  lSkipOnSuccessGroup:=0;
  Result:=UnexpectedEnd;
  lNegateCharGroup:=false;
  while aMaskIndex<=cMaskCompiledLimit do begin
    case TMaskOpCode(cMaskCompiled[aMaskIndex]) of
      TMaskOpCode.Literal:
        begin
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          c:=PChar(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex);
          if cMatchString[aMatchOffset]<>c then begin
            Result:=TMaskFailCause.MaskNotMatch;
            Exit;
          end;
          inc(aMatchOffset);
        end;
      TMaskOpCode.AnyChar:
        begin
          inc(aMaskIndex);
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMatchOffset);
        end;
      TMaskOpCode.Negate:
        begin
          lNegateCharGroup:=true;
          inc(aMaskIndex);
        end;
      TMaskOpCode.CharsGroupBegin:
        begin
          lNegateCharGroup:=false;
          inc(aMaskIndex);
          lSkipOnSuccessGroup:=PInteger(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex,sizeof(integer));
        end;
      TMaskOpCode.CharsGroupEnd:
        begin
          if lNegateCharGroup then begin
            aMaskIndex:=lSkipOnSuccessGroup+1;
            inc(aMatchOffset);
          end else begin
            Result:=TMaskFailCause.MaskNotMatch;
            exit;
          end;
        end;
      TMaskOpCode.OptionalChar:
        begin
          inc(aMaskIndex);
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          c:=PChar(@cMaskCompiled[aMaskIndex])^;
          if cMatchString[aMatchOffset]=c then begin
            if lNegateCharGroup then begin
              Result:=TMaskFailCause.MaskNotMatch;
              exit;
            end;
            aMaskIndex:=lSkipOnSuccessGroup+1;
            inc(aMatchOffset);
          end else begin
            inc(aMaskIndex);
          end;
        end;
      TMaskOpCode.Range:
        begin
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          c:=PChar(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex);
          c2:=PChar(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex);
          t1:=(cMatchString[aMatchOffset]>=c) and (cMatchString[aMatchOffset]<=c2);
          if t1 then begin
            if not lNegateCharGroup then begin
              //Jump to CharsGroupEnd+1 because if CharsGroupEnd is reached
              //it means that all optional chars and ranges have not matched the string.
              aMaskIndex:=lSkipOnSuccessGroup+1;
              inc(aMatchOffset);
            end else begin
              Result:=TMaskFailCause.MaskNotMatch;
              exit;
            end;
          end
        end;
      TMaskOpCode.AnyCharToNext:
        begin
          // if last is *, everything in remain match matches
          if aMaskIndex=cMaskCompiledLimit then begin
            Result:=TMaskFailCause.Success;
            exit;
          end;
          if aMatchOffset>cMatchStringLimit then begin
            if aMaskIndex=cMaskCompiledLimit then begin
              Result:=TMaskFailCause.Success;
              exit;
            end;
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          while aMatchOffset<=cMatchStringLimit do begin
            lFailCause:=intfMatches(aMatchOffset,aMaskIndex);
            if lFailCause=TMaskFailCause.Success then begin
              Result:=TMaskFailCause.Success;
              exit;
            end else if lFailCause=TMaskFailCause.MatchStringExhausted then begin
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            inc(aMatchOffset);
          end;
          Result:=TMaskFailCause.MatchStringExhausted;
          exit;
        end;
      TMaskOpCode.AnyCharOrNone:
        begin
          inc(aMaskIndex);
          lTryCounter:=PInteger(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex,sizeof(integer));
          if TMaskOpCode(cMaskCompiled[aMaskIndex])<>TMaskOpCode.CharsGroupEnd then begin
            Exception_InternalError();
          end else begin
            aMaskIndex:=lSkipOnSuccessGroup+1;
          end;

          // Try to match remain mask eating, 0,1,2,...,lTryCounter chars.
          for j := 0 to lTryCounter do begin
            if aMatchOffset>cMatchStringLimit then begin
              if aMaskIndex=cMaskCompiledLimit+1 then begin
                Result:=TMaskFailCause.Success;
                exit;
              end;
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            lFailCause:=intfMatches(aMatchOffset,aMaskIndex);
            if lFailCause=TMaskFailCause.Success then begin
              Result:=TMaskFailCause.Success;
              exit;
            end else if lFailCause=TMaskFailCause.MatchStringExhausted then begin
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            inc(aMatchOffset);
          end;
          Result:=TMaskFailCause.MatchStringExhausted;
          exit;
        end;
      else
        begin
          Exception_InternalError();
        end;
    end;
  end;
  if (aMaskIndex>cMaskCompiledLimit) and (aMatchOffset>cMatchStringLimit) then begin
    Result:=TMaskFailCause.Success;
  end else begin
    if aMaskIndex>cMaskCompiledLimit then begin
      Result:=TMaskFailCause.MaskExhausted;
    end else begin
      Result:=TMaskFailCause.MatchStringExhausted;
    end;
  end;
end;

constructor TMaskANSI.Create(const aMask: RawByteString;
  const aCaseSensitive: Boolean);
begin
  inherited Create(aCaseSensitive);
  FOriginalMask:=aMask;
end;

constructor TMaskANSI.CreateAdvanced(const aMask: RawByteString;
  const aCaseSensitive: Boolean; const aOpcodesAllowed: TMaskOpcodesSet);
begin
  inherited CreateAdvanced(aCaseSensitive,aOpcodesAllowed);
  FOriginalMask:=aMask;
end;

function TMaskANSI.Matches(const aStringToMatch: RawByteString): Boolean;
begin
  if not cMaskIsCompiled then Compile;
  if not FCaseSensitive then begin
    cMatchString:=AnsiLowerCase(aStringToMatch);
  end else begin
    cMatchString:=aStringToMatch;
  end;
  cMatchStringLimit:=length(cMatchString);
  if (cMatchStringLimit>=cMatchMinimumLiteralBytes) and (cMatchStringLimit<=cMatchMaximumLiteralBytes) then begin
    if intfMatches(1,0)=TMaskFailCause.Success then begin
      Result:=true;
    end else begin
      Result:=false;
    end;
  end else begin
    // There are too many or not enough bytes to match the string
    Result:=false;
  end;
end;

{ TMaskUTF8 }

procedure TMaskUTF8.Compile;
var
  j: Integer;
  lCharsGroupInsertSize: integer;
  lCPLength: integer;
  lLast: TMaskOpCode;
  lMask: RawByteString;
  lFirstRange, lSecondRange: integer;

  function IsARange(aPosition: integer; out aFirstSequence: integer; out aSecondSequence: integer): Boolean; {$IFDEF USE_INLINE}inline;{$ENDIF}
  var
    llCPL: integer;
  begin
    Result:=false;
    aFirstSequence:=0;
    aSecondSequence:=0;
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[aPosition]=FMaskEscapeChar) then begin
      llCPL:=UTF8Length(@lMask[aPosition]);
      if aPosition+llCPL>cMaskLimit then exit; // ==>
      inc(aPosition,llCPL);
      aFirstSequence:=aPosition;
    end else begin
      aFirstSequence:=aPosition;
    end;
    llCPL:=UTF8Length(@lMask[aPosition]);
    inc(aPosition,llCPL);
    if aPosition+llCPL>cMaskLimit then exit; // ==>
    // It is not a range, return false
    if lMask[aPosition]<>'-' then exit;

    llCPL:=UTF8Length(@lMask[aPosition]);
    inc(aPosition,llCPL);
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[aPosition]=FMaskEscapeChar) then begin
      llCPL:=UTF8Length(@lMask[aPosition]);
      if aPosition+llCPL>cMaskLimit then exit; // ==>
      inc(aPosition,llCPL);
      aSecondSequence:=aPosition;
    end else begin
      aSecondSequence:=aPosition;
    end;
    Result:=true;
  end;

begin
  inherited Compile;
  if not FCaseSensitive then begin
    lMask:=LowerCase(FOriginalMask);
  end else begin
    lMask:=FOriginalMask;
  end;
  cMaskLimit:=Length(lMask);
  lLast:=TMaskOpCode.Literal;
  SetLength(cMaskCompiled,0);
  j:=1;
  while j<=cMaskLimit do begin
    lCPLength:=UTF8Length(@lMask[j]);
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[j]=FMaskEscapeChar) then begin
      // next is Literal
      inc(j,lCPLength);
      if j<=cMaskLimit then begin
        lCPLength:=UTF8Length(@lMask[j]);
        Add(TMaskOpCode.Literal);
        Add(lCPLength,@lMask[j]);
        inc(cMatchMinimumLiteralBytes,lCPLength);
        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength);
        lLast:=TMaskOpCode.Literal;
        inc(j,lCPLength);
      end else begin
        Exception_IncompleteMask();
      end;
    end else begin
      if lMask[j] in ['*','?','['] then begin
        case lMask[j] of
          '*':
            begin
              if eMaskOpcodeAnyText in FMaskOpcodesAllowed then begin
                if lLast<>TMaskOpCode.AnyCharToNext then begin
                  Add(TMaskOpCode.AnyCharToNext);
                  lLast:=TMaskOpCode.AnyCharToNext;
                  // * = No limit
                  cMatchMaximumLiteralBytes:=High(cMatchMaximumLiteralBytes);
                end;
              end else begin
                Add(TMaskOpCode.Literal);
                Add(lCPLength,@lMask[j]);
                inc(cMatchMinimumLiteralBytes,lCPLength);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
          '?':
            begin
              if eMaskOpcodeAnyChar in FMaskOpcodesAllowed then begin
                Add(TMaskOpCode.AnyChar);
                inc(cMatchMinimumLiteralBytes,1);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,4);
                lLast:=TMaskOpCode.AnyChar;
              end else begin
                Add(TMaskOpCode.Literal);
                Add(lCPLength,@lMask[j]);
                inc(cMatchMinimumLiteralBytes,lCPLength);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
          '[':
            begin
              if (eMaskOpcodeOptionalChar in FMaskOpcodesAllowed) or
                 (eMaskOpcodeRange in FMaskOpcodesAllowed) or
                 (eMaskOpcodeAnyCharOrNone in FMaskOpcodesAllowed)
                 then begin
                lLast:=TMaskOpCode.CharsGroupBegin;
                Add(TMaskOpCode.CharsGroupBegin);
                inc(cMatchMinimumLiteralBytes,1);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,4);
                lCharsGroupInsertSize:=cMaskCompiledIndex;
                Add(0);
                inc(j); // CP length is 1 because it is "["
                if j<cMaskLimit then begin
                  if (lMask[j]='!') and (eMaskOpcodeNegateGroup in FMaskOpcodesAllowed) then begin
                    Add(TMaskOpCode.Negate);
                    inc(j); // CP length is 1 because it is "!"
                    lLast:=TMaskOpCode.Negate;
                  end;
                end;

                while j<=cMaskLimit do begin
                  lCPLength:=UTF8Length(@lMask[j]);

                  if (lMask[j]='?') and (eMaskOpcodeAnyCharOrNone in FMaskOpcodesAllowed) then begin
                    // This syntax is permitted [??] but not this one [?a] or [a?]
                    if (lLast=TMaskOpCode.CharsGroupBegin) or (lLast=TMaskOpCode.AnyCharOrNone) then begin
                      if lLast=TMaskOpCode.AnyCharOrNone then begin
                        // Increment counter
                        IncrementLastCounterBy(TMaskOpCode.AnyCharOrNone,1);
                      end else begin
                        Add(TMaskOpCode.AnyCharOrNone);
                        Add(1); // Counter
                        // Discount minimun bytes added at the "CharGroupBegin"
                        // because [?] could be 1 or 0 chars, so minimum is zero
                        // but the CharsGroupBegin assumes 1 char as all other
                        // masks replace the group by 1 char position.
                        // This code will run 1 time per group at maximun.
                        dec(cMatchMinimumLiteralBytes,1);
                        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then dec(cMatchMaximumLiteralBytes,4);
                      end;
                      if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,4);
                      lLast:=TMaskOpCode.AnyCharOrNone;
                    end else begin
                      Exception_InvalidCharMask(lMask[j],j);
                    end;

                  end else if (lLast=TMaskOpCode.AnyCharOrNone) and (lMask[j]<>']') then begin
                    //lMask[j] is not '?', but previous mask was '?' and it is an invalid sequence.
                    // "[??] = Valid" // "[a?] or [?a] = Invalid"
                    Exception_InvalidCharMask(lMask[j],j);

                  end else if (eMaskOpcodeRange in FMaskOpcodesAllowed) and IsARange(j,lFirstRange,lSecondRange) then begin
                    Add(TMaskOpCode.Range);
                    // Check if reverse range is needed
                    {$IFDEF RANGES_AUTOREVERSE}
                    if CompareUTF8Sequences(@lMask[lFirstRange],@lMask[lSecondRange])<0 then begin
                      lCPLength:=UTF8Length(@lMask[lFirstRange]);
                      Add(lCPLength,@lMask[lFirstRange]);
                      lCPLength:=UTF8Length(@lMask[lSecondRange]);
                      Add(lCPLength,@lMask[lSecondRange]);
                      j:=lSecondRange;
                    end else begin
                      lCPLength:=UTF8Length(@lMask[lSecondRange]);
                      Add(lCPLength,@lMask[lSecondRange]);
                      Add(UTF8Length(@lMask[lFirstRange]),@lMask[lFirstRange]);
                      j:=lSecondRange;
                    end;
                    {$ELSE}
                      lCPLength:=UTF8Length(@lMask[lFirstRange]);
                      Add(lCPLength,@lMask[lFirstRange]);
                      lCPLength:=UTF8Length(@lMask[lSecondRange]);
                      Add(lCPLength,@lMask[lSecondRange]);
                      j:=lSecondRange;
                    {$ENDIF}
                    lLast:=TMaskOpCode.Range;

                  end else if lMask[j]=']' then begin
                    if (lLast=TMaskOpCode.CharsGroupBegin) or (lLast=TMaskOpCode.Negate) then begin
                      //Error empty match
                      Exception_InvalidCharMask(lMask[j],j);
                    end;
                    // Insert the new offset in case of a positive match in CharsGroup
                    PInteger(@cMaskCompiled[lCharsGroupInsertSize])^:=cMaskCompiledIndex;
                    Add(TMaskOpCode.CharsGroupEnd);
                    lLast:=TMaskOpCode.CharsGroupEnd;
                    break;
                  end else begin
                    if (lMask[j]=FMaskEscapeChar) and (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) then begin
                      // next is optional char in set or literal
                      inc(j,lCPLength);
                      if j<=cMaskLimit then begin
                        lCPLength:=UTF8Length(@lMask[j]);
                      end else begin
                        Exception_IncompleteMask();
                      end;
                    end;
                    if eMaskOpcodeOptionalChar in FMaskOpcodesAllowed then begin
                      Add(TMaskOpCode.OptionalChar);
                      Add(lCPLength,@lMask[j]);
                      lLast:=TMaskOpCode.OptionalChar;
                    end else begin
                      Exception_InvalidCharMask(lMask[j],j);
                    end;
                  end;
                  inc(j,lCPLength);
                end;
                if j>cMaskLimit then begin
                  Exception_MissingCloseChar(']',cMaskLimit);
                end;
              end else begin
                if (lMask[j]=FMaskEscapeChar) and (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) then begin
                  // next is Literal
                  inc(j,lCPLength);
                  if j<=cMaskLimit then begin
                    lCPLength:=UTF8Length(@lMask[j]);
                  end else begin
                    Exception_IncompleteMask();
                  end;
                end;
                Add(TMaskOpCode.Literal);
                Add(lCPLength,@lMask[j]);
                inc(cMatchMinimumLiteralBytes,lCPLength);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
        end;
      end else begin
        // Literal
        Add(TMaskOpCode.Literal);
        Add(lCPLength,@lMask[j]);
        inc(cMatchMinimumLiteralBytes,lCPLength);
        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength);
        lLast:=TMaskOpCode.Literal;
      end;
      inc(j,lCPLength);
    end;
  end;
  SetLength(cMaskCompiled,cMaskCompiledIndex);
  cMaskCompiledLimit:=cMaskCompiledIndex-1;
end;

class function TMaskUTF8.UTF8Length(const P: PBYTE): integer; static;
{$IFDEF USE_BRANCHLESS_UTF8LEN}
const
  UTF8LengthsArray: array [0..31] of BYTE = (
          1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
          0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 2, 2, 3, 3, 4, 0
      );
begin
  Result:=UTF8LengthsArray[P^ shr 3];
{$ELSE}
var
  c: BYTE;
begin

  c:=P^;
  if (c and %10000000) = %00000000 then begin
    Result:=1;
  end else if (c and %11100000) = %11000000 then begin
    Result:=2;
  end else if (c and %11110000) = %11100000 then begin
    Result:=3;
  end else if (c and %11111000) = %11110000 then begin
    Result:=4;
  end else if (c and %11111100) = %11111000 then begin
    Result:=5;
  end else if (c and %11111110) = %11111100 then begin
    Result:=6;
  end else begin
    raise EMaskException.CreateFmt(rsUTF8WrongEncoding, [(P-1)^, P^, (P+1)^], eMaskException_InvalidUTF8Sequence);
  end;

{$ENDIF}
end;

class function TMaskUTF8.CompareUTF8Sequences(const P1, P2: PBYTE): integer;
var
  l1,l2: integer;
  l: integer;
begin
  l1:=UTF8Length(p1);
  l2:=UTF8Length(p2);
  Result:=0;
  l:=0;
  while (l<l1) and (l<l2) do begin
    Result:=Integer((P1+l)^)-integer((P2+l)^);
    if Result<>0 then exit;
    inc(l);
  end;
  Result:=l1-l2;
end;

function TMaskUTF8.intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause;
var
  c1,c2: PBYTE;
  lFailCause: TMaskFailCause;
  lNegateCharGroup: Boolean;
  lSkipOnSuccessGroup: integer;
  t1: Boolean;
  j: integer;
  lTryCounter: integer;
begin
  lSkipOnSuccessGroup:=0;
  Result:=UnexpectedEnd;
  lNegateCharGroup:=false;
  while aMaskIndex<=cMaskCompiledLimit do begin
    case TMaskOpCode(cMaskCompiled[aMaskIndex]) of
      TMaskOpCode.Literal:
        begin
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          if CompareUTF8Sequences(@cMaskCompiled[aMaskIndex],@cMatchString[aMatchOffset])<>0 then begin
            Result:=TMaskFailCause.MaskNotMatch;
            Exit;
          end;
          inc(aMaskIndex,UTF8Length(@cMaskCompiled[aMaskIndex]));
          inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
        end;
      TMaskOpCode.AnyChar:
        begin
          inc(aMaskIndex);
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
        end;
      TMaskOpCode.Negate:
        begin
          lNegateCharGroup:=true;
          inc(aMaskIndex);
        end;
      TMaskOpCode.CharsGroupBegin:
        begin
          lNegateCharGroup:=false;
          inc(aMaskIndex);
          lSkipOnSuccessGroup:=PInteger(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex,sizeof(integer));
        end;
      TMaskOpCode.CharsGroupEnd:
        begin
          if lNegateCharGroup then begin
            aMaskIndex:=lSkipOnSuccessGroup+1;
            inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
          end else begin
            Result:=TMaskFailCause.MaskNotMatch;
            exit;
          end;
        end;
      TMaskOpCode.OptionalChar:
        begin
          inc(aMaskIndex);
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          if CompareUTF8Sequences(@cMaskCompiled[aMaskIndex],@cMatchString[aMatchOffset])=0 then begin
            if lNegateCharGroup then begin
              Result:=TMaskFailCause.MaskNotMatch;
              exit;
            end;
            aMaskIndex:=lSkipOnSuccessGroup+1;
            inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
          end else begin
            inc(aMaskIndex,UTF8Length(@cMaskCompiled[aMaskIndex]));
          end;
        end;
      TMaskOpCode.Range:
        begin
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          c1:=@cMaskCompiled[aMaskIndex];
          inc(aMaskIndex,UTF8Length(C1));
          c2:=@cMaskCompiled[aMaskIndex];
          inc(aMaskIndex,UTF8Length(C2));
          t1:=(CompareUTF8Sequences(@cMatchString[aMatchOffset],c1)>=0) and (CompareUTF8Sequences(@cMatchString[aMatchOffset],c2)<=0);
          if t1 then begin
            if not lNegateCharGroup then begin
              //Jump to CharsGroupEnd+1 because if CharsGroupEnd is reached
              //it means that all optional chars and ranges have not matched the string.
              aMaskIndex:=lSkipOnSuccessGroup+1;
              inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
            end else begin
              Result:=TMaskFailCause.MaskNotMatch;
              exit;
            end;
          end
        end;
      TMaskOpCode.AnyCharToNext:
        begin
          // if last is "*", everything in remain data matches
          if aMaskIndex=cMaskCompiledLimit then begin
            Result:=TMaskFailCause.Success;
            exit;
          end;
          if aMatchOffset>cMatchStringLimit then begin
            if aMaskIndex=cMaskCompiledLimit then begin
              Result:=TMaskFailCause.Success;
              exit;
            end;
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          while aMatchOffset<=cMatchStringLimit do begin
            lFailCause:=intfMatches(aMatchOffset,aMaskIndex);
            if lFailCause=TMaskFailCause.Success then begin
              Result:=TMaskFailCause.Success;
              exit;
            end else if lFailCause=TMaskFailCause.MatchStringExhausted then begin
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
          end;
          Result:=TMaskFailCause.MatchStringExhausted;
          exit;
        end;
      TMaskOpCode.AnyCharOrNone:
        begin
          inc(aMaskIndex);
          lTryCounter:=PInteger(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex,sizeof(integer));
          if TMaskOpCode(cMaskCompiled[aMaskIndex])<>TMaskOpCode.CharsGroupEnd then begin
            Exception_InternalError();
          end else begin
            aMaskIndex:=lSkipOnSuccessGroup+1;
          end;

          // Try to match remain mask eating, 0,1,2,...,lTryCounter chars.
          for j := 0 to lTryCounter do begin
            if aMatchOffset>cMatchStringLimit then begin
              if aMaskIndex=cMaskCompiledLimit+1 then begin
                Result:=TMaskFailCause.Success;
                exit;
              end;
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            lFailCause:=intfMatches(aMatchOffset,aMaskIndex);
            if lFailCause=TMaskFailCause.Success then begin
              Result:=TMaskFailCause.Success;
              exit;
            end else if lFailCause=TMaskFailCause.MatchStringExhausted then begin
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            inc(aMatchOffset,UTF8Length(@cMatchString[aMatchOffset]));
          end;
          Result:=TMaskFailCause.MatchStringExhausted;
          exit;
        end;
      else
        begin
          Exception_InternalError();
        end;
    end;
  end;
  if (aMaskIndex>cMaskCompiledLimit) and (aMatchOffset>cMatchStringLimit) then begin
    Result:=TMaskFailCause.Success;
  end else begin
    if aMaskIndex>cMaskCompiledLimit then begin
      Result:=TMaskFailCause.MaskExhausted;
    end else begin
      Result:=TMaskFailCause.MatchStringExhausted;
    end;
  end;
end;

constructor TMaskUTF8.Create(const aMask: RawByteString;
  const aCaseSensitive: Boolean);
begin
  inherited Create(aCaseSensitive);
  FOriginalMask:=aMask;
end;

constructor TMaskUTF8.CreateAdvanced(const aMask: RawByteString;
  const aCaseSensitive: Boolean; const aOpcodesAllowed: TMaskOpcodesSet);
begin
  inherited CreateAdvanced(aCaseSensitive,aOpcodesAllowed);
  FOriginalMask:=aMask;
end;

function TMaskUTF8.Matches(const aStringToMatch: RawByteString): Boolean;
begin
  if not cMaskIsCompiled then Compile;
  if not FCaseSensitive then begin
    cMatchString:=LowerCase(aStringToMatch);
  end else begin
    cMatchString:=aStringToMatch;
  end;
  cMatchStringLimit:=length(cMatchString);
  if (cMatchStringLimit>=cMatchMinimumLiteralBytes) and (cMatchStringLimit<=cMatchMaximumLiteralBytes) then begin
    if intfMatches(1,0)=TMaskFailCause.Success then begin
      Result:=true;
    end else begin
      Result:=false;
    end;
  end else begin
    // There are too many or not enough bytes to match the string
    Result:=false;
  end;
end;

{ TMaskUnicode }

class function TMaskUnicode.InUnicodeChars(aUnicodeChar: UnicodeChar; aCharSet: array of UnicodeChar): Boolean; static;
var
  k: integer;
begin
  Result:=false;
  for k := 0 to High(aCharSet) do begin
    if aUnicodeChar=aCharSet[k] then begin
      Result:=true;
      Break;
    end;
  end;
end;

procedure TMaskUnicode.Compile;
var
  j: Integer;
  lCharsGroupInsertSize: integer;
  lCPLength: integer;
  lLast: TMaskOpCode;
  lMask: UnicodeString;
  lFirstRange, lSecondRange: integer;

  function IsARange(aPosition: integer; out aFirstSequence: integer; out aSecondSequence: integer): Boolean; {$IFDEF USE_INLINE}inline;{$ENDIF}
  var
    llCPL: integer;
  begin
    Result:=false;
    aFirstSequence:=0;
    aSecondSequence:=0;
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[aPosition]=FMaskEscapeChar) then begin
      llCPL:=UTF16Length(@lMask[aPosition]);
      if aPosition+llCPL>cMaskLimit then exit; // ==>
      inc(aPosition,llCPL);
      aFirstSequence:=aPosition;
    end else begin
      aFirstSequence:=aPosition;
    end;
    llCPL:=UTF16Length(@lMask[aPosition]);
    inc(aPosition,llCPL);
    if aPosition+llCPL>cMaskLimit then exit; // ==>
    // It is not a range, return false
    if lMask[aPosition]<>'-' then exit;

    llCPL:=UTF16Length(@lMask[aPosition]);
    inc(aPosition,llCPL);
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[aPosition]=FMaskEscapeChar) then begin
      llCPL:=UTF16Length(@lMask[aPosition]);
      if aPosition+llCPL>cMaskLimit then exit; // ==>
      inc(aPosition,llCPL);
      aSecondSequence:=aPosition;
    end else begin
      aSecondSequence:=aPosition;
    end;
    Result:=true;
  end;

begin
  inherited Compile;
  if not FCaseSensitive then begin
    lMask:=LowerCase(FOriginalMask);
  end else begin
    lMask:=FOriginalMask;
  end;
  cMaskLimit:=Length(lMask);
  lLast:=TMaskOpCode.Literal;
  SetLength(cMaskCompiled,0);
  j:=1;
  while j<=cMaskLimit do begin
    lCPLength:=UTF16Length(@lMask[j]);
    if (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) and (lMask[j]=FMaskEscapeChar) then begin
      // next is Literal
      inc(j,lCPLength);
      if j<=cMaskLimit then begin
        lCPLength:=UTF16Length(@lMask[j]);
        Add(TMaskOpCode.Literal);
        Add(lCPLength*UTF16_CP_BYTES,@lMask[j]);
        inc(cMatchMinimumLiteralBytes,lCPLength*UTF16_CP_BYTES);
        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength*UTF16_CP_BYTES);
        lLast:=TMaskOpCode.Literal;
        inc(j,lCPLength);
      end else begin
        Exception_IncompleteMask();
      end;
    end else begin
      if InUnicodeChars(lMask[j],['*','?','[']) then begin
        case lMask[j] of
          '*':
            begin
              if eMaskOpcodeAnyText in FMaskOpcodesAllowed then begin
                if lLast<>TMaskOpCode.AnyCharToNext then begin
                  Add(TMaskOpCode.AnyCharToNext);
                  lLast:=TMaskOpCode.AnyCharToNext;
                  // * = No limit
                  cMatchMaximumLiteralBytes:=High(cMatchMaximumLiteralBytes);
                end;
              end else begin
                Add(TMaskOpCode.Literal);
                Add(lCPLength,@lMask[j]);
                inc(cMatchMinimumLiteralBytes,lCPLength*UTF16_CP_BYTES);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength*UTF16_CP_BYTES);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
          '?':
            begin
              if eMaskOpcodeAnyChar in FMaskOpcodesAllowed then begin
                Add(TMaskOpCode.AnyChar);
                lLast:=TMaskOpCode.AnyChar;
                inc(cMatchMinimumLiteralBytes,1*UTF16_CP_BYTES);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,2*UTF16_CP_BYTES);
              end else begin
                Add(TMaskOpCode.Literal);
                Add(lCPLength,@lMask[j]);
                inc(cMatchMinimumLiteralBytes,lCPLength*UTF16_CP_BYTES);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength*UTF16_CP_BYTES);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
          '[':
            begin
              if (eMaskOpcodeOptionalChar in FMaskOpcodesAllowed) or
                 (eMaskOpcodeRange in FMaskOpcodesAllowed) or
                 (eMaskOpcodeAnyCharOrNone in FMaskOpcodesAllowed)
                 then begin
                lLast:=TMaskOpCode.CharsGroupBegin;
                Add(TMaskOpCode.CharsGroupBegin);
                inc(cMatchMinimumLiteralBytes,1*UTF16_CP_BYTES);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,2*UTF16_CP_BYTES);
                lCharsGroupInsertSize:=cMaskCompiledIndex;
                Add(0);
                inc(j,lCPLength); // CP length is 1 because it is "["
                if j<cMaskLimit then begin
                  if (lMask[j]='!') and (eMaskOpcodeNegateGroup in FMaskOpcodesAllowed) then begin
                    Add(TMaskOpCode.Negate);
                    inc(j); // CP length is 1 because it is "!"
                    lLast:=TMaskOpCode.Negate;
                  end;
                end;

                while j<=cMaskLimit do begin
                  lCPLength:=UTF16Length(@lMask[j]);

                  if (lMask[j]='?') and (eMaskOpcodeAnyCharOrNone in FMaskOpcodesAllowed) then begin
                    // This syntax is permitted [??] but not this one [?a] or [a?]
                    if (lLast=TMaskOpCode.CharsGroupBegin) or (lLast=TMaskOpCode.AnyCharOrNone) then begin
                      if lLast=TMaskOpCode.AnyCharOrNone then begin
                        // Increment counter
                        IncrementLastCounterBy(TMaskOpCode.AnyCharOrNone,1);
                      end else begin
                        Add(TMaskOpCode.AnyCharOrNone);
                        Add(1); // Counter
                        // Discount minimun bytes added at the "CharGroupBegin"
                        // because [?] could be 1 or 0 chars, so minimum is zero
                        // but the CharsGroupBegin assumes 1 char as all other
                        // masks replace the group by 1 char position.
                        // This code will run 1 time per group at maximun.
                        dec(cMatchMinimumLiteralBytes,2);
                        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then dec(cMatchMaximumLiteralBytes,4);
                      end;
                      if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,4);
                      lLast:=TMaskOpCode.AnyCharOrNone;
                    end else begin
                      Exception_InvalidCharMask(string(lMask[j]),j);
                    end;

                  end else if (lLast=TMaskOpCode.AnyCharOrNone) and (lMask[j]<>']') then begin
                    //lMask[j] is not '?', but previous mask was '?' and it is an invalid sequence.
                    // "[??] = Valid" // "[a?] or [?a] = Invalid"
                    Exception_InvalidCharMask(string(lMask[j]),j);

                  end else if (eMaskOpcodeRange in FMaskOpcodesAllowed) and IsARange(j,lFirstRange,lSecondRange) then begin
                    Add(TMaskOpCode.Range);
                    // Check if reverse range is needed
                    {$IFDEF RANGES_AUTOREVERSE}
                    if CompareUTF16Sequences(@lMask[lFirstRange],@lMask[lSecondRange])<0 then begin
                      lCPLength:=UTF16Length(@lMask[lFirstRange]);
                      Add(lCPLength*UTF16_CP_BYTES,@lMask[lFirstRange]);
                      lCPLength:=UTF16Length(@lMask[lSecondRange]);
                      Add(lCPLength*UTF16_CP_BYTES,@lMask[lSecondRange]);
                      j:=lSecondRange;
                    end else begin
                      lCPLength:=UTF16Length(@lMask[lSecondRange]);
                      Add(lCPLength*UTF16_CP_BYTES,@lMask[lSecondRange]);
                      Add(UTF16Length(@lMask[lFirstRange])*UTF16_CP_BYTES,@lMask[lFirstRange]);
                      j:=lSecondRange;
                    end;
                    {$ELSE}
                      lCPLength:=UTF16Length(@lMask[lFirstRange]);
                      Add(lCPLength*UTF16_CP_BYTES,@lMask[lFirstRange]);
                      lCPLength:=UTF16Length(@lMask[lSecondRange]);
                      Add(lCPLength*UTF16_CP_BYTES,@lMask[lSecondRange]);
                      j:=lSecondRange;
                    {$ENDIF}
                    lLast:=TMaskOpCode.Range;

                  end else if lMask[j]=']' then begin
                    if (lLast=TMaskOpCode.CharsGroupBegin) or (lLast=TMaskOpCode.Negate) then begin
                      //Error empty match
                      Exception_InvalidCharMask(']',j);
                    end;
                    // Insert the new offset in case of a positive match in CharsGroup
                    PInteger(@cMaskCompiled[lCharsGroupInsertSize])^:=cMaskCompiledIndex;
                    Add(TMaskOpCode.CharsGroupEnd);
                    lLast:=TMaskOpCode.CharsGroupEnd;
                    break;
                  end else begin
                    if (lMask[j]=FMaskEscapeChar) and (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) then begin
                      // next is loptional char in set
                      inc(j,lCPLength);
                      if j<=cMaskLimit then begin
                        lCPLength:=UTF16Length(@lMask[j]);
                      end else begin
                        Exception_IncompleteMask();
                      end;
                    end;
                    if eMaskOpcodeOptionalChar in FMaskOpcodesAllowed then begin
                      Add(TMaskOpCode.OptionalChar);
                      Add(lCPLength*UTF16_CP_BYTES,@lMask[j]);
                      lLast:=TMaskOpCode.OptionalChar;
                    end else begin
                      Exception_InvalidCharMask(lMask[j],j);
                    end;
                  end;
                  inc(j,lCPLength);
                end;
                if j>cMaskLimit then begin
                  Exception_MissingCloseChar(']',cMaskLimit);
                end;
              end else begin
                if (lMask[j]=FMaskEscapeChar) and (eMaskOpcodeEscapeChar in FMaskOpcodesAllowed) then begin
                  // next is Literal
                  inc(j,lCPLength);
                  if j<=cMaskLimit then begin
                    lCPLength:=UTF16Length(@lMask[j]);
                  end else begin
                    Exception_IncompleteMask();
                  end;
                end;
                Add(TMaskOpCode.Literal);
                Add(lCPLength*UTF16_CP_BYTES,@lMask[j]);
                inc(cMatchMinimumLiteralBytes,lCPLength);
                if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength);
                lLast:=TMaskOpCode.Literal;
              end;
            end;
        end;
      end else begin
        // Literal
        Add(TMaskOpCode.Literal);
        Add(lCPLength*UTF16_CP_BYTES,@lMask[j]);
        inc(cMatchMinimumLiteralBytes,lCPLength*UTF16_CP_BYTES);
        if cMatchMaximumLiteralBytes<High(cMatchMaximumLiteralBytes) then inc(cMatchMaximumLiteralBytes,lCPLength*UTF16_CP_BYTES);
        lLast:=TMaskOpCode.Literal;
      end;
      inc(j,lCPLength);
    end;
  end;
  SetLength(cMaskCompiled,cMaskCompiledIndex);
  cMaskCompiledLimit:=cMaskCompiledIndex-1;
end;

class function TMaskUnicode.UTF16Length(const P: PWORD): integer;
begin
  if (P^>=WORD($D800)) and (P^<=WORD($DBFF)) then begin
    Result:=2; // Surrogated pairs
  end else begin
    Result:=1;
  end;
end;

class function TMaskUnicode.CompareUTF16Sequences(const P1, P2: PWORD): integer;
var
  l1,l2: integer;
  l: integer;
begin
  l1:=UTF16Length(p1);
  l2:=UTF16Length(p2);
  Result:=0;
  l:=0;
  while (l<l1) and (l<l2) do begin
    Result:=Integer((P1+l)^)-integer((P2+l)^);
    if Result<>0 then exit;
    inc(l);
  end;
  Result:=l1-l2;
end;

function TMaskUnicode.intfMatches(aMatchOffset: integer; aMaskIndex: integer): TMaskFailCause;
var
  c1,c2: PWORD;
  lFailCause: TMaskFailCause;
  lNegateCharGroup: Boolean;
  lSkipOnSuccessGroup: integer;
  t1: Boolean;
  lTryCounter: integer;
  j: integer;
begin
  lSkipOnSuccessGroup:=0;
  Result:=UnexpectedEnd;
  lNegateCharGroup:=false;
  while aMaskIndex<=cMaskCompiledLimit do begin
    case TMaskOpCode(cMaskCompiled[aMaskIndex]) of
      TMaskOpCode.Literal:
        begin
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          if CompareUTF16Sequences(@cMaskCompiled[aMaskIndex],@cMatchString[aMatchOffset])<>0 then begin
            Result:=TMaskFailCause.MaskNotMatch;
            Exit;
          end;
          inc(aMaskIndex,UTF16Length(@cMaskCompiled[aMaskIndex])*UTF16_CP_BYTES);
          inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
        end;
      TMaskOpCode.AnyChar:
        begin
          inc(aMaskIndex);
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
        end;
      TMaskOpCode.Negate:
        begin
          lNegateCharGroup:=true;
          inc(aMaskIndex);
        end;
      TMaskOpCode.CharsGroupBegin:
        begin
          lNegateCharGroup:=false;
          inc(aMaskIndex);
          lSkipOnSuccessGroup:=PInteger(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex,sizeof(integer));
        end;
      TMaskOpCode.CharsGroupEnd:
        begin
          if lNegateCharGroup then begin
            aMaskIndex:=lSkipOnSuccessGroup+1;
            inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
          end else begin
            Result:=TMaskFailCause.MaskNotMatch;
            exit;
          end;
        end;
      TMaskOpCode.OptionalChar:
        begin
          inc(aMaskIndex);
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          if CompareUTF16Sequences(@cMaskCompiled[aMaskIndex],@cMatchString[aMatchOffset])=0 then begin
            if lNegateCharGroup then begin
              Result:=TMaskFailCause.MaskNotMatch;
              exit;
            end;
            aMaskIndex:=lSkipOnSuccessGroup+1;
            inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
          end else begin
            inc(aMaskIndex,UTF16Length(@cMaskCompiled[aMaskIndex])*UTF16_CP_BYTES);
          end;
        end;
      TMaskOpCode.Range:
        begin
          if aMatchOffset>cMatchStringLimit then begin
            // Error, no char to match.
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          c1:=@cMaskCompiled[aMaskIndex];
          inc(aMaskIndex,UTF16Length(C1)*UTF16_CP_BYTES);
          c2:=@cMaskCompiled[aMaskIndex];
          inc(aMaskIndex,UTF16Length(C2)*UTF16_CP_BYTES);
          t1:=(CompareUTF16Sequences(@cMatchString[aMatchOffset],c1)>=0) and (CompareUTF16Sequences(@cMatchString[aMatchOffset],c2)<=0);
          if t1 then begin
            if not lNegateCharGroup then begin
              //Jump to CharsGroupEnd+1 because if CharsGroupEnd is reached
              //it means that all optional chars and ranges have not matched the string.
              aMaskIndex:=lSkipOnSuccessGroup+1;
              inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
            end else begin
              Result:=TMaskFailCause.MaskNotMatch;
              exit;
            end;
          end
        end;
      TMaskOpCode.AnyCharToNext:
        begin
          // if last is "*", everything in remain data matches
          if aMaskIndex=cMaskCompiledLimit then begin
            Result:=TMaskFailCause.Success;
            exit;
          end;
          if aMatchOffset>cMatchStringLimit then begin
            if aMaskIndex=cMaskCompiledLimit then begin
              Result:=TMaskFailCause.Success;
              exit;
            end;
            Result:=TMaskFailCause.MatchStringExhausted;
            exit;
          end;
          inc(aMaskIndex);
          while aMatchOffset<=cMatchStringLimit do begin
            lFailCause:=intfMatches(aMatchOffset,aMaskIndex);
            if lFailCause=TMaskFailCause.Success then begin
              Result:=TMaskFailCause.Success;
              exit;
            end else if lFailCause=TMaskFailCause.MatchStringExhausted then begin
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
          end;
          Result:=TMaskFailCause.MatchStringExhausted;
          exit;
        end;
      TMaskOpCode.AnyCharOrNone:
        begin
          inc(aMaskIndex);
          lTryCounter:=PInteger(@cMaskCompiled[aMaskIndex])^;
          inc(aMaskIndex,sizeof(integer));
          if TMaskOpCode(cMaskCompiled[aMaskIndex])<>TMaskOpCode.CharsGroupEnd then begin
            Exception_InternalError();
          end else begin
            aMaskIndex:=lSkipOnSuccessGroup+1;
          end;

          // Try to match remain mask eating, 0,1,2,...,lTryCounter chars.
          for j := 0 to lTryCounter do begin
            if aMatchOffset>cMatchStringLimit then begin
              if aMaskIndex=cMaskCompiledLimit+1 then begin
                Result:=TMaskFailCause.Success;
                exit;
              end;
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            lFailCause:=intfMatches(aMatchOffset,aMaskIndex);
            if lFailCause=TMaskFailCause.Success then begin
              Result:=TMaskFailCause.Success;
              exit;
            end else if lFailCause=TMaskFailCause.MatchStringExhausted then begin
              Result:=TMaskFailCause.MatchStringExhausted;
              exit;
            end;
            inc(aMatchOffset,UTF16Length(@cMatchString[aMatchOffset]));
          end;
          Result:=TMaskFailCause.MatchStringExhausted;
          exit;
        end;
      else
        begin
          Exception_InternalError();
        end;
    end;
  end;
  if (aMaskIndex>cMaskCompiledLimit) and (aMatchOffset>cMatchStringLimit) then begin
    Result:=TMaskFailCause.Success;
  end else begin
    if aMaskIndex>cMaskCompiledLimit then begin
      Result:=TMaskFailCause.MaskExhausted;
    end else begin
      Result:=TMaskFailCause.MatchStringExhausted;
    end;
  end;
end;

constructor TMaskUnicode.Create(const aMask: UnicodeString;
  const aCaseSensitive: Boolean);
begin
  inherited Create(aCaseSensitive);
  FOriginalMask:=aMask;
end;

constructor TMaskUnicode.CreateAdvanced(const aMask: UnicodeString;
  const aCaseSensitive: Boolean; const aOpcodesAllowed: TMaskOpcodesSet);
begin
  inherited CreateAdvanced(aCaseSensitive,aOpcodesAllowed);
  FOriginalMask:=aMask;
end;

function TMaskUnicode.Matches(const aStringToMatch: UnicodeString): Boolean;
begin
  if not cMaskIsCompiled then Compile;
  if not FCaseSensitive then begin
    cMatchString:=LowerCase(aStringToMatch);
  end else begin
    cMatchString:=aStringToMatch;
  end;
  cMatchStringLimit:=length(cMatchString);
  if (cMatchStringLimit*UTF16_CP_BYTES>=cMatchMinimumLiteralBytes) and (cMatchStringLimit*UTF16_CP_BYTES<=cMatchMaximumLiteralBytes) then begin
    if intfMatches(1,0)=TMaskFailCause.Success then begin
      Result:=true;
    end else begin
      Result:=false;
    end;
  end else begin
    // There are too many or not enough bytes to match the string
    Result:=false;
  end;
end;

end.

