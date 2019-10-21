unit DCPMAIN;

interface

uses
  SysUtils, Classes, DCPmd5;

function dcpcrypt_md5(stream: TStream): AnsiString;

implementation

function dcpcrypt_md5(stream: TStream): AnsiString;
var
  md5: TDCP_md5;
  s: AnsiString;
begin
  md5 := TDCP_md5.Create(nil);
  md5.Init;
  md5.UpdateStream(stream, stream.Size);
  SetLength(Result, 16);
  md5.final(PAnsiChar(Result)^);
  SetLength(s, 32);
  BinToHex(PAnsiChar(Result), PAnsiChar(s),16);
  md5.Free;
  Result := PAnsiChar(s);
end;

end.

