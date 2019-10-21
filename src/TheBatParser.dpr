program TheBatParser;

{$APPTYPE CONSOLE}

uses
  Classes,
  SysUtils,
  Windows,
  IdMessage,
  IdAttachment,
  IdText,
  IdMessageParts,
  IdEMailAddress,
  DCPMAIN;

const
  terminator: string = Chr(13) + Chr(10) + '.' + Chr(13) + Chr(10);
  tempname: string = 'TheBatParserTemp';

function StreamWideOpen(Name: WideString): THandleStream;
var
  han: Integer;
begin
  han := CreateFileW(PWideChar(Name), GENERIC_WRITE, 7, nil, CREATE_ALWAYS,
    FILE_ATTRIBUTE_NORMAL, 0);
  if (han = 0) or (han = -1) then
    Abort;
  Result := THandleStream.Create(han);
end;

procedure StreamWideClose(var hand: THandleStream);
var
  han: Integer;
begin
  han := hand.Handle;
  hand.Free;
  hand := nil;
  CloseHandle(han);
end;

var
  curdirs: array[0..31] of WideString;
  curdir: integer;

procedure pushd(dir: WideString);
begin
  if curdir > 31 then
    Abort;
  SetLength(curdirs[curdir], GetCurrentDirectoryW(0, nil));
  GetCurrentDirectoryW(Length(curdirs[curdir]), PWideChar(curdirs[curdir]));
  Inc(curdir);
  if dir <> '' then
    SetCurrentDirectoryW(PWideChar(dir));
end;

procedure pusht();
var
  s: WideString;
begin
  SetLength(s, GetTempPathW(0, nil));
  GetTempPathW(Length(s), PWideChar(s));
  pushd(s);
end;

procedure popd();
begin
  if curdir < 1 then
    Abort;
  Dec(curdir);
  SetCurrentDirectoryW(PWideChar(curdirs[curdir]));
end;

function ValidFileName(Name: WideString): Boolean;
var
  t: WideString;
begin
  t := '_';
  Result := True;
  if Name = t then
    Exit;
  DeleteFileW(PWideChar(Name));
  CloseHandle(CreateFileW(PWideChar(t), GENERIC_WRITE, 0, nil, CREATE_ALWAYS,
    FILE_ATTRIBUTE_NORMAL, 0));
  MoveFileW(PWideChar(t), PWideChar(Name));
  DeleteFileW(PWideChar(t));
  Result := DeleteFileW(PWideChar(Name));
end;

function SecureFileName(const Name: WideString; nodot: Boolean = False): WideString;
var
  i: Integer;
  a: AnsiChar;
  s: AnsiString;

  function range(x, a, b: AnsiChar): Boolean;
  begin
    Result := (Ord(x) >= Ord(a)) and (Ord(x) <= Ord(b));
  end;

begin
  Result := Name;
  s := '.,~+-@#$)[ ]{}!å¸Å¨';
  if nodot then
    s[1] := ' ';
  for i := 1 to Length(Name) do
  begin
    a := AnsiChar(Name[i]);
    if range(a, '0', '9') or range(a, 'a', 'z') or range(a, 'A', 'Z') or range(a,
      'à', 'ÿ') or range(a, 'À', 'ß') or (Pos(a, s) > 0) then
      Result[i] := WideChar(a)
    else
      Result[i] := '_';
  end;
end;

function EscapeFileName(const Name: WideString; max: Integer = 0): WideString;
var
  i, p, s: Integer;
  c: WideChar;
  b, e: WideString;
begin
  Result := Name;
  s := Length(Name);
  p := 0;
  for i := 1 to s do
  begin
    c := Name[i];
    if c = '.' then
      p := i;
    if (c = '/') or (c = '\') or (c = '<') or (c = '>') or (c = '"') or (c = ':')
      or (c = '/') or (c = '|') or (c = '?') or (c = '*') or (Ord(c) < 32) then
      Result[i] := '_'
    else
      Result[i] := c;
  end;
  b := Result;
  e := '';
  if p > 0 then
  begin
    b := Copy(Result, 1, p - 1);
    e := Copy(Result, p, s);
  end;
  if (max > 0) and (s > max) then  // max must be 0 or >=8
  begin
    if Length(e) <= 6 then
      b := Copy(b, 1, max - (Length(e) + 1)) + '~'
    else
    begin
      b := Copy(Result, 1, max - 1) + '~';
      e := '';
    end;
  end;
  Result := b + e;
end;

procedure ExtensionFromName(const Name: WideString; out base: WideString; out
  ext: WideString);
var
  i: Integer;
begin
  for i := Length(Name) downto 1 do
    if Name[i] = '.' then
    begin
      base := Copy(Name, 1, i - 1);
      ext := Copy(Name, i, Length(Name));
      Exit;
    end;
  base := Name;
  ext := '';
end;

function FilenameFromString(const Name: WideString): WideString;
begin
  pusht();
  Result := EscapeFileName(Name, 250);
  if not ValidFileName(Result) then
    Result := SecureFileName(Result);
  if not ValidFileName(Result) then
    Result := SecureFileName(Result, True) + '_';
  popd();
end;

function NewFileName({path: WideString;} Name: WideString): WideString;
var
  n, e: WideString;
  i: Integer;

  function i2s(i: Integer): WideString;
  begin
    if i < 10 then
      Result := '0' + Chr(Ord('0') + i)
    else
      Result := inttostr(i);
  end;

begin
  Result := FilenameFromString(Trim(Name));
  if not FileExists(Result) then
    Exit;
  ExtensionFromName(Result, n, e);
  for i := 1 to 9 do
  begin
    Result := n + '_0' + Chr(Ord('0') + i) + e;
    if not FileExists(Result) then
      Exit;
  end;
  for i := 10 to 99 do
  begin
    Result := n + '_' + inttostr(i) + e;
    if not FileExists(Result) then
      Exit;
  end;
  Result := '';
end;

function FilenameFromContentType(ct: AnsiString): WideString;
var
  p: Integer;
begin
  Result := ct;
  p := Pos(';', ct);
  if p > 0 then
    Delete(Result, p, Length(ct));
end;

function IsTextHtml(const text: string): Boolean;
var
  s: string;
begin
  Result := false;
  if Length(text) < 13 then
    Exit;
  s := LowerCase(Copy(text, Length(text) - 8, 7));
  if (s <> '</html>') and (s <> '</body>') then
    Exit;
  s := LowerCase(Copy(text, 1, 6));
  if (s <> '<html>') and (s <> '<body>') then
    Exit;
  Result := true;
end;

procedure SaveAttachment(att: TIdMessagePart);
var
  Name: WideString;
  s: string;
begin
  Name := att.FileName;
  if Name = '' then
    Name := FilenameFromContentType(att.ContentType);
  if att.PartType = mptText then
  begin
    s := Trim(TIdText(att).Body.text);
    if s = '' then
      Exit;
    if IsTextHtml(s) then
      Name := Name + '.htm'
    else
      Name := Name + '.txt';
  end;
  Name := NewFileName(Name);
  if att.PartType = mptAttachment then
    TIdAttachment(att).SaveToFile(Name)
  else
    TIdText(att).Body.SaveToFile(Name);
end;

function DateToFilename(d: TDateTime): WideString;
var
  s: AnsiString;
begin
  DateTimeToString(s, 'yyyy.mm.dd, hh.nn', d);
  Result := s;
end;

procedure SaveUtf(save: TFileStream; text: WideString); overload;
var
  s: AnsiString;
begin
  s := UTF8Encode(text);
  save.WriteBuffer(PAnsiChar(s)^, Length(s));
end;

procedure SaveUtf(save: TFileStream; text: AnsiString); overload;
begin
  save.WriteBuffer(PAnsiChar(text)^, Length(text));
end;

function PrintAddressList(list: TIdEMailAddressList): string;
var
  i: Integer;
begin
  for i := 0 to list.Count - 1 do
  begin
    Result := Result + list.Items[i].text;
    if i <> list.Count - 1 then
      Result := Result + ', ';
  end;
end;

var
  infile, folder, mode: string;
  stream: TMemoryStream;
  hand: THandleStream;
  save: TFileStream;
  MSG: TIdMessage;
  i: integer;
  s, d, f, t, p: WideString;
  m: AnsiString;
  From: Boolean;

begin
  curdir := 0;
  pushd('');
  if ParamCount <> 3 then
  begin
    Writeln('TheBatParser - export .eml/.msg files to nice filetree');
    Writeln('');
    Writeln('Usage:');
    Writeln('thebatparser.exe "file.msg" "outputdir\" (in|out)');
    Writeln('');
    Writeln('Where:');
    Writeln('file.msg - input email message');
    Writeln('outputdir - output base folder');
    Writeln('mode - "in" to parse as incoming, "out" - as sended');
    Writeln('');
    Writeln('File structure:');
    Writeln('\%base%\%address%\%date%\files...');
    Writeln('%base% - your "outputdir\"');
    Writeln('%address% - sender/recipient like "admin@example.com"');
    Writeln('%date% - timestamp in "yyyy.mm.dd, mm.ss"');
    Writeln('files... - here they are:');
    Writeln('%MD5%.msg - a copy of original input file');
    Writeln('body.txt - contents of the first message body (if not empty)');
    Writeln('%SUBJECT%.txt - UTF-8 extracted headers and texts');
    Writeln('%NAME%.%EXT% - attachments, with original filename and extension');
    Writeln('%TYPE%.txt - alternate bodies, if not empty');
    Writeln('%TYPE%.htm - body part with detected HTML contents');
    Writeln('..\%PERSON%.txt (in parent folder) - a copy of sender/recipient name');
    Writeln('');
    Writeln('Here:');
    Writeln('%MD5% - unique message hash. If already present - then parsing will fail');
    Writeln('%SUBJECT% - message subject');
    Writeln('%TYPE% - part content-type (like "text_plain")');
    Writeln('');
    Writeln('Contents of file %MD5%.txt:');
    Writeln('" Date: ..." - a copy of %date%');
    Writeln('" Subj: ..." - a copy of %SUBJECT%');
    Writeln('" From: ..." - list of senders');
    Writeln('" To:   ..." - list of recipients');
    Writeln('" Sndr: ..." - real sender of a message (if forwarded from)');
    Writeln('" Copy: ..." - Carbon Copy list');
    Writeln('"  BCc: ..." - Blind Carbon copy list (used for outcoming mail)');
    Writeln('" RpTo: ..." - Reply-To address');
    Writeln('" ReRe: ..." - Receipt-Recipient address');
    Writeln('Next there are all textual parts of the message.');
    Writeln('');
    Exit;
  end;
  infile := ParamStr(1);
  folder := ParamStr(2);
  mode := LowerCase(ParamStr(3));
  if not FileExists(infile) then
  begin
    Writeln('File not found: "', infile, '"');
    Exit;
  end;
  if not DirectoryExists(folder) then
  begin
    Writeln('Directory not found: "', folder, '"');
    Exit;
  end;
  if (mode <> 'in') and (mode <> 'out') then
  begin
    Writeln('Wrong mode: "', mode, '"');
    Exit;
  end;
  Write(infile, #9);
  From := (mode = 'in');
  p := folder;
  pusht();
  CreateDir(tempname);
  popd();
  try
    MSG := TIdMessage.Create(nil);

    stream := TMemoryStream.Create;
    stream.LoadFromFile(infile);
    stream.Seek(0, soFromEnd);
    stream.WriteBuffer(PAnsiChar(terminator)^, Length(terminator));
    stream.Seek(0, soFromBeginning);
    MSG.LoadFromStream(stream, False);
    stream.Seek(0, soFromBeginning);
    d := DateToFilename(MSG.Date);
    if From then
      t := WideString(MSG.From.Address)
    else
      t := WideString(MSG.Recipients.Items[0].Address);
    f := FilenameFromString(t);
    pushd(p);
    CreateDirectoryW(PWideChar(f), nil);
    if From then
      t := WideString(MSG.From.Name)
    else
      t := WideString(MSG.Recipients.Items[0].Name);
    pushd(f);
    hand := StreamWideOpen(WideString(FilenameFromString(t) + '.txt'));
    s := MSG.From.text;
    hand.WriteBuffer(PWideChar(s)^, Length(s) * 2);
    StreamWideClose(hand);
    CreateDirectoryW(PWideChar(d), nil);
    pushd(d);
    stream.Seek(0, soFromBeginning);
    m := dcpcrypt_md5(stream);
    if FileExists(m + '.msg') then
    begin
      Writeln('(same)');
      pusht();
      RemoveDir(tempname);
      Halt(0);
    end;
    if Length(MSG.Body.Text) > 0 then
      MSG.Body.SaveToFile(NewFileName('body.txt'));
    for i := 0 to MSG.MessageParts.Count - 1 do
      SaveAttachment(MSG.MessageParts[i]);
    stream.Seek(0, soFromBeginning);
    stream.SaveToFile(m + '.msg');
    stream.Free;
    t := ' Date:' + Chr(9) + d + Chr(13) + Chr(10);
    t := t + ' Subj:' + Chr(9) + MSG.Subject + Chr(13) + Chr(10);
    p := PrintAddressList(MSG.FromList);
    t := t + ' From:' + Chr(9) + p + Chr(13) + Chr(10);
    p := PrintAddressList(MSG.Recipients);
    t := t + ' To:  ' + Chr(9) + p + Chr(13) + Chr(10);
    p := MSG.Sender.text;
    if p <> '' then
      t := t + ' Sndr:' + Chr(9) + p + Chr(13) + Chr(10);
    p := PrintAddressList(MSG.CCList);
    if p <> '' then
      t := t + ' Copy:' + Chr(9) + p + Chr(13) + Chr(10);
    p := PrintAddressList(MSG.BccList);
    if p <> '' then
      t := t + ' BCc: ' + Chr(9) + p + Chr(13) + Chr(10);
    p := PrintAddressList(MSG.ReplyTo);
    if p <> '' then
      t := t + ' RpTo:' + Chr(9) + p + Chr(13) + Chr(10);
    p := MSG.ReceiptRecipient.text;
    if p <> '' then
      t := t + ' ReRe:' + Chr(9) + p + Chr(13) + Chr(10);
    save := TFileStream.Create(NewFileName(MSG.Subject + '.log'), fmCreate);
    save.WriteBuffer(PChar(Chr(239) + Chr(187) + Chr(191))^, 3);
    SaveUtf(save, t + Chr(13) + Chr(10));
    t := Chr(13) + Chr(10) + Chr(13) + Chr(10);
    m := t + MSG.Body.text + t;
    SaveUtf(save, m);
    for i := 0 to MSG.MessageParts.Count - 1 do
      if MSG.MessageParts.Items[i].PartType = mptText then
      begin
        m := TIdText(MSG.MessageParts.Items[i]).Body.text + t;
        SaveUtf(save, m);
      end;
    save.Free;
    popd();
    popd();
    popd();
    Writeln('OK');
  except
    on e: Exception do
      Writeln('Error: ', e.message);
  end;
  pusht();
  RemoveDir(tempname);
end.

