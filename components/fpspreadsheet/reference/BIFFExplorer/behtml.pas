unit beHTML;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics;

type
  THTMLHeader = (h1, h2, h3, h4, h5);
  THeaderColors = Array[THTMLHeader] of TColor;

  THTMLDocument = class
  private
    FLines: TStrings;
    FRawMode: Boolean;
    FIndent: Integer;
    function Indent: String;
    function Raw(const AText: String): String;
    function ColorToHTML(AColor: TColor): String;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddEmptyLine;
    procedure AddListItem(const AText: String);
    procedure AddHeader(AHeader: THTMLHeader; const AText: String);
    procedure AddParagraph(const AText: String);
    procedure BeginDocument(const ATitle: String; const AHeaderColors: THeaderColors;
      ARawMode: Boolean=false);
    procedure BeginBulletList;
    procedure BeginNumberedList;
    function Bold(const AText: String): String;
    procedure EndDocument;
    procedure EndBulletList;
    procedure EndNumberedList;
    function Hyperlink(const AText, ALink: String): String;
    function Italic(const AText: String): String;
    property Lines: TStrings read FLines;
  end;

implementation

uses
  StrUtils, LCLIntf;

constructor THTMLDocument.Create;
begin
  inherited;
  FLines := TStringList.Create;
end;

destructor THTMLDocument.Destroy;
begin
  FLines.Free;
  inherited;
end;

procedure THTMLDocument.AddHeader(AHeader: THTMLHeader; const AText: String);
begin
  if FRawMode then
    FLines.Add(Raw(AText))
  else
    FLines.Add(Format('%s<h%d>%s</h%d>', [Indent, ord(AHeader)+1, AText, ord(AHeader)+1]));
end;

                                 (*

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- HTML Codes by Quackit.com -->
<title>
Title appears in the browser's title bar...</title>
<meta name="keywords" content="Separate keywords or phrases with a comma (example: html code generator, generate html, ...)">
<meta name="description" content="Make it nice and short, but descriptive. The description may appear in search engines' search results pages...">
<style type="text/css">
body {background-color:ffffff;background-image:url(http://);background-repeat:no-repeat;background-position:top left;background-attachment:fixed;}
h4{font-family:Arial;color:003366;}
p {font-family:Cursive;font-size:14px;font-style:normal;font-weight:normal;color:000000;}
</style>
</head>
<body>
<h4>Heading goes here...</h4>
<p>Enter your paragraph text here...</p>
</body>
</html>
                               *)

procedure THTMLDocument.BeginDocument(const ATitle: String;
  const AHeaderColors: THeaderColors; ARawMode: Boolean = false);
begin
  FRawMode := ARawMode;
  FLines.Clear;
  if not FRawMode then begin
    FLines.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">');
    FLines.Add('<html>');
    FLines.Add('  <head>');
    FLines.Add('    <meta http-equiv="content-type" content="text/html; charset=UTF-8">');
    FLines.Add('    <title>' + ATitle + '</title>');
    FLines.Add('    <style type="text/css">');
    FLines.Add(Format(
               '      h1{color:%s;}', [ColorToHTML(AHeaderColors[h1])]));
    FLines.Add(Format(
               '      h2{color:%s;}', [ColorToHTML(AHeaderColors[h2])]));
    FLines.Add(Format(
               '      h3{color:%s;}', [ColorToHTML(AHeaderColors[h3])]));
    FLines.Add(Format(
               '      h4{color:%s;}', [ColorToHTML(AHeaderColors[h4])]));
    FLines.Add(Format(
               '      h5{color:%s;}', [ColorToHTML(AHeaderColors[h5])]));
    FLines.Add('    </style>');
    FLines.Add('  </head>');
    FLines.Add('  <body>');
    FIndent := 4;
  end;
end;

procedure THTMLDocument.BeginBulletList;
begin
  if not FRawMode then begin
    FLines.Add(Indent + '<ul>');
    inc(FIndent, 2);
  end;
end;

procedure THTMLDocument.BeginNumberedList;
begin
  if not FRawMode then begin
    FLines.Add(Indent + '<ol>');
    inc(FIndent, 2);
  end;
end;

procedure THTMLDocument.AddEmptyLine;
begin
  if FRawMode then
    FLines.Add('')
  else
    FLines.Add('<br>');
end;

procedure THTMLDocument.AddListItem(const AText: String);
begin
  if FRawMode then
    FLines.Add('- ' + Raw(AText))
  else
    FLines.Add(Indent + '<li>' + AText + '</li>');
end;

procedure THTMLDocument.AddParagraph(const AText: String);
begin
  if FRawMode then
    FLines.Add(Raw(AText))
  else
    FLines.Add(Indent + '<p>' + AText + '</p>');
end;

function THTMLDocument.Bold(const AText: String): String;
begin
  if FRawMode then
    Result := AText
  else
    Result := '<b>' + AText + '</b>';
end;

function THTMLDocument.ColorToHTML(AColor: TColor): String;
var
  tmpRGB: LongInt;
begin
  tmpRGB := ColorToRGB(AColor) ;
  Result := Format('#%.2x%.2x%.2x', [
    GetRValue(tmpRGB),
    GetGValue(tmpRGB),
    GetBValue(tmpRGB)
  ]) ;
end;

procedure THTMLDocument.EndDocument;
begin
  if not FRawMode then begin
    FLines.Add('  </body>');
    FLines.Add('</html>');
  end;
end;

procedure THTMLDocument.EndBulletList;
begin
  if not FRawMode then begin
    dec(FIndent, 2);
    FLines.Add(Indent + '</ul>');
  end;
end;

procedure THTMLDocument.EndNumberedList;
begin
  if not FRawMode then begin
    dec(FIndent, 2);
    FLines.Add(Indent + '</ol>');
  end;
end;

function THTMLDocument.Hyperlink(const AText, ALink: String): String;
begin
  if FRawMode then
    Result := Format('%s (%s)', [AText, ALink])
  else
    Result := Format('<a href="%s">%s</a>', [ALink, AText]);
end;

function THTMLDocument.Indent: String;
begin
  Result := DupeString(' ', FIndent);
end;

function THTMLDocument.Italic(const AText: String): String;
begin
  if FRawMode then
    Result := AText
  else
    Result := '<i>' + AText + '</i>';
end;

function THTMLDocument.Raw(const AText: String): String;
var
  i, n: Integer;
begin
  Result := '';
  if AText = '' then
    exit;
  n := Length(AText);
  i := 1;
  while (i <= n) do begin
    if AText[i] = '<' then
      repeat
        inc(i);
      until (i = n) or (AText[i] = '>')
    else
      Result := Result + AText[i];
    inc(i);
  end;
end;

end.

