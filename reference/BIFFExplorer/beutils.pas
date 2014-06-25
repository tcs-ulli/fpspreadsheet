unit beUtils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, IniFiles, Forms,
  fpspreadsheet;

function  CreateIni : TCustomIniFile;
procedure ReadFormFromIni(ini: TCustomIniFile; ASection: String; AForm: TCustomForm);
procedure WriteFormToIni(ini: TCustomIniFile; ASection: String; AForm: TCustomForm);

function GetFileFormatName(AFormat: TsSpreadsheetFormat): String;

implementation

function CreateIni : TCustomIniFile;
var
  cfg : string;
begin
  cfg := GetAppConfigDir(false);
  if not DirectoryExists(cfg) then
    CreateDir(cfg);
  result := TMemIniFile.Create(GetAppConfigFile(false));
end;

procedure ReadFormFromIni(ini: TCustomIniFile; ASection: String; AForm: TCustomForm);
var
  L,T,W,H: Integer;
  isMax: Boolean;
begin
  L := ini.ReadInteger(ASection, 'Left', AForm.Left);
  T := Ini.ReadInteger(ASection, 'Top', AForm.Top);
  W := ini.ReadInteger(ASection, 'Width', AForm.Width);
  H := ini.ReadInteger(ASection, 'Height', AForm.Height);
  isMax := ini.ReadBool(ASection, 'Maximized', AForm.WindowState = wsMaximized);
  if W > Screen.Width then W := Screen.Width;
  if H > Screen.Height then H := Screen.Height;
  if L < 0 then L := 0;
  if T < 0 then T := 0;
  if L + W > Screen.Width then L := Screen.Width - W;
  if T + H > Screen.Height then T := Screen.Height - H;
  AForm.Left := L;
  AForm.Top := T;
  AForm.Width := W;
  AForm.Height := H;
  if IsMax then
    AForm.WindowState := wsMaximized
  else
    AForm.WindowState := wsNormal;
end;


procedure WriteFormToIni(ini: TCustomIniFile; ASection: String; AForm: TCustomForm);
begin
  ini.WriteBool(ASection, 'Maximized', (AForm.WindowState = wsMaximized));
  if AForm.WindowState = wsNormal then begin
    ini.WriteInteger(ASection, 'Left', AForm.Left);
    ini.WriteInteger(ASection, 'Top', AForm.Top);
    ini.WriteInteger(ASection, 'Width', AForm.Width);
    ini.WriteInteger(ASection, 'Height', AForm.Height);
  end;
end;

function GetFileFormatName(AFormat: TsSpreadsheetFormat): string;
begin
  case AFormat of
    sfExcel2              : Result := 'BIFF2';
    { Excel3/4 not supported fpspreadsheet
    sfExcel3              : Result := 'BIFF3';
    sfExcel4              : Result := 'BIFF4';
    }
    sfExcel5              : Result := 'BIFF5';
    sfExcel8              : Result := 'BIFF8';
    sfooxml               : Result := 'OOXML';
    sfOpenDocument        : Result := 'Open Document';
    sfCSV                 : Result := 'CSV';
    sfWikiTable_Pipes     : Result := 'WikiTable Pipes';
    sfWikiTable_WikiMedia : Result := 'WikiTable WikiMedia';
    else                    Result := '-unknown format-';
  end;
end;

end.

