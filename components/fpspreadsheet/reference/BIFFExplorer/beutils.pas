unit beUtils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, IniFiles, Forms;

function  CreateIni : TCustomIniFile;
procedure ReadFormFromIni(ini: TCustomIniFile; ASection: String; AForm: TCustomForm);
procedure WriteFormToIni(ini: TCustomIniFile; ASection: String; AForm: TCustomForm);


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

end.

