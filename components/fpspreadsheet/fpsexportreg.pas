{
  Registration for fpsexport into the Lazarus component palette
  This requires package lazdbexport for property editors etc
}
unit fpsexportreg;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LazarusPackageIntf, lresources, fpdataexporter;

procedure Register;

implementation

//todo: add component graphic
//{$R fpsexportimg.res}

uses
  fpsexport;

{@@ ----------------------------------------------------------------------------
  Registers the export component in the Lazarus component palette,
  page "Data Export".
-------------------------------------------------------------------------------}
procedure Register;
begin
  RegisterComponents('Data Export',[TFPSExport]);
end;

end.

