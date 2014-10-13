{
  Registration for fpsexport into the Lazarus component palette
  This requires package lazdbexport for property editors etc
}
unit fpsexportreg;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LazarusPackageIntf, lresources, fpdataexporter;

Procedure Register;

implementation

//todo: add component graphic
//{$R fpsexportimg.res}

uses
  fpsexport;

Procedure Register;

begin
  RegisterComponents('Data Export',[TFPSExport]);
end;

initialization
  RegisterPackage('Data Export', @Register );

end.

