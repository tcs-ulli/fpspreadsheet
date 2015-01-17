{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit laz_fpspreadsheetexport_visual;

interface

uses
  fpsexport, fpsexportreg, LazarusPackageIntf;

implementation

procedure Register;
begin
  RegisterUnit('fpsexportreg', @fpsexportreg.Register);
end;

initialization
  RegisterPackage('laz_fpspreadsheetexport_visual', @Register);
end.
