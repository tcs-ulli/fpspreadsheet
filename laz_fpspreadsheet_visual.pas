{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit laz_fpspreadsheet_visual;

interface

uses
  fpspreadsheetctrls, fpspreadsheetgrid, fpspreadsheetchart, fpsActions, 
  LazarusPackageIntf;

implementation

procedure Register;
begin
  RegisterUnit('fpspreadsheetctrls', @fpspreadsheetctrls.Register);
  RegisterUnit('fpspreadsheetgrid', @fpspreadsheetgrid.Register);
  RegisterUnit('fpspreadsheetchart', @fpspreadsheetchart.Register);
  RegisterUnit('fpsactions', @fpsactions.Register);
end;

initialization
  RegisterPackage('laz_fpspreadsheet_visual', @Register);
end.
