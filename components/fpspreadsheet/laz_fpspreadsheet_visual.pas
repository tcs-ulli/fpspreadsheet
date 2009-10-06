{ This file was automatically created by Lazarus. do not edit!
  This source is only used to compile and install the package.
 }

unit laz_fpspreadsheet_visual; 

interface

uses
  fpspreadsheetgrid, LazarusPackageIntf;

implementation

procedure Register; 
begin
  RegisterUnit('fpspreadsheetgrid', @fpspreadsheetgrid.Register); 
end; 

initialization
  RegisterPackage('laz_fpspreadsheet_visual', @Register); 
end.
