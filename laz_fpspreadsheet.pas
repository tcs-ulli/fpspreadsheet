{ This file was automatically created by Lazarus. do not edit!
  This source is only used to compile and install the package.
 }

unit laz_fpspreadsheet; 

interface

uses
  fpolestorage, fpsallformats, fpsopendocument, fpspreadsheet, xlsbiff2, 
  xlsbiff5, xlsbiff8, xlsxooxml, fpsutils, fpszipper, LazarusPackageIntf;

implementation

procedure Register; 
begin
end; 

initialization
  RegisterPackage('laz_fpspreadsheet', @Register); 
end.
