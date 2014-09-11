program fpsgrid;

{$mode objfpc}{$H+}

uses
  Interfaces, // this includes the LCL widgetset
  Forms, mainform, LResources, laz_fpspreadsheet_visual;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

