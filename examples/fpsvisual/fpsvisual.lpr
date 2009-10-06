program fpsvisual;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, mainform, LResources, laz_fpspreadsheet_visual
  { you can add units after this };

{$IFDEF WINDOWS}{$R fpsvisual.rc}{$ENDIF}

begin
  {$I fpsvisual.lrs}
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

