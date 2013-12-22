program spreadtestgui;

{$mode objfpc}{$H+}

uses
  Interfaces, Forms, GuiTestRunner,
  datetests, stringtests,
  numberstests, manualtests, testsutility, internaltests, formattests;

begin
  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

