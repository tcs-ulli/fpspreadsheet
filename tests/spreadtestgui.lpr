program spreadtestgui;

{$mode objfpc}{$H+}

{.$DEFINE HEAPTRC}    // Instead of using -gh activate this to write the heap trace to file

uses
 {$IFDEF HEAPTRC}
  HeapTrc, SysUtils,
 {$ENDIF}
  Interfaces, Forms, GuiTestRunner, datetests, stringtests, numberstests,
  manualtests, testsutility, internaltests, formattests, colortests, fonttests,
  optiontests, numformatparsertests, formulatests, rpnFormulaUnit,
  emptycelltests, errortests, virtualmodetests, insertdeletetests,
  celltypetests, sortingtests, copytests, commenttests, enumeratortests;

begin
 {$IFDEF HEAPTRC}
  // Assuming your build mode sets -dDEBUG in Project Options/Other when defining -gh
  // This avoids interference when running a production/default build without -gh

  if FileExists('heap.trc') then
    DeleteFile('heap.trc');
  SetHeapTraceOutput('heap.trc');
 {$ENDIF HEAPTRC}

  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

