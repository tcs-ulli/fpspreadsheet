program demo_virtualmode_read;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  {$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}
  {$ENDIF}
  Classes, SysUtils, lazutf8, variants,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, xlsxooxml;

type
  TDataAnalyzer = class
    NumberCellCount: integer;
    LabelCellCount: Integer;
    procedure ReadCellDataHandler(Sender: TObject; ARow,ACol: Cardinal;
      const ADataCell: PCell);
  end;

const
  TestFileName = 'test_virtual.xls';

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  dataAnalyzer: TDataAnalyzer;
  t: TTime;

  procedure TDataAnalyzer.ReadCellDataHandler(Sender: TObject;
    ARow, ACol: Cardinal; const ADataCell: PCell);
  { This is just a sample stupidly counting the number and label cells.
    A more serious example could write the cell data to a database. }
  var
    s: String;
  begin
    if ADataCell^.ContentType = cctNumber then
      inc(NumberCellCount);
    if ADataCell^.ContentType = cctUTF8String then
      inc(LabelCellCount);

    // you can use the event handler also to provide feedback on how the process
    // progresses:
    if (ACol = 0) and (ARow mod 1000 = 0) then
      WriteLn('Reading row ', ARow, '...');
  end;

begin
  if not FileExists(TestFileName) then begin
    WriteLn('The test file does not exist. Please run demo_virtual_write first.');
    Halt;
  end;

  dataAnalyzer := TDataAnalyzer.Create;
  try
    workbook := TsWorkbook.Create;
    try
      { These are the essential commands to activate virtual mode: }
      workbook.Options := [boVirtualMode];
//      workbook.Options := [boVirtualMode, buBufStream];
      { boBufStream can be omitted, but is important for large files: it reads
        large pieces of the file to a memory stream from which the data are
        analyzed faster. }

      { The event handler for OnReadCellData links the workbook to the method
        from which analyzes the data. }
      workbook.OnReadCellData := @dataAnalyzer.ReadCellDataHandler;

      t := Now;
      workbook.ReadFromFile(TestFileName);
      t := Now - t;

      WriteLn(Format('The workbook containes %d number and %d label cells, total %d.', [
        dataAnalyzer.NumberCellCount,
        dataAnalyzer.LabelCellCount,
        dataAnalyzer.NumberCellCount + dataAnalyzer.LabelCellCount]));

      WriteLn(Format('Execution time: %.3f sec', [t*24*60*60]));

    finally
      workbook.Free;
    end;

  finally
    dataAnalyzer.Free;
  end;

  WriteLn('Press [ENTER] to quit...');
  ReadLn;
end.

