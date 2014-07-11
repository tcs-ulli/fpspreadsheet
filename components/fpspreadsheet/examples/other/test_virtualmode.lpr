program test_virtualmode;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes, laz_fpspreadsheet,
  { you can add units after this }
  SysUtils, variants, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, xlsxooxml;

type
  TDataProvider = class
    procedure NeedCellData(Sender: TObject; ARow,ACol: Cardinal; var AData: variant);
  end;

  procedure TDataProvider.NeedCellData(Sender: TObject; ARow, ACol: Cardinal;
    var AData: variant);
  { This is just a sample using random data. Normally, in case of a database,
    you would read a record and return its field values, such as:

    Dataset.Fields[ACol].AsVariant := AData;
    if ACol = Dataset.FieldCount then Dataset.Next;
    // NOTE: you have to take care of advancing the database cursor!
  }
  var
    s: String;
    n: Double;
  begin
    if odd(random(10)) then begin
      s := Format('R=%d-C=%d', [ARow, ACol]);
      AData := s;
    end else
      AData := 10000*ARow + ACol;

    // you can use the OnNeedData also to provide feedback on how the process
    // progresses:
    if (ACol = 0) and (ARow mod 1000 = 0) then
      WriteLn('Writing row ', ARow, '...');
  end;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  dataprovider: TDataProvider;

begin

  dataprovider := TDataProvider.Create;
  try
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Sheet1');
      worksheet.WriteFontStyle(0, 1, [fssBold]);

      { These are the essential commands to activate virtual mode: }

//      workbook.WritingOptions := [woVirtualMode, woSaveMemory];
      workbook.WritingOptions := [woVirtualMode];

        // woSaveMemory can be omitted, but is essential for large files: it causes
        // writing temporaray data to a file stream instead of a memory stream.
        // woSaveMemory, however, considerably slows down writing of biff files.

      workbook.VirtualRowCount := 10000;
      workbook.VirtualColCount := 100;
        // These two numbers define the size of virtual spreadsheet.
        // In case of a database, VirtualRowCount is the RecordCount, VirtualColCount
        // the number of fields to be written to the spreadsheet file

      workbook.OnNeedCellData := @dataprovider.NeedCellData;
        // This links the worksheet to the method from which it gets the
        // data to write.

      // In case of a database, you would open the dataset before calling this:
      workbook.WriteToFile('test_virtual.xlsx', sfOOXML, true);
//      workbook.WriteToFile('test_virtual.xls', sfExcel5, true);

    finally
      workbook.Free;
    end;

    WriteLn('Press [ENTER] to quit...');
    ReadLn;
  finally
    dataprovider.Free;
  end;
end.

