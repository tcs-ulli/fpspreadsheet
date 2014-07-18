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
    procedure NeedCellData(Sender: TObject; ARow,ACol: Cardinal;
      var AData: variant; var AStyleCell: PCell);
  end;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  dataprovider: TDataProvider;
  headerTemplate: PCell;
  t: TTime;

  procedure TDataProvider.NeedCellData(Sender: TObject; ARow, ACol: Cardinal;
    var AData: variant; var AStyleCell: PCell);
  { This is just a sample using random data. Normally, in case of a database,
    you would read a record and return its field values, such as:

    Dataset.Fields[ACol].AsVariant := AData;
    if ACol = Dataset.FieldCount then Dataset.Next;
    // NOTE: you have to take care of advancing the database cursor!
  }
  var
    s: String;
  begin
    if ARow = 0 then begin
      AData := Format('Column %d', [ACol + 1]);
      AStyleCell := headerTemplate;
      // This makes the style of the "headerTemplate" cell available to
      // formatting of all virtual cells in row 0.
      // Important: The template cell must be an existing cell in the worksheet.
    end else    {
    if odd(random(10)) then }begin
      s := Format('R=%d-C=%d', [ARow, ACol]);
      AData := s;
    end {else
      AData := 10000*ARow + ACol};

    // you can use the OnNeedData also to provide feedback on how the process
    // progresses:
    if (ACol = 0) and (ARow mod 1000 = 0) then
      WriteLn('Writing row ', ARow, '...');
  end;

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
      { woSaveMemory can be omitted, but is essential for large files: it causes
        writing temporaray data to a file stream instead of a memory stream.
        woSaveMemory, however, considerably slows down writing of biff files. }

      { Next two numbers define the size of virtual spreadsheet.
        In case of a database, VirtualRowCount is the RecordCount, VirtualColCount
        the number of fields to be written to the spreadsheet file }
      workbook.VirtualRowCount := 60000;
      workbook.VirtualColCount := 100;

      { The event handler for OnNeedCellData links the workbook to the method
        from which it gets the data to be written. }
      workbook.OnNeedCellData := @dataprovider.NeedCellData;

      { If we want to change the format of some cells we have to provide this
        format in template cells of the worksheet. In the example, the first
        row whould be in bold letters and have a gray background.
        Therefore, we define a "header template cell" and pass this in the
        NeedCellData event handler.}
      worksheet.WriteFontStyle(0, 0, [fssBold]);
      worksheet.WriteBackgroundColor(0, 0, scSilver);
      headerTemplate := worksheet.FindCell(0, 0);

      worksheet.WriteRowHeight(0, 3);
      worksheet.WriteColWidth(0, 30);
      { In case of a database, you would open the dataset before calling this: }

      t := Now;
      workbook.WriteToFile('test_virtual.xlsx', sfOOXML, true);
      //workbook.WriteToFile('test_virtual.xls', sfExcel8, true);
      t := Now - t;

    finally
      workbook.Free;
    end;

    WriteLn(Format('Execution time: %.3f sec', [t*24*60*60]));
    WriteLn('Press [ENTER] to quit...');
    ReadLn;
  finally
    dataprovider.Free;
  end;
end.

