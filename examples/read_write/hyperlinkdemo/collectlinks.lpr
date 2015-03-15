{ This program seeks all spreadsheet hyperlinks in the file "source.xls" and
  adds the linked worksheet to a new workbook. }

program collectlinks;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes, SysUtils, uriParser,
  fpstypes, fpspreadsheet, fpsUtils, fpsAllFormats
  { you can add units after this };

const
  srcFile = 'source.ods';
  destFile = 'result';

var
  srcWorkbook, destWorkbook, linkedWorkbook: TsWorkbook;
  sheet, linkedSheet, destSheet: TsWorksheet;
  cell: PCell;
  hyperlink: PsHyperlink;
  u: TURI;
  fn: String;  // Name of linked file
  bookmark: String;  // Bookmark of hyperlink
  sheetFormat: TsSpreadsheetFormat;
  sheetName: String;
  r, c: Cardinal;

begin
  // Just for the demo: create the file "source.xls". It contains hyperlinks to
  // some the "test" files created in the XXXXdemo projects
  Write('Creating source workbook...');
  srcWorkbook := TsWorkbook.Create;
  try
    sheet := srcWorkbook.AddWorksheet('Sheet');

    sheet.WriteUTF8Text(0, 0, 'Link to biff8 test file');
    sheet.WriteHyperlink(0, 0, '../excel8demo/test.xls#''My Worksheet 2''!A1');
    //sheet.WriteHyperlink(0, 0, '../excel8demo/test.xls#''Meu Relatório''!A1');

    sheet.WriteUTF8Text(1, 0, 'Link to ods test file');
    sheet.WriteHyperlink(1, 0, '..\opendocdemo\test.ods');

    sheet.WriteUTF8Text(2, 0, 'E-Mail Link');
    sheet.WriteHyperlink(2, 0, 'mailto:someone@mail.com;someoneelse@mail.com?Subject=This is a test');

    sheet.WriteUTF8Text(3, 0, 'Web-Hyperlink');
    sheet.WriteHyperlink(3, 0, 'http://www.lazarus-ide.org/');

    sheet.WriteUTF8Text(4, 0, 'File-Link (absolute path)');
    sheet.WriteHyperlink(4, 0, 'file:///'+ExpandFilename('..\..\..\tests\testooxml_1899.xlsx'));
    // This creates the URI such as "file:///D:\Prog_Lazarus\svn\lazarus-ccr\components\fpspreadsheet\tests\testooxml_1899.xlsx"
    // but makes sure that the file exists on your system.

    sheet.WriteUTF8Text(5, 0, 'Jump to A10');
    sheet.WriteHyperlink(5, 0, '#A10');

    sheet.WriteColWidth(0, 40);

    srcWorkbook.WriteToFile(srcFile, true);
  finally
    srcWorkbook.Free;
  end;
  WriteLn('Done.');

  // Prepare destination workbook
  destWorkbook := nil;

  // Now open the source file and seek hyperlinks
  Write('Reading source workbook, sheet ');
  srcWorkbook := TsWorkbook.Create;
  try
    srcWorkbook.ReadFromFile(srcFile);
    sheet := srcWorkbook.GetWorksheetByIndex(0);
    WriteLn(sheet.Name, '...');

    for cell in sheet.Cells do
    begin
      hyperlink := sheet.FindHyperlink(cell);
      if (hyperlink <> nil) then    // Ignore cells without hyperlink
      begin
        WriteLn;
        WriteLn('Cell ', GetCellString(cell^.Row, cell^.Col), ':');
        WriteLn('  Hyperlink "', hyperlink^.Target, '"');
        if (hyperlink^.Target[1] = '#') then
          WriteLn('  Ignoring internal hyperlink')
        else
        begin
          u := ParseURI(hyperlink^.Target);
          if u.Protocol = '' then begin
            Write('  Local file (relative path)');
            SplitHyperlink(hyperlink^.Target, fn, bookmark)
          end else
          if URIToFileName(hyperlink^.Target, fn) then
            Write('  File (absolute path)')
          else
          begin
            WriteLn('  Ignoring protocol "', u.Protocol, '"');
            continue;  // Ignore http, mailto etc.
          end;

          if not FileExists(fn) then
            WriteLn(' does not exist.')
          else
          if GetFormatFromFileName(fn, sheetFormat) then
          begin
            Write(' supported. ');
            // Create destination workbook if not yet done so far...
            if destWorkbook = nil then
              destWorkbook := TsWorkbook.Create;
            // Open linked workbook
            linkedworkbook := TsWorkbook.Create;
            try
              linkedworkbook.ReadFromFile(fn, sheetFormat);
              // Get linked worksheet
              if bookmark = '' then
                linkedSheet := linkedWorkbook.GetWorksheetByIndex(0)
              else
                if not linkedWorkbook.TryStrToCell(bookmark, linkedsheet, r, c)
                then begin
                  WriteLn('Failure finding linked worksheet.');
                  continue;
                end;
//                linkedSheet := linkedWorkbook.GetWorksheetByName(bookmark);
              // Copy linked worksheet to new sheet in destination workbook
              destSheet := destWorkbook.CopyWorksheetFrom(linkedSheet);
              // Create sheet name
              sheetName := ExtractFileName(fn) + '#' +linkedSheet.Name;
              destWorkbook.ValidWorksheetName(sheetName, true);
              destSheet.Name := sheetName;
              // Done
              WriteLn(' Copied.');
            finally
              linkedworkbook.Free;
            end;
          end;
        end;
      end;
    end;

    // Save destination workbook
    WriteLn;
    if destWorkbook <> nil then
    begin
      destworkbook.WriteToFile(destFile+'.xls', true);
      destworkbook.WriteToFile(destFile+'.xlsx', true);
      destworkbook.WriteToFile(destFile+'.ods', true);
      WriteLn('All hyperlinks to spreadsheets are collected in files ' + destFile + '.*');
    end else
      WriteLn('No hyperlinks found.');

    WriteLn('Press ENTER to close...');
    ReadLn;

  finally
    // Clean up
    srcWorkbook.Free;
    if destWorkbook <> nil then destWorkbook.Free;
  end;

end.

