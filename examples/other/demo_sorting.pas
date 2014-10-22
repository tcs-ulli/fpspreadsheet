program demo_sorting;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  SysUtils, Classes
  { you can add units after this },
  TypInfo, fpSpreadsheet, fpsutils;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  s: String;
  sortParams: TsSortParams;

  procedure SortSingleColumn;
  var
    i: Integer;
    n: Double;
  begin
    WriteLn('Sorting of a single column');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteNumber(0, 0, 10);        // A1
      worksheet.WriteNumber(1, 0, 2);         // A2
      worksheet.WriteNumber(2, 0, 5);         // A3
      worksheet.WriteNumber(3, 0, 1);         // A4

      sortParams := InitSortParams(true, 1);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 3, 0);

      WriteLn(#9, 'A');
      for i:=0 to 3 do
      begin
        n := worksheet.ReadAsNumber(i, 0);
        WriteLn(i, #9, FloatToStr(n));
      end;
      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortSingleRow;
  var
    i: Integer;
    n: Double;
  begin
    WriteLn('Sorting of a single row');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteNumber(0, 0, 10);        // A1
      worksheet.WriteNumber(0, 1, 2);         // B1
      worksheet.WriteNumber(0, 2, 5);         // C1
      worksheet.WriteNumber(0, 3, 1);         // D1

      sortParams := InitSortParams(false, 1);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 0, 3);

      for i:=0 to 3 do
        Write(char(ord('A')+i) + '1', #9);
      WriteLn;

      for i:=0 to 3 do
      begin
        n := worksheet.ReadAsNumber(0, i);
        Write(FloatToStr(n), #9);
      end;
      WriteLn;
      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortTwoColumns_OneKey;
  var
    i: Integer;
    n1, n2: Double;
  begin
    WriteLn('Sorting of two columns using a single key column');
    WriteLn('(The 2nd column must be the negative of the 1st one)');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteNumber(0, 0, 10);        // A1
      worksheet.WriteNumber(1, 0, 2);         // A2
      worksheet.WriteNumber(2, 0, 5);         // A3
      worksheet.WriteNumber(3, 0, 1);         // A4

      worksheet.WriteNumber(0, 1, -10);        // B1
      worksheet.WriteNumber(1, 1, -2);         // B2
      worksheet.WriteNumber(2, 1, -5);         // B3
      worksheet.WriteNumber(3, 1, -1);         // B4

      sortParams := InitSortParams(true, 1);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 3, 1);

      WriteLn(#9, 'A', #9, 'B');
      for i:=0 to 3 do
      begin
        n1 := worksheet.ReadAsNumber(i, 0);
        n2 := worksheet.ReadAsNumber(i, 1);
        WriteLn(i, #9, FloatToStr(n1), #9, FloatToStr(n2));
      end;
      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortTwoRows_OneKey;
  var
    i: Integer;
    n1, n2: Double;
  begin
    WriteLn('Sorting of two rows using a single key column');
    WriteLn('(The 2nd row must be the negative of 1st row)');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteNumber(0, 0, 10);        // A1
      worksheet.WriteNumber(0, 1, 2);         // B1
      worksheet.WriteNumber(0, 2, 5);         // C1
      worksheet.WriteNumber(0, 3, 1);         // D1

      worksheet.WriteNumber(1, 0, -10);       // A2
      worksheet.WriteNumber(1, 1, -2);        // B2
      worksheet.WriteNumber(1, 2, -5);        // C2
      worksheet.WriteNumber(1, 3, -1);        // D2

      sortParams := InitSortParams(false, 1);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 1, 3);

      Write(#9);
      for i:=0 to 3 do
        Write(char(ord('A')+i) + '1', #9);
      WriteLn;

      Write('1', #9);
      for i:=0 to 3 do begin
        n1 := worksheet.ReadAsNumber(0, i);
        Write(FloatToStr(n1), #9);
      end;
      WriteLn;

      Write('2', #9);
      for i:=0 to 3 do begin
        n1 := worksheet.ReadAsNumber(1, i);
        Write(FloatToStr(n1), #9);
      end;
      WriteLn;

      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortTwoColumns_TwoKeys;
  var
    i: Integer;
    n1, n2: Double;
  begin
    WriteLn('Sorting of two columns on column "A" and "B"');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteNumber(0, 0, 10);         // A1
      worksheet.WriteNumber(1, 0, 1);          // A2
      worksheet.WriteNumber(2, 0, 1);          // A3
      worksheet.WriteNumber(3, 0, 1);          // A4

      worksheet.WriteNumber(0, 1, -10);        // B1
      worksheet.WriteNumber(1, 1, -2);         // B2
      worksheet.WriteNumber(2, 1, -5);         // B3
      worksheet.WriteNumber(3, 1, -1);         // B4

      sortParams := InitSortParams(true, 2);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;
      sortParams.Keys[1].ColRowIndex := 1;
      sortParams.Keys[1].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 3, 1);

      WriteLn(#9, 'A', #9, 'B');
      for i:=0 to 3 do
      begin
        n1 := worksheet.ReadAsNumber(i, 0);
        n2 := worksheet.ReadAsNumber(i, 1);
        WriteLn(i, #9, FloatToStr(n1), #9, FloatToStr(n2));
      end;
      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortTwoRows_TwoKeys;
  var
    i: Integer;
    n1, n2: Double;
  begin
    WriteLn('Sorting of two rows on row "1" and "2"');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteNumber(0, 0, 10);         // A1
      worksheet.WriteNumber(0, 1, 1);          // B1
      worksheet.WriteNumber(0, 2, 1);          // C1
      worksheet.WriteNumber(0, 3, 1);          // D1

      worksheet.WriteNumber(1, 0, -10);        // A2
      worksheet.WriteNumber(1, 1, -2);         // B2
      worksheet.WriteNumber(1, 2, -5);         // C2
      worksheet.WriteNumber(1, 3, -1);         // D2

      sortParams := InitSortParams(false, 2);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;
      sortParams.Keys[1].ColRowIndex := 1;
      sortParams.Keys[1].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 1, 3);

      Write(#9);
      for i:=0 to 3 do
        Write(char(ord('A')+i) + '1', #9);
      WriteLn;

      Write('1', #9);
      for i:=0 to 3 do begin
        n1 := worksheet.ReadAsNumber(0, i);
        Write(FloatToStr(n1), #9);
      end;
      WriteLn;

      Write('2', #9);
      for i:=0 to 3 do begin
        n1 := worksheet.ReadAsNumber(1, i);
        Write(FloatToStr(n1), #9);
      end;
      WriteLn;

      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortTwoColumns_TwoKeys_1;
  var
    i: Integer;
    n: Double;
    s: String;
  begin
    WriteLn('Sorting of two columns on column "A" and "B"');
    WriteLn('(Expecting an ascending columns of characters and numbers)');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteUTF8Text(0, 0, 'E');
      worksheet.WriteUTF8Text(1, 0, 'E');
      worksheet.WriteUTF8Text(2, 0, 'C');
      worksheet.WriteUTF8Text(3, 0, 'B');
      worksheet.WriteUTF8Text(4, 0, 'D');
      worksheet.WriteUTF8Text(5, 0, 'D');
      worksheet.WriteUTF8Text(6, 0, 'A');
      worksheet.WriteUTF8Text(7, 0, 'B');
      worksheet.WriteUTF8Text(8, 0, 'C');
      worksheet.WriteUTF8Text(9, 0, 'A');

      worksheet.WriteNumber(0, 1, 9);         // A2        --> E
      worksheet.WriteNumber(1, 1, 8);         // B2        --> E
      worksheet.WriteNumber(2, 1, 5);         // C2        --> C
      worksheet.WriteNumber(3, 1, 2);         // D2        --> B
      worksheet.WriteNumber(4, 1, 6);         // E2        --> D
      worksheet.WriteNumber(5, 1, 7);         // F2        --> D
      worksheet.WriteNumber(6, 1, 1);         // G2        --> A
      worksheet.WriteNumber(7, 1, 3);         // H2        --> B
      worksheet.WriteNumber(8, 1, 4);         // I2        --> C
      worksheet.WriteNumber(9, 1, 0);         // J2        --> A

      sortParams := InitSortParams(true, 2);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;
      sortParams.Keys[1].ColRowIndex := 1;
      sortParams.Keys[1].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 9, 1);

      WriteLn(#9, 'A', #9, 'B');
      for i:=0 to 9 do
      begin
        s := worksheet.ReadAsUTF8Text(i, 0);
        n := worksheet.ReadAsNumber(i, 1);
        WriteLn(i, #9, s, #9, FloatToStr(n));
      end;

      WriteLn;

    finally
      workbook.Free;
    end;
  end;

  procedure SortTwoRows_TwoKeys_1;
  var
    i: Integer;
    n1, n2: Double;
  begin
    WriteLn('Sorting of two rows on row "1" and "2"');
    WriteLn('(Expecting an ascending row of numbers)');
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet('Test');

      worksheet.WriteUTF8Text(0, 0, 'E');
      worksheet.WriteUTF8Text(0, 1, 'E');
      worksheet.WriteUTF8Text(0, 2, 'C');
      worksheet.WriteUTF8Text(0, 3, 'B');
      worksheet.WriteUTF8Text(0, 4, 'D');
      worksheet.WriteUTF8Text(0, 5, 'D');
      worksheet.WriteUTF8Text(0, 6, 'A');
      worksheet.WriteUTF8Text(0, 7, 'B');
      worksheet.WriteUTF8Text(0, 8, 'C');
      worksheet.WriteUTF8Text(0, 9, 'A');

      worksheet.WriteNumber(1, 0, 9);         // A2        --> E
      worksheet.WriteNumber(1, 1, 8);         // B2        --> E
      worksheet.WriteNumber(1, 2, 5);         // C2        --> C
      worksheet.WriteNumber(1, 3, 2);         // D2        --> B
      worksheet.WriteNumber(1, 4, 6);         // E2        --> D
      worksheet.WriteNumber(1, 5, 7);         // F2        --> D
      worksheet.WriteNumber(1, 6, 1);         // G2        --> A
      worksheet.WriteNumber(1, 7, 3);         // H2        --> B
      worksheet.WriteNumber(1, 8, 4);         // I2        --> C
      worksheet.WriteNumber(1, 9, 0);         // J2        --> A

      sortParams := InitSortParams(false, 2);
      sortParams.Keys[0].ColRowIndex := 0;
      sortParams.Keys[0].Order := ssoAscending;
      sortParams.Keys[1].ColRowIndex := 1;
      sortParams.Keys[1].Order := ssoAscending;

      worksheet.Sort(sortParams, 0, 0, 1, 9);

      Write(#9);
      for i:=0 to 9 do
        Write(char(ord('A')+i) + '1', #9);
      WriteLn;

      Write('1', #9);
      for i:=0 to 9 do
        Write(worksheet.ReadAsUTF8Text(0, i), #9);
      WriteLn;

      Write('2', #9);
      for i:=0 to 9 do begin
        n1 := worksheet.ReadAsNumber(1, i);
        Write(FloatToStr(n1), #9);
      end;
      WriteLn;

      WriteLn;

    finally
      workbook.Free;
    end;
  end;

begin
  SortSingleColumn;
  SortSingleRow;

  SortTwoColumns_OneKey;
  SortTwoRows_OneKey;

  SortTwoColumns_TwoKeys;
  SortTwoRows_TwoKeys;

  SortTwoColumns_TwoKeys_1;
  SortTwoRows_TwoKeys_1;
end.

