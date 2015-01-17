{ This demo shows how user-provided functions can be used for calculation of
  RPN formulas that are built-in to fpspreadsheet, but don't have their own
  calculation procedure.

  The example will show implementation of some financial formulas:
  - FV()    (future value)
  - PV()    (present value)
  - PMT()   (payment)
  - NPER()  (number of payment periods)

  The demo writes a spreadsheet file which uses these formulas and then displays
  the result in a console window. (Open the generated file in Excel or
  Open/LibreOffice and compare).
}

program demo_formula_func;

{$mode delphi}{$H+}

uses
 {$IFDEF UNIX}
 {$IFDEF UseCThreads}
 cthreads,
 {$ENDIF}
 {$ENDIF}
 Classes, SysUtils, math,
 fpstypes, fpspreadsheet, fpsallformats, fpsexprparser, financemath;

{ Base data used in this demonstration }
const
  INTEREST_RATE = 0.03;                         // interest rate per period
  NUMBER_PAYMENTS = 10;                         // number of payment periods
  REG_PAYMENT = 1000;                           // regular payment per period
  PRESENT_VALUE = 10000.0;                      // present value of investment
  PAYMENT_WHEN: TPaymentTime = ptEndOfPeriod;   // when is the payment made

{------------------------------------------------------------------------------}
{          Adaption of financial functions to usage by fpspreadsheet           }
{         The functions are implemented in the unit "financemath.pas".         }
{------------------------------------------------------------------------------}

procedure fpsFV(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result.ResFloat := FutureValue(
    ArgToFloat(Args[0]),                 // interest rate
    ArgToInt(Args[1]),                   // number of payments
    ArgToFloat(Args[2]),                 // payment
    ArgToFloat(Args[3]),                 // present value
    TPaymentTime(ArgToInt(Args[4]))      // payment type
  );
end;

procedure fpsPMT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result.ResFloat := Payment(
    ArgToFloat(Args[0]),                 // interest rate
    ArgToInt(Args[1]),                   // number of payments
    ArgToFloat(Args[2]),                 // present value
    ArgToFloat(Args[3]),                 // future value
    TPaymentTime(ArgToInt(Args[4]))      // payment type
  );
end;

procedure fpsPV(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result.ResFloat := PresentValue(
    ArgToFloat(Args[0]),                 // interest rate
    ArgToInt(Args[1]),                   // number of payments
    ArgToFloat(Args[2]),                 // payment
    ArgToFloat(Args[3]),                 // future value
    TPaymentTime(ArgToInt(Args[4]))      // payment type
  );
end;

procedure fpsNPER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result.ResFloat := NumberOfPeriods(
    ArgToFloat(Args[0]),                 // interest rate
    ArgToFloat(Args[1]),                 // payment
    ArgToFloat(Args[2]),                 // present value
    ArgToFloat(Args[3]),                 // future value
    TPaymentTime(ArgToInt(Args[4]))      // payment type
  );
end;

procedure fpsRATE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result.ResFloat := InterestRate(
    ArgToInt(Args[0]),                   // number of payments
    ArgToFloat(Args[1]),                 // payment
    ArgToFloat(Args[2]),                 // present value
    ArgToFloat(Args[3]),                 // future value
    TPaymentTime(ArgToInt(Args[4]))      // payment type
  );
end;


{------------------------------------------------------------------------------}
{        Write xls file comparing our own calculations with Excel result       }
{------------------------------------------------------------------------------}
procedure WriteFile(AFileName: String);
const
  INT_EXCEL_SHEET_FUNC_PV    = 56;
  INT_EXCEL_SHEET_FUNC_FV    = 57;
  INT_EXCEL_SHEET_FUNC_NPER  = 58;
  INT_EXCEL_SHEET_FUNC_PMT   = 59;
  INT_EXCEL_SHEET_FUNC_RATE  = 60;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  fval, pval, pmtval, nperval, rateval: Double;
  formula: String;
  fs: TFormatSettings;

begin
  { We have to register our financial functions in fpspreadsheet. Otherwise an
    error code would be displayed in the reading part of this demo for these
    formula cells.
    The 1st parameter is the data type of the function result ('F'=float)
    The 2nd parameter shows the data types of the arguments ('F=float, 'I'=integer)
    The 3rd parameter is the Excel ID needed when writing to xls files. (see
    "OpenOffice Documentation of Microsoft Excel File Format", section 3.11)
    The 4th parameter is the address of the function to be used for calculation. }

  RegisterFunction('FV',   'F', 'FIFFI', INT_EXCEL_SHEET_FUNC_FV,   @fpsFV);
  RegisterFunction('PMT',  'F', 'FIFFI', INT_EXCEL_SHEET_FUNC_PMT,  @fpsPMT);
  RegisterFunction('PV',   'F', 'FIFFI', INT_EXCEL_SHEET_FUNC_PV,   @fpsPV);
  RegisterFunction('NPER', 'F', 'FFFFI', INT_EXCEL_SHEET_FUNC_NPER, @fpsNPER);
  RegisterFunction('RATE', 'F', 'IFFFI', INT_EXCEL_SHEET_FUNC_RATE, @fpsRATE);

  // The formula parser requires a point as decimals separator.
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';

  workbook := TsWorkbook.Create;
  try
    //workbook.Options := workbook.Options + [boCalcBeforeSaving];

    worksheet := workbook.AddWorksheet('Financial');
    worksheet.WriteColWidth(0, 40);
    worksheet.WriteColWidth(1, 15);

    worksheet.WriteUTF8Text(0, 0, 'INPUT DATA');
    worksheet.WriteFontStyle(0, 0, [fssBold]);

    worksheet.WriteUTF8Text(1, 0, 'Interest rate');
    worksheet.WriteNumber(1, 1, INTEREST_RATE, nfPercentage, 1);        // B2

    worksheet.WriteUTF8Text(2, 0, 'Number of payments');
    worksheet.WriteNumber(2, 1, NUMBER_PAYMENTS);                       // B3

    worksheet.WriteUTF8Text(3, 0, 'Payment');
    worksheet.WriteCurrency(3, 1, REG_PAYMENT, nfCurrency, 2, '$');     // B4

    worksheet.WriteUTF8Text(4, 0, 'Present value');
    worksheet.WriteCurrency(4, 1, PRESENT_VALUE, nfCurrency, 2, '$');   // B5

    worksheet.WriteUTF8Text(5, 0, 'Payment at end (0) or at begin (1)');
    worksheet.WriteNumber(5, 1, ord(PAYMENT_WHEN));                     // B6

    // future value calculation
    fval := FutureValue(INTEREST_RATE, NUMBER_PAYMENTS, REG_PAYMENT, PRESENT_VALUE, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(7, 0, 'CALCULATION OF THE FUTURE VALUE');
    worksheet.WriteFontStyle(7, 0, [fssBold]);
    worksheet.WriteUTF8Text(8, 0, 'Direct calculation');
    worksheet.WriteCurrency(8, 1, fval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(9, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(9, 1, nfCurrency, 2, '$');
    worksheet.WriteNumberFormat(9, 1, nfCurrency, 2, '$');
    formula := Format('FV(%f,%d,%f,%f,%d)',
      [1.0*INTEREST_RATE, NUMBER_PAYMENTS, 1.0*REG_PAYMENT, 1.0*PRESENT_VALUE, ord(PAYMENT_WHEN)], fs
    );
    worksheet.WriteFormula(9, 1, formula);
    worksheet.WriteUTF8Text(10, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(10, 1, nfCurrency, 2, '$');
    worksheet.WriteFormula(10, 1, 'FV(B2,B3,B4,B5,B6)');

    // present value calculation
    pval := PresentValue(INTEREST_RATE, NUMBER_PAYMENTS, REG_PAYMENT, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(12, 0, 'CALCULATION OF THE PRESENT VALUE');
    worksheet.WriteFontStyle(12, 0, [fssBold]);
    worksheet.WriteUTF8Text(13, 0, 'Direct calculation');
    worksheet.WriteCurrency(13, 1, pval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(14, 0, 'Worksheet calculation using constants');
    formula := Format('PV(%f,%d,%f,%f,%d)',
      [1.0*INTEREST_RATE, NUMBER_PAYMENTS, 1.0*REG_PAYMENT, fval, ord(PAYMENT_WHEN)], fs
    );
    worksheet.WriteNumberFormat(14, 1, nfCurrency, 2, '$');
    worksheet.WriteFormula(14, 1, formula);
    Worksheet.WriteUTF8Text(15, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(15, 1, nfCurrency, 2, '$');
    worksheet.WriteFormula(15, 1, 'PV(B2,B3,B4,B11,B6)');

    // payments calculation
    pmtval := Payment(INTEREST_RATE, NUMBER_PAYMENTS, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(17, 0, 'CALCULATION OF THE PAYMENT');
    worksheet.WriteFontStyle(17, 0, [fssBold]);
    worksheet.WriteUTF8Text(18, 0, 'Direct calculation');
    worksheet.WriteCurrency(18, 1, pmtval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(19, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(19, 1, nfCurrency, 2, '$');
    formula := Format('PMT(%g,%d,%g,%g,%d)',
      [INTEREST_RATE, NUMBER_PAYMENTS, PRESENT_VALUE, fval, ord(PAYMENT_WHEN)], fs
    );
    worksheet.WriteFormula(19, 1, formula);
    Worksheet.WriteUTF8Text(20, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(20, 1, nfCurrency, 2, '$');
    worksheet.WriteFormula(20, 1, 'PMT(B2,B3,B5,B11,B6)');

    // number of periods calculation
    nperval := NumberOfPeriods(INTEREST_RATE, REG_PAYMENT, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(22, 0, 'CALCULATION OF THE NUMBER OF PAYMENT PERIODS');
    worksheet.WriteFontStyle(22, 0, [fssBold]);
    worksheet.WriteUTF8Text(23, 0, 'Direct calculation');
    worksheet.WriteNumber(23, 1, nperval, nfFixed, 2);

    worksheet.WriteUTF8Text(24, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(24, 1, nfFixed, 2);
    formula := Format('NPER(%g,%g,%g,%g,%d)',
      [1.0*INTEREST_RATE, 1.0*REG_PAYMENT, 1.0*PRESENT_VALUE, fval, ord(PAYMENT_WHEN)], fs
    );
    worksheet.WriteFormula(24, 1, formula);
    Worksheet.WriteUTF8Text(25, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(25, 1, nfFixed, 2);
    worksheet.WriteFormula(25, 1, 'NPER(B2,B4,B5,B11,B6)');

    // interest rate calculation
    rateval := InterestRate(NUMBER_PAYMENTS, REG_PAYMENT, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(27, 0, 'CALCULATION OF THE INTEREST RATE');
    worksheet.WriteFontStyle(27, 0, [fssBold]);
    worksheet.WriteUTF8Text(28, 0, 'Direct calculation');
    worksheet.WriteNumber(28, 1, rateval, nfPercentage, 2);

    worksheet.WriteUTF8Text(29, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(29, 1, nfPercentage, 2);
    formula := Format('RATE(%d,%g,%g,%g,%d)',
      [NUMBER_PAYMENTS, 1.0*REG_PAYMENT, 1.0*PRESENT_VALUE, fval, ord(PAYMENT_WHEN)], fs
    );
    worksheet.WriteFormula(29, 1, formula);
    Worksheet.WriteUTF8Text(30, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(30, 1, nfPercentage, 2);
    worksheet.WriteFormula(30, 1, 'RATE(B3,B4,B5,B11,B6)');

    workbook.WriteToFile(AFileName, true);

  finally
    workbook.Free;
  end;
end;

{------------------------------------------------------------------------------}
{                  Read xls file to display Excel's results                    }
{------------------------------------------------------------------------------}
procedure ReadFile(AFileName: String);
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  r: Cardinal;
  s1, s2: String;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boReadFormulas, boAutoCalc];
    workbook.ReadFromFile(AFilename);
    worksheet := workbook.GetFirstWorksheet;

    // Write all cells with contents to the console
    WriteLn('');
    WriteLn('Contents of file "', AFileName, '"');
    WriteLn('');

    for r := 0 to worksheet.GetLastRowIndex do
    begin
      s1 := UTF8ToAnsi(worksheet.ReadAsUTF8Text(r, 0));
      s2 := UTF8ToAnsi(worksheet.ReadAsUTF8Text(r, 1));
      if s1 = '' then
        WriteLn
      else
      if s2 = '' then
        WriteLn(s1)
      else
        WriteLn(s1+': ':50, s2);
    end;

  finally
    workbook.Free;
  end;
end;

const
  TestFile='test_user_formula.xlsx';  // Format depends on extension selected
  // !!!! ods not working yet !!!!

begin
  WriteLn('This demo registers user-defined functions for financial calculations');
  WriteLn('and writes and reads the corresponding spreadsheet file.');
  WriteLn;

  WriteFile(TestFile);
  ReadFile(TestFile);

  WriteLn;
  WriteLn('Open the file in Excel or OpenOffice/LibreOffice.');
  WriteLn('Press [ENTER] to close...');
  ReadLn;
end.


