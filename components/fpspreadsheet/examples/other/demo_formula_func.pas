{ This demo shows how user-provided functions can be used for calculation of
  RPN formulas that are built-in to fpspreadsheet, but don't have their own
  calculation procedure.

  The example will show implementation of some financial formulas:
  - FV()    (future value)
  - PV()    (present value)
  - PMT()   (payment)
  - NPER()  (number of payment periods)

  The demo writes an xls file which uses these formulas and then displays
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
 Classes, SysUtils,
 math, fpspreadsheet, xlsbiff8, fpsfunc, financemath;


{------------------------------------------------------------------------------}
{          Adaption of financial functions to usage by fpspreadsheet           }
{         The functions are implemented in the unit "financemath.pas".         }
{------------------------------------------------------------------------------}

function fpsFV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  // Pop the argument from the stack. This can be done by means of PopNumberValues
  // which brings the values back in the right order and reports an error
  // in case of non-numerical values.
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    // Call the FutureValue function with the NumberValues of the arguments.
    Result := CreateNumberArg(FutureValue(
      data[0],         // interest rate
      round(data[1]),  // number of payments
      data[2],         // payment
      data[3],         // present value
      TPaymentTime(round(data[4]))   // payment type
    ));
end;

function fpsPMT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    Result := CreateNumberArg(Payment(
      data[0],         // interest rate
      round(data[1]),  // number of payments
      data[2],         // present value
      data[3],         // future value
      TPaymentTime(round(data[4]))   // payment type
    ));
end;

function fpsPV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// Present value
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    Result := CreateNumberArg(PresentValue(
      data[0],         // interest rate
      round(data[1]),  // number of payments
      data[2],         // payment
      data[3],         // future value
      TPaymentTime(round(data[4]))   // payment type
    ));
end;

function fpsNPER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    Result := CreateNumberArg(NumberOfPeriods(
      data[0],         // interest rate
      data[1],         // payment
      data[2],         // present value
      data[3],         // future value
      TPaymentTime(round(data[4]))   // payment type
    ));
end;

function fpsRATE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    Result := CreateNumberArg(InterestRate(
      round(data[0]),  // number of payment periods
      data[1],         // payment
      data[2],         // present value
      data[3],         // future value
      TPaymentTime(round(data[4]))   // payment type
    ));
end;

{------------------------------------------------------------------------------}
{        Write xls file comparing our own calculations with Excel result       }
{------------------------------------------------------------------------------}
procedure WriteFile(AFileName: String);
const
  INTEREST_RATE = 0.03;                         // interest rate per period
  NUMBER_PAYMENTS = 10;                         // number of payment periods
  REG_PAYMENT = 1000;                           // regular payment per period
  PRESENT_VALUE = 10000;                        // present value of investment
  PAYMENT_WHEN: TPaymentTime = ptEndOfPeriod;   // when is the payment made

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  fval, pval, pmtval, nperval, rateval: Double;
begin
  { We have to register our financial functions in fpspreadsheet. Otherwise an
    error code would be displayed in the reading part of this demo for these
    formula cells. }
  RegisterFormulaFunc(fekFV, @fpsFV);
  RegisterFormulaFunc(fekPMT, @fpsPMT);
  RegisterFormulaFunc(fekPV, @fpsPV);
  RegisterFormulaFunc(fekNPER, @fpsNPER);
  RegisterFormulaFunc(fekRATE, @fpsRATE);

  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Financial');
    worksheet.Options := worksheet.Options + [soCalcBeforeSaving];
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
    worksheet.WriteRPNFormula(9, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(REG_PAYMENT,
      RPNNumber(PRESENT_VALUE,
      RPNNumber(ord(PAYMENT_WHEN),
      RPNFunc(fekFV, 5,
      nil))))))));
    worksheet.WriteUTF8Text(10, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(10, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(10, 1, CreateRPNFormula(
      RPNCellValue('B2',   // interest rate
      RPNCellValue('B3',   // number of periods
      RPNCellValue('B4',   // payment
      RPNCellValue('B5',   // present value
      RPNCellValue('B6',   // payment at end or at start
      RPNFunc(fekFV, 5,    // Call Excel's FV formula
      nil))))))));

    // present value calculation
    pval := PresentValue(INTEREST_RATE, NUMBER_PAYMENTS, REG_PAYMENT, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(12, 0, 'CALCULATION OF THE PRESENT VALUE');
    worksheet.WriteFontStyle(12, 0, [fssBold]);
    worksheet.WriteUTF8Text(13, 0, 'Direct calculation');
    worksheet.WriteCurrency(13, 1, pval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(14, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(14, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(14, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(REG_PAYMENT,
      RPNNumber(fval,
      RPNNumber(ord(PAYMENT_WHEN),
      RPNFunc(fekPV, 5,
      nil))))))));
    Worksheet.WriteUTF8Text(15, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(15, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(15, 1, CreateRPNFormula(
      RPNCellValue('B2',   // interest rate
      RPNCellValue('B3',   // number of periods
      RPNCellValue('B4',   // payment
      RPNCellValue('B11',  // future value
      RPNCellValue('B6',   // payment at end or at start
      RPNFunc(fekPV, 5,    // Call Excel's PV formula
      nil))))))));

    // payments calculation
    pmtval := Payment(INTEREST_RATE, NUMBER_PAYMENTS, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(17, 0, 'CALCULATION OF THE PAYMENT');
    worksheet.WriteFontStyle(17, 0, [fssBold]);
    worksheet.WriteUTF8Text(18, 0, 'Direct calculation');
    worksheet.WriteCurrency(18, 1, pmtval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(19, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(19, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(19, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(PRESENT_VALUE,
      RPNNumber(fval,
      RPNNumber(ord(PAYMENT_WHEN),
      RPNFunc(fekPMT, 5,
      nil))))))));
    Worksheet.WriteUTF8Text(20, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(20, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(20, 1, CreateRPNFormula(
      RPNCellValue('B2',   // interest rate
      RPNCellValue('B3',   // number of periods
      RPNCellValue('B5',   // present value
      RPNCellValue('B11',  // future value
      RPNCellValue('B6',   // payment at end or at start
      RPNFunc(fekPMT, 5,   // Call Excel's PMT formula
      nil))))))));

    // number of periods calculation
    nperval := NumberOfPeriods(INTEREST_RATE, REG_PAYMENT, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(22, 0, 'CALCULATION OF THE NUMBER OF PAYMENT PERIODS');
    worksheet.WriteFontStyle(22, 0, [fssBold]);
    worksheet.WriteUTF8Text(23, 0, 'Direct calculation');
    worksheet.WriteNumber(23, 1, nperval, nfFixed, 2);

    worksheet.WriteUTF8Text(24, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(24, 1, nfFixed, 2);
    worksheet.WriteRPNFormula(24, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(REG_PAYMENT,
      RPNNumber(PRESENT_VALUE,
      RPNNumber(fval,
      RPNNumber(ord(PAYMENT_WHEN),
      RPNFunc(fekNPER, 5,
      nil))))))));
    Worksheet.WriteUTF8Text(25, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(25, 1, nfFixed, 2);
    worksheet.WriteRPNFormula(25, 1, CreateRPNFormula(
      RPNCellValue('B2',   // interest rate
      RPNCellValue('B4',   // payment
      RPNCellValue('B5',   // present value
      RPNCellValue('B11',  // future value
      RPNCellValue('B6',   // payment at end or at start
      RPNFunc(fekNPER, 5,  // Call Excel's PMT formula
      nil))))))));

    // interest rate calculation
    rateval := InterestRate(NUMBER_PAYMENTS, REG_PAYMENT, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(27, 0, 'CALCULATION OF THE INTEREST RATE');
    worksheet.WriteFontStyle(27, 0, [fssBold]);
    worksheet.WriteUTF8Text(28, 0, 'Direct calculation');
    worksheet.WriteNumber(28, 1, rateval, nfPercentage, 2);

    worksheet.WriteUTF8Text(29, 0, 'Worksheet calculation using constants');
    worksheet.WriteNumberFormat(29, 1, nfPercentage, 2);
    worksheet.WriteRPNFormula(29, 1, CreateRPNFormula(
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(REG_PAYMENT,
      RPNNumber(PRESENT_VALUE,
      RPNNumber(fval,
      RPNNumber(ord(PAYMENT_WHEN),
      RPNFunc(fekRATE, 5,
      nil))))))));
    Worksheet.WriteUTF8Text(30, 0, 'Worksheet calculation using cell values');
    worksheet.WriteNumberFormat(30, 1, nfPercentage, 2);
    worksheet.WriteRPNFormula(30, 1, CreateRPNFormula(
      RPNCellValue('B3',   // number of payments
      RPNCellValue('B4',   // payment
      RPNCellValue('B5',   // present value
      RPNCellValue('B11',  // future value
      RPNCellValue('B6',   // payment at end or at start
      RPNFunc(fekRATE, 5,  // Call Excel's PMT formula
      nil))))))));

    workbook.WriteToFile(AFileName, sfExcel8, true);

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
    workbook.ReadFormulas := true;
    workbook.ReadFromFile(AFilename, sfExcel8);

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

    WriteLn;
    WriteLn('Press [ENTER] to close...');
    ReadLn;
  finally
    workbook.Free;
  end;
end;

const
  TestFile='test_fv.xls';
begin
  WriteFile(TestFile);
  ReadFile(TestFile);
end.


