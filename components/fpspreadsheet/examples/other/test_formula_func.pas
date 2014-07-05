{ This demo show how user-provided functions can be used for calculation of
  rpn formulas that are built-in to fpspreadsheet, but don't have an own
  calculation procedure.

  The example will show implementation of the some financial formulas:
  - FV(...)  (future value)
  - PV(...)  (present value)
  - PMT(...) (payment)

  The demo writes an xls file which uses these formulas and then displays
  the result in a console window. (Open the generated file in Excel or
  Open/LibreOffice and compare).
}

program test_formula_func;

//{$mode objfpc}{$H+}
{$mode delphi}{$H+}

uses
 {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes, SysUtils, laz_fpspreadsheet
  { you can add units after this },
  math, fpspreadsheet, xlsbiff8, fpsfunc;


{------------------------------------------------------------------------------}
{           Basic implmentation of the three financial funtions                }
{------------------------------------------------------------------------------}

const
  paymentAtEnd = 0;
  paymentAtBegin = 1;

{ Calculates the future value of an investment based on an interest rate and
  a constant payment schedule:
  - "interest_rate" is the interest rate for the investment (as decimal, not percent)
  - "number_periods" is the number of payment periods, i.e. number of payments
    for the annuity.
  - "payment" is the amount of the payment made each period
  - "pv" is the present value of the payments.
  - "payment_type" indicates when the payments are due (see paymentAtXXX constants)
  see: http://en.wikipedia.org/wiki/Time_value_of_money

  In Excel's implementation the payments and the FV add up to 0:
      FV + PV q^n + PMT (q^n - 1) / (q - 1) = 0
}
function FV(interest_rate: Double; number_periods: Integer; payment, pv: Double;
  payment_type: integer): Double;
var
  q, qn, factor: Double;
begin
  q := 1.0 + interest_rate;
  qn := power(q, number_periods);
  factor := (qn - 1) / (q - 1);
  if payment_type = paymentAtBegin then
    factor := factor * q;

  Result := -(pv * qn + payment*factor);
end;

{ Calculates the regular payments for a loan based on an interest rate and a
  constant payment schedule
  Arguments as shown for FV(), in addition:
  - "fv" is the future value of the payments.
  see: http://en.wikipedia.org/wiki/Time_value_of_money
  }
function PMT(interest_rate: Double; number_periods: Integer; pv, fv: Double;
  payment_type: Integer): Double;
var
  q, qn, factor: Double;
begin
  q := 1.0 + interest_rate;
  qn := power(q, number_periods);
  factor := (qn - 1) / (q - 1);
  if payment_type = paymentAtBegin then
    factor := factor * q;

  Result := -(fv + pv * qn) / factor;
end;

{ Calculates the present value of an investment based on an interest rate and
  a constant payment schedule.
  Arguments as shown for FV(), in addition:
  - "fv" is the future value of the payments.
  see: http://en.wikipedia.org/wiki/Time_value_of_money
}
function PV(interest_rate: Double; number_periods: Integer; payment, fv: Double;
  payment_type: Integer): Double;
var
  q, qn, factor: Double;
begin
  q := 1.0 + interest_rate;
  qn := power(q, number_periods);
  factor := (qn - 1) / (q - 1);
  if payment_type = paymentAtBegin then
    factor := factor * q;

  Result := -(fv + payment*factor) / qn;
end;


{------------------------------------------------------------------------------}
{                    Adaption for usage by fpspreadsheet                       }
{------------------------------------------------------------------------------}

function fpsFV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  // Pop the argument off the stack. This can be done by means of PopNumberValues
  // which brings the values back into the right order and reports an error
  // in case of non-numerical values.
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    // Call our FV function with the NumberValues of the arguments.
    Result := CreateNumberArg(FV(
      data[0],         // interest rate
      round(data[1]),  // number of payments
      data[2],         // payment
      data[3],         // present value
      round(data[4])   // payment type
    ));
end;

function fpsPMT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    Result := CreateNumberArg(PMT(
      data[0],         // interest rate
      round(data[1]),  // number of payments
      data[2],         // present value
      data[3],         // future value
      round(data[4])   // payment type
    ));
end;

function fpsPV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    Result := CreateNumberArg(PV(
      data[0],         // interest rate
      round(data[1]),  // number of payments
      data[2],         // payment
      data[3],         // future value
      round(data[4])   // payment type
    ));
end;


{------------------------------------------------------------------------------}
{        Write xls file comparing our own calculations with Excel result       }
{------------------------------------------------------------------------------}
procedure WriteFile(AFileName: String);
const
  INTEREST_RATE = 0.03;
  NUMBER_PAYMENTS = 10;
  PAYMENT = 1000;
  PRESENT_VALUE = 10000;
  PAYMENT_WHEN = paymentAtEnd;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  fval, pval, pmtval: Double;

begin
  { We have to register our financial function in fpspreadsheet. Otherwise an
    error code would be displayed in the reading part of this demo in these
    formula cells. }
  RegisterFormulaFunc(fekFV, @fpsFV);
  RegisterFormulaFunc(fekPMT, @fpsPMT);
  RegisterFormulaFunc(fekPV, @fpsPV);

  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Financial');
    worksheet.Options := worksheet.Options + [soCalcBeforeSaving];
    worksheet.WriteColWidth(0, 40);

    worksheet.WriteUTF8Text(0, 0, 'Interest rate');
    worksheet.WriteNumber(0, 1, INTEREST_RATE, nfPercentage, 1);        // B1

    worksheet.WriteUTF8Text(1, 0, 'Number of payments');
    worksheet.WriteNumber(1, 1, NUMBER_PAYMENTS);                       // B2

    worksheet.WriteUTF8Text(2, 0, 'Payment');
    worksheet.WriteCurrency(2, 1, PAYMENT, nfCurrency, 2, '$');         // B3

    worksheet.WriteUTF8Text(3, 0, 'Present value');
    worksheet.WriteCurrency(3, 1, PRESENT_VALUE, nfCurrency, 2, '$');   // B4

    worksheet.WriteUTF8Text(4, 0, 'Payment at end (0) or at begin (1)');
    worksheet.WriteNumber(4, 1, PAYMENT_WHEN);                          // B5

    // future value calculation
    fval := FV(INTEREST_RATE, NUMBER_PAYMENTS, PAYMENT, PRESENT_VALUE, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(6, 0, 'Future value');
    worksheet.WriteFontStyle(6, 0, [fssBold]);
    worksheet.WriteUTF8Text(7, 0, 'Our calculation');
    worksheet.WriteCurrency(7, 1, fval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(8, 0, 'Excel''s calculation using constants');
    worksheet.WriteNumberFormat(8, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(8, 1, CreateRPNFormula(                  // B9
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(PAYMENT,
      RPNNumber(PRESENT_VALUE,
      RPNNumber(PAYMENT_WHEN,
      RPNFunc(fekFV, 5,
      nil))))))));
    worksheet.WriteUTF8Text(9, 0, 'Excel''s calculation using cell values');
    worksheet.WriteNumberFormat(9, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(9, 1, CreateRPNFormula(                  // B9
      RPNCellValue('B1',   // interest rate
      RPNCellValue('B2',   // number of periods
      RPNCellValue('B3',   // payment
      RPNCellValue('B4',   // present value
      RPNCellValue('B5',   // payment at end or at start
      RPNFunc(fekFV, 5,    // Call Excel's FV formula
      nil))))))));

    // present value calculation
    pval := PV(INTEREST_RATE, NUMBER_PAYMENTS, PAYMENT, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(11, 0, 'Present value');
    worksheet.WriteFontStyle(11, 0, [fssBold]);
    worksheet.WriteUTF8Text(12, 0, 'Our calculation');
    worksheet.WriteCurrency(12, 1, pval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(13, 0, 'Excel''s calculation using constants');
    worksheet.WriteNumberFormat(13, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(13, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(PAYMENT,
      RPNNumber(fval,
      RPNNumber(PAYMENT_WHEN,
      RPNFunc(fekPV, 5,
      nil))))))));
    Worksheet.WriteUTF8Text(14, 0, 'Excel''s calculation using cell values');
    worksheet.WriteNumberFormat(14, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(14, 1, CreateRPNFormula(
      RPNCellValue('B1',   // interest rate
      RPNCellValue('B2',   // number of periods
      RPNCellValue('B3',   // payment
      RPNCellValue('B10',  // future value
      RPNCellValue('B5',   // payment at end or at start
      RPNFunc(fekPV, 5,    // Call Excel's PV formula
      nil))))))));

    // payments calculation
    pmtval := PMT(INTEREST_RATE, NUMBER_PAYMENTS, PRESENT_VALUE, fval, PAYMENT_WHEN);
    worksheet.WriteUTF8Text(16, 0, 'Payment');
    worksheet.WriteFontStyle(16, 0, [fssBold]);
    worksheet.WriteUTF8Text(17, 0, 'Our calculation');
    worksheet.WriteCurrency(17, 1, pmtval, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(18, 0, 'Excel''s calculation using constants');
    worksheet.WriteNumberFormat(18, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(18, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(PRESENT_VALUE,
      RPNNumber(fval,
      RPNNumber(PAYMENT_WHEN,
      RPNFunc(fekPMT, 5,
      nil))))))));
    Worksheet.WriteUTF8Text(19, 0, 'Excel''s calculation using cell values');
    worksheet.WriteNumberFormat(19, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(19, 1, CreateRPNFormula(
      RPNCellValue('B1',   // interest rate
      RPNCellValue('B2',   // number of periods
      RPNCellValue('B4',   // present value
      RPNCellValue('B10',  // future value
      RPNCellValue('B5',   // payment at end or at start
      RPNFunc(fekPMT, 5,   // Call Excel's PMT formula
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
    WriteLn('Contents of the first worksheet of the file:');
    WriteLn('');

    for r := 0 to worksheet.GetLastRowIndex do begin
      s1 := UTF8ToAnsi(worksheet.ReadAsUTF8Text(r, 0));
      s2 := UTF8ToAnsi(worksheet.ReadAsUTF8Text(r, 1));
      if s1 = '' then
        WriteLn
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


begin
  WriteFile('test_fv.xls');
  ReadFile('test_fv.xls');
end.


