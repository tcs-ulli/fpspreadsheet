{ This demo show how a user-provided function can be used for calculation of
  rpn formulas that are built-in to fpspreadsheet, but don't have an own
  calculation procedure. }

program test_formula_func;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes, laz_fpspreadsheet
  { you can add units after this },
  math, fpspreadsheet, fpsfunc;

const
  paymentAtEnd = 0;
  paymentAtBegin = 1;

{ Calculates the future value of an investment based on an interest rate and
  a constant payment schedule:
  - "interest_rate" is the interest rate for the investment (as decimal, not percent)
  - "number_periods" is the number of payment periods, i.e. number of payments
    for the annuity.
  - "payment" is the amount of the payment made each period
  - "PV" is the present value of the payments.
  - "payment_type" indicates when the payments are due (see paymentAtXXX constants)
  see: http://en.wikipedia.org/wiki/Future_value
}
function FV(interest_rate, number_periods, payment, pv: Double;
  payment_type: integer): Double;
var
  q: Double;
begin
  q := 1.0 + interest_rate;

  Result := pv * power(q, number_periods) +
            (power(q, number_periods) - 1) / (q - 1) * payment;

  if payment_type = paymentAtBegin then
    Result := Result * q;
end;

function fpsFV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  // Pop the argument off the stack. This can be done by means of PopNumberValues
  // which brings the values back into the right order and reports an error
  // in case of non-numerical values.
  if Args.PopNumberValues(NumArgs, false, data, Result) then
    // Call our FV function with the NumberValues of the arguments.
    Result := CreateNumber(FV(
      data[0],   // interest rate
      data[1],   // number of payments
      data[2],   // payment
      data[3],   // present value
      round(data[4]) // payment type
    ));
end;

const
  INTEREST_RATE = 0.03;
  NUMBER_PAYMENTS = 10;
  PAYMENT = 1000;
  PRESENT_VALUE = 10000;
  PAYMENT_WHEN = paymentAtEnd;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;

begin
  RegisterFormulaFunc(fekFV, @fpsFV);

  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Financial');
    worksheet.Options := worksheet.Options + [soCalcBeforeSaving];
    worksheet.WriteColWidth(0, 20);

    worksheet.WriteUTF8Text(0, 0, 'Interest rate');
    worksheet.WriteNumber(0, 1, INTEREST_RATE, nfPercentage, 1);

    worksheet.WriteUTF8Text(1, 0, 'Number of payments');
    worksheet.WriteNumber(1, 1, NUMBER_PAYMENTS);

    worksheet.WriteUTF8Text(2, 0, 'Payment');
    worksheet.WriteCurrency(2, 1, PAYMENT, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(3, 0, 'Present value');
    worksheet.WriteCurrency(3, 1, PRESENT_VALUE, nfCurrency, 2, '$');

    worksheet.WriteUTF8Text(4, 0, 'Payment at end');
    worksheet.WriteBoolValue(4, 1, PAYMENT_WHEN = paymentAtEnd);

    worksheet.WriteUTF8Text(6, 0, 'Future value');
    worksheet.WriteFontStyle(6, 0, [fssBold]);
    worksheet.WriteUTF8Text(7, 0, 'Our calculation');
    worksheet.WriteCurrency(7, 1,
      FV(INTEREST_RATE, NUMBER_PAYMENTS, PAYMENT, PRESENT_VALUE, PAYMENT_WHEN),
      nfCurrency, 2, '$'
    );

    worksheet.WriteUTF8Text(8, 0, 'Excel''s calculation');
    worksheet.WriteNumberFormat(8, 1, nfCurrency, 2, '$');
    worksheet.WriteRPNFormula(8, 1, CreateRPNFormula(
      RPNNumber(INTEREST_RATE,
      RPNNumber(NUMBER_PAYMENTS,
      RPNNumber(-PAYMENT,
      RPNNumber(-PRESENT_VALUE,
      RPNNumber(PAYMENT_WHEN,
      RPNFunc(fekFV, 5,
      nil))))))));

    workbook.WriteToFile('test_fv.xls', sfExcel8, true);

  finally
    workbook.Free;
  end;
end.

