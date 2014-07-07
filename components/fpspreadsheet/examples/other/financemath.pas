unit financemath;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

{ Cash flow equation:

    FV + PV * q^n + PMT (q^n - 1) / (q - 1) = 0                (1)

  with

    q   = 1 + interest rate (RATE)
    PV  = present value of an investment
    FV  = future value of an investment
    PMT = regular payment (per period)
    n   = NPER = number of payment periods

  This is valid for payments occuring at the end of each period. If payments
  occur at the start of each period the payments are multiplied by a factor q.
  This case is indicated by means of the parameter PaymentTime below.

  The interest rate is considered as "per period" - whatever that is.
  If the period is 1 year then we use the "usual" interest rate.
  If the period is 1 month then we use 1/12 of the yearly interest rate.

  Sign rules:
  - Money that I receive is to a positive number
  - Money that I pay is to a negative number.

  Example 1: Saving account
  - A saving account has an initial balance of 1000 $ (PV).
    I paid this money to the bank --> negative number
  - I deposit 100$ regularly to this account (PMT): I pay this money --> negative number.
  - At the end of the payments (NPER periods) I get the money back --> positive number.
    This is the FV.

  Example 2: Loan
  - I borrow 1000$ from the bank: I get money --> positive PV
  - I pay 100$ back to the bank in regular intervals --> negative PMT
  - At the end, all debts are paid --> FV = 0.

  The cash flow equation (1) contains 5 parameters (Rate, PV, FV, PMT, NPER).
  The functions below solve this equation always for one of these parameters.

  References:
  - http://en.wikipedia.org/wiki/Time_value_of_money
  - https://wiki.openoffice.org/wiki/Documentation/How_Tos/Calc:_Derivation_of_Financial_Formulas
}

type
  TPaymentTime = (ptEndOfPeriod, ptStartOfPeriod);

function FutureValue(ARate: Extended; NPeriods: Integer;
  APayment, APresentValue: Extended; APaymentTime: TPaymentTime): Extended;

function InterestRate(NPeriods: Integer; APayment, APresentValue, AFutureValue: Extended;
  APaymentTime: TPaymentTime): Extended;

function NumberOfPeriods(ARate, APayment, APresentValue, AFutureValue: Extended;
  APaymentTime: TPaymentTime): Extended;

function Payment(ARate: Extended; NPeriods: Integer;
  APresentValue, AFutureValue: Extended; APaymentTime: TPaymentTime): Extended;

function PresentValue(ARate: Extended; NPeriods: Integer;
  APayment, AFutureValue: Extended; APaymentTime: TPaymentTime): Extended;


implementation

uses
  math;

function FutureValue(ARate: Extended; NPeriods: Integer;
  APayment, APresentValue: Extended; APaymentTime: TPaymentTime): Extended;
var
  q, qn, factor: Extended;
begin
  if ARate = 0 then
    Result := -APresentValue - APayment * NPeriods
  else begin
    q := 1.0 + ARate;
    qn := power(q, NPeriods);
    factor := (qn - 1) / (q - 1);
    if APaymentTime = ptStartOfPeriod then
      factor := factor * q;
    Result := -(APresentValue * qn + APayment*factor);
  end;
end;

function InterestRate(NPeriods: Integer; APayment, APresentValue, AFutureValue: Extended;
  APaymentTime: TPaymentTime): Extended;
{ The interest rate cannot be calculated analytically. We solve the equation
  numerically by means of the Newton method:
  - guess value for the interest reate
  - calculate at which interest rate the tangent of the curve fv(rate)
    (straight line!) has the requested future vale.
  - use this rate for the next iteration. }
const
  DELTA = 0.001;
  EPS = 1E-9;   // required precision of interest rate (after typ. 6 iterations)
  MAXIT = 20;   // max iteration count to protect agains non-convergence
var
  r1, r2, dr: Extended;
  fv1, fv2: Extended;
  iteration: Integer;
begin
  iteration := 0;
  r1 := 0.05;  // inital guess
  repeat
    r2 := r1 + DELTA;
    fv1 := FutureValue(r1, NPeriods, APayment, APresentValue, APaymentTime);
    fv2 := FutureValue(r2, NPeriods, APayment, APresentValue, APaymentTime);
    dr := (AFutureValue - fv1) / (fv2 - fv1) * delta;  // tangent at fv(r)
    r1 := r1 + dr;      // next guess
    inc(iteration);
  until (abs(dr) < EPS) or (iteration >= MAXIT);
  Result := r1;
end;

function NumberOfPeriods(ARate, APayment, APresentValue, AFutureValue: extended;
  APaymentTime: TPaymentTime): extended;
{ Solve the cash flow equation (1) for q^n and take the logarithm }
var
  q, x1, x2: Extended;
begin
  if ARate = 0 then
    Result := -(APresentValue + AFutureValue) / APayment
  else begin
    q := 1.0 + ARate;
    if APaymentTime = ptStartOfPeriod then
      APayment := APayment * q;
    x1 := APayment - AFutureValue * ARate;
    x2 := APayment + APresentValue * ARate;
    if   (x2 = 0)                    // we have to divide by x2
      or (sign(x1) * sign(x2) < 0)   // the argument of the log is negative
    then
      Result := Infinity
    else begin
      Result := ln(x1/x2) / ln(q);
    end;
  end;
end;

function Payment(ARate: Extended; NPeriods: Integer;
  APresentValue, AFutureValue: Extended; APaymentTime: TPaymentTime): Extended;
var
  q, qn, factor: Extended;
begin
  if ARate = 0 then
    Result := -(AFutureValue + APresentValue) / NPeriods
  else begin
    q := 1.0 + ARate;
    qn := power(q, NPeriods);
    factor := (qn - 1) / (q - 1);
    if APaymentTime = ptStartOfPeriod then
      factor := factor * q;
    Result := -(AFutureValue + APresentValue * qn) / factor;
  end;
end;

function PresentValue(ARate: Extended; NPeriods: Integer;
  APayment, AFutureValue: Extended; APaymentTime: TPaymentTime): Extended;
var
  q, qn, factor: Extended;
begin
  if ARate = 0.0 then
    Result := -AFutureValue - APayment * NPeriods
  else begin
    q := 1.0 + ARate;
    qn := power(q, NPeriods);
    factor := (qn - 1) / (q - 1);
    if APaymentTime = ptStartOfPeriod then
      factor := factor * q;
    Result := -(AFutureValue + APayment*factor) / qn;
  end;
end;

end.

