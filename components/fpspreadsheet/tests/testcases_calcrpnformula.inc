{ include file for "formulatests.pas", containing the test cases for the
  calcrpnformula test. }

  // Addition
  Row := 0;
  MyWorksheet.WriteUTF8Text(Row, 0, '=1+1');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNFunc(fekAdd, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1.0+1.0);

  // Subtraction
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=1-10');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(10,
    RPNFunc(fekSub, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1-10);

  // Multiplication
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10*-3');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(-3,
    RPNFunc(fekMul, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(10*(-3));

  // Multiplication  w/Parenthesis
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10*(-3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(-3,
    RPNParenthesis(
    RPNFunc(fekMul, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(10*(-3));

  // Division
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10/200');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(200,
    RPNFunc(fekDiv, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(10/200);

  // Division: Error case - divide by zero
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10/0');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(0,
    RPNFunc(fekDiv, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errDivideByZero);

  // Percentage
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10%');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNFunc(fekPercent, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(10*0.01);

  // Power
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=power(2.0, 0.5)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2.0,
    RPNNumber(0.5,
    RPNFunc(fekPower, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(power(2, 0.5));

{$IFDEF ENABLE_CALC_RPN_EXCEPTIONS}
  // Power: Error case  "power( (negative number), (fractional number) )"
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=power(-2.0, 0.5)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-2.0,
    RPNNumber(0.5,
    RPNFunc(fekPower, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);
{$ENDIF}

  // Unary minus
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=-(-1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-1,
    RPNParenthesis(
    RPNFunc(fekUMinus, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1);

  // Unary plus
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=+(-1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-1,
    RPNParenthesis(
    RPNFunc(fekUPlus, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(-1);

  // String concatenation
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '="Hallo"&" world"');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString(' world',
    RPNFunc(fekConcat, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('Hallo' + ' world');

  // Equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true=false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true = false);

  // Equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"="world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool('Hallo' = 'world');

  // Equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1=1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(1=1);

  // Greater (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true>false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekGreater, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true > false);

  // Greater (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo">"world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekGreater, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool('Hallo' > 'world');

  // Greater (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1>1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekGreater, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(1>1);

  // Greater equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true>=false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekGreaterEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true >= false);

  // Greater equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo">="world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekGreaterEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool('Hallo' >= 'world');

  // Greater equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1>=1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekGreaterEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(1>=1);

  // Less (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true<false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekLess, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true < false);

  // Less (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"<"world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekLess, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool('Hallo' < 'world');

  // Less (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1<1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekLess, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(1<1);

  // Less equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true<=false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekLessEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true <= false);

  // Less equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"<="world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekLessEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool('Hallo' <= 'world');

  // Less equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1<=1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekLessEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(1<=1);

  // Not equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true<>false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekNotEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true <> false);

  // Not equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"<>"world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekNotEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool('Hallo' <> 'world');

  // Not equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1<>1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekNotEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(1<>1);



{------------------------------------------------------------------------------}
{                                     Math                                     }
{------------------------------------------------------------------------------}

  // ABS
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=abs(-1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-1,
    RPNFunc(fekABS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(abs(-1));

  // ARCCOS - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acos(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekACOS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(arccos(0.1));

  // ACOS - error result (arccos is not defined outside the interval [-1..1]
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acos(-2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-2,
    RPNFunc(fekACOS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverFlow);

   // ARCCOSH - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acosh(1.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.1,
    RPNFunc(fekACOSH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(arccosh(1.1));

  // ACOSH - error result (arccos is not defined below 1
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acosh(0.9)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.9,
    RPNFunc(fekACOSH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverFlow);

  // ASIN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=asin(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekASIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(arcsin(0.1));

  // ASIN - error result (arcsin is not defined outside the interval [-1..1]
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=asin(-2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-2,
    RPNFunc(fekASIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverFlow);

   // ARCSINH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=asinh(1.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.1,
    RPNFunc(fekASINH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(arcsinh(1.1));

  // ATAN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=atan(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekATAN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(arctan(0.1));

  // ATANH - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=atanh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekATANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(arctanh(0.1));

  // ATANH - error result (arctan is only defined within ]-1,1[
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=atanh(1.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNFunc(fekATANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverFlow);

  // COS
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=cos(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekCOS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(cos(0.1));

  // COSH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=cosh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekCOSH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(cosh(0.1));

  // DEGREES
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=degrees(1.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNFunc(fekDEGREES, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(radToDeg(1.0));

  // EXP
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=exp(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekEXP, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(exp(0.1));

  // INT (positive argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=int(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekINT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(floor(0.1));

  // INT (negative argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=int(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekINT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(floor(-0.1));

  // LN - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ln(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekLN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(ln(0.1));

  // LN - error due to argument = 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ln(0.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.0,
    RPNFunc(fekLN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // LN - error due to argument < 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ln(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekLN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // LOG10 - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log10(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekLOG10, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(log10(0.1));

  // LOG10 - error due to argument = 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log10(0.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.0,
    RPNFunc(fekLOG10, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // LOG10 - error due to argument < 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log10(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekLOG10, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // LOG - valid result (2 arguments)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1, 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNNumber(2,
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(logn(2, 0.1));

  // LOG - valid result (2 arguments, base missing)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1, )');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNMissingArg(
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(logn(10, 0.1));

  // LOG - valid result (1 argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekLOG, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(logn(10, 0.1));

  // LOG - negative base
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1, -2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNNumber(-2,
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // LOG - negative value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(-0.1, 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNNumber(2.0,
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // PI
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=pi()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(pi);

  // RADIANS
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=radians(60)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(60,
    RPNFunc(fekRADIANS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(degtorad(60));

  // RAND
  // Test omitted because we cannot enforce getting the same random number back
  // when we are reading the file.

  // ROUND
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ROUND(pi(), 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI,
    RPNNumber(2,
    RPNFunc(fekROUND, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(RoundTo(pi, 2));

  // SIGN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sign(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekSIGN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(sign(-0.1));

  // SIN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sin(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekSIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(sin(0.1));

  // SINH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sinh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekSINH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(sinh(0.1));

  // SQRT - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sqrt(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekSQRT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(sqrt(0.1));

  // SQRT - error (negative argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sqrt(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekSQRT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errOverflow);

  // TAN - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=tan(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekTAN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(tan(0.1));

  // TAN - error (argument = pi/2)
  // This test is omitted because it is difficult to exactly hit pi/2.

  // TANH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=tanh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekTANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(tanh(0.1));


{------------------------------------------------------------------------------}
{                             Date/time functions                              }
{------------------------------------------------------------------------------}

  // DATE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=DATE(2014, 7, 1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(7,
    RPNNumber(1,
    RPNFunc(fekDATE, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(EncodeDate(2014,7,1));

  // DATEVALUE
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=DATEVALUE("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekDATEVALUE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(EncodeDate(2014,7,1));

  // DAY  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=DAY(DATE(2014,7,1))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(7,
    RPNNumber(1,
    RPNFunc(fekDATE,
    RPNFunc(fekDAY, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1);

  // DAY  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=DAY("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekDAY, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1);

  // HOUR  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=HOUR(TIME(9, 59, 20))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(9,
    RPNNumber(59,
    RPNNumber(20,
    RPNFunc(fekTIME,
    RPNFunc(fekHOUR, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(9);

  // HOUR  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=HOUR("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekHOUR, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(9);

  // MINUTE  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MINUTE(TIME(9, 59, 20))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(9,
    RPNNumber(59,
    RPNNumber(20,
    RPNFunc(fekTIME,
    RPNFunc(fekMINUTE, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(59);

  // MINUTE  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=MINUTE("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekMINUTE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(59);

  // MONTH  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MONTH(DATE(2014,7,1))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(7,
    RPNNumber(1,
    RPNFunc(fekDATE,
    RPNFunc(fekMONTH, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(7);

  // MONTH  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=MONTH("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekMONTH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(7);

  // NOW
  inc(Row);
  // Make sure that, when the file is read we still have the same second
  // Otherwise there would be a mismatch.
  repeat
    t := now();
    DecodeTime(t, hr, min, sec, msec);
  until msec < 500;
  MyWorksheet.WriteUTF8Text(Row, 0, '=NOW()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekNOW, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(t);

  // SECOND  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=SECOND(TIME(9, 59, 20))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(9,
    RPNNumber(59,
    RPNNumber(20,
    RPNFunc(fekTIME,
    RPNFunc(fekSECOND, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(20);

  // SECOND  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=SECOND("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekSECOND, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(20);

  // TIME
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TIME(9, 59, 20)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(9,
    RPNNumber(59,
    RPNNumber(20,
    RPNFunc(fekTIME, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(EncodeTime(9, 59, 20, 0));

  // TIMEVALUE
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));  // Localization!
  MyWorksheet.WriteUTF8Text(Row, 0, '=TIMEVALUE("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekTIMEVALUE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(EncodeTime(9, 59, 20, 0));

  // TODAY
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TODAY()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekTODAY, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(Date());

  // WEEKDAY  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=WEEKDAY(DATE(2014,7,1))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(7,
    RPNNumber(1,
    RPNFunc(fekDATE,
    RPNFunc(fekWEEKDAY, 1, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(DayOfWeek(EncodeDate(2014,7,1)));

  // WEEKDAY  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=WEEKDAY("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekWEEKDAY, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(DayOfWeek(EncodeDate(2014,7,1)));

  // YEAR  / argument number
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=YEAR(DATE(2014,7,1))');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(7,
    RPNNumber(1,
    RPNFunc(fekDATE,
    RPNFunc(fekYEAR, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(2014);

  // YEAR  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=YEAR("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekYEAR, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(2014);


{------------------------------------------------------------------------------}
{                          Statistical functions                               }
{------------------------------------------------------------------------------}

  // AVEDEV
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AVEDEV(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekAVEDEV, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(mean([0.0, 0.1, 0.2, 0.1, 0.2]));
                            // these values are the absolute deviations from mean (1.0)
  // AVERAGE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AVERAGE(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekAVERAGE, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(mean([1.0, 1.1, 1.2, 0.9, 0.8]));

  // COUNT (no missing values)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=COUNT(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
   RPNNumber(1.0,
   RPNNumber(1.1,
   RPNNumber(1.2,
   RPNNumber(0.9,
   RPNNumber(0.8,
   RPNFunc(fekCOUNT, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(5);

  // COUNT (with missing values)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=COUNT(1, , 1.2, , 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
   RPNNumber(1.0,
   RPNMissingArg(
   RPNNumber(1.2,
   RPNMissingArg(
   RPNNumber(0.8,
   RPNFunc(fekCOUNT, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(3);

   // MAX
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MAX(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekMAX, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(MaxValue([1.0, 1.1, 1.2, 0.9, 0.8]));

  // MAX (with missing values)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MAX(1, , , 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
   RPNNumber(1.0,
   RPNMissingArg(
   RPNMissingArg(
   RPNNumber(0.9,
   RPNNumber(0.8,
   RPNFunc(fekMAX, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(MaxValue([1.0, {1.1, 1.2,} 0.9, 0.8]));

   // MIN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MIN(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekMIN, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(MinValue([1.0, 1.1, 1.2, 0.9, 0.8]));

  // PRODUCT
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=PRODUCT(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekPRODUCT, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1.0*1.1*1.2*0.9*0.8);

  // STDEV
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=STDEV(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekSTDEV, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(StdDev([1.0, 1.1, 1.2, 0.9, 0.8]));

  // STDEVP
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=STDEVP(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekSTDEVP, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(PopnStdDev([1.0, 1.1, 1.2, 0.9, 0.8]));

  // SUM
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=SUM(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekSUM, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(Sum([1.0, 1.1, 1.2, 0.9, 0.8]));

  // SUMSQ
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=SUMSQ(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekSUMSQ, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(SumOfSquares([1.0, 1.1, 1.2, 0.9, 0.8]));

  // VAR
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=VAR(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekVAR, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(Variance([1.0, 1.1, 1.2, 0.9, 0.8]));

  // VARP
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=VARP(1, 1.1, 1.2, 0.9, 0.8)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.1,
    RPNNumber(1.2,
    RPNNumber(0.9,
    RPNNumber(0.8,
    RPNFunc(fekVARP, 5, nil))))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(PopnVariance([1.0, 1.1, 1.2, 0.9, 0.8]));


{------------------------------------------------------------------------------}
{                            Logical functions                                 }
{------------------------------------------------------------------------------}

  // AND of one values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AND(true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNFunc(fekAND, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true);

  // AND of two values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AND(true,false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNFunc(fekAND, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true and false);

  // AND of three values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AND(true,false,true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNBool(true,
    RPNFunc(fekAND, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true and false and true);

  // OR of one values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=OR(true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNFunc(fekOR, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true);

  // OR of two values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=OR(true,false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNFunc(fekOR, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true or false);

  // OR of three values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=OR(true,false,true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNBool(true,
    RPNFunc(fekOR, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true or false or true);

  // function =FALSE()
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=FALSE()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekFALSE, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(false);

  // function =TRUE()
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TRUE()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekTRUE, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true);

  // NOT
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=NOT(false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNFunc(fekNOT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(not false);

  // IF (2 parameters)/strings/case true
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNString('A',
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('A');

  // IF (2 parameters) /floats/case true
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1);

  // IF (2 parameters)/strings/case falsee
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(false,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNString('A',
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(false);

  // IF (2 parameters) /floats/case tfalse
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(false,1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNNumber(1.0,
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(false);

  // IF (3 parameters)/strings
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,"A","B")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNString('A',
    RPNString('B',
    RPNFunc(fekIF, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('A');

  // IF (3 parameters) /floats
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1,2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNNumber(2.0,
    RPNFunc(fekIF,3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1);

  // IF (3 parameters) /floats / mixed types, case true
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNString('A',
    RPNFunc(fekIF,3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(1);

  // IF (3 parameters) /floats / mixed types, case false
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(false,1,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNNumber(1.0,
    RPNString('A',
    RPNFunc(fekIF, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('A');


{------------------------------------------------------------------------------}
{                             String functions                                 }
{------------------------------------------------------------------------------}

  // CHAR
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=CHAR(72)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(72,
    RPNFunc(fekCHAR, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString(char(72));

   // CODE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=CODE("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekCODE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(ord('H'));

  // LEFT (2 parameters)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Hallo word", 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNNumber(2,
    RPNFunc(fekLEFT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('Ha');

  // LEFT (2 parameters, utf8)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Ändern", 3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Ändern',
    RPNNumber(3,
    RPNFunc(fekLEFT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('Änd');

  // LEFT (2 parameters, 1 of them missing)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Hallo word", )');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNMissingArg(
    RPNFunc(fekLEFT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('H');

  // LEFT (1 parameter)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekLEFT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('H');

  // Lower case
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LOWER("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekLOWER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString(LowerCase('Hallo world'));

  // Lower case / utf8
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LOWER("Viele Grüße")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Viele Grüße',
    RPNFunc(fekLOWER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString(UTF8LowerCase('Viele Grüße'));

  // MID
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MID("Hallo word", 3, 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNNumber(3,
    RPNNumber(2,
    RPNFunc(fekMID, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('ll');

  // REPLACE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=REPLACE("weather", 2, 2, "he")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('weather',
    RPNNumber(2,
    RPNNumber(2,
    RPNString('he',
    RPNFunc(fekREPLACE, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('whether');

  // REPLACE / utf8
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=REPLACE("würde", 2, 1, "u")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('würde',
    RPNNumber(2,
    RPNNumber(1,
    RPNString('u',
    RPNFunc(fekREPLACE, nil)))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('wurde');

// RIGHT (2 parameters)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=RIGHT("Hallo word", 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNNumber(2,
    RPNFunc(fekRIGHT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('ld');

  // RIGHT (2 parameters, one of them missing)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=RIGHT("Hallo word", )');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNMissingArg(
    RPNFunc(fekRIGHT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('d');

  // RIGHT (1 parameter)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=RIGHT("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekRIGHT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('d');

  // SUBSTITUTE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=SUBSTITUTE("lAzArus", "A", "a")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('lAzArus',
    RPNString('A',
    RPNString('a',
    RPNFunc(fekSUBSTITUTE, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString('lazarus');

  // Trim
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TRIM("   Hallo word   ")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('   Hallo world   ',
    RPNFunc(fekTRIM, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString(UTF8Trim('   Hallo world   '));

  // Upper case
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=UPPER("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekUPPER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString(UpperCase('Hallo world'));

  // Upper case / utf8
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=UPPER("Viele Grüße")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Viele Grüße',
    RPNFunc(fekUPPER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateString(UTF8UpperCase('Viele Grüße'));


{------------------------------------------------------------------------------}
{                           Information functions                              }
{------------------------------------------------------------------------------}

  // IsError
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISERROR(1/0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekDiv,
    RPNFunc(fekISERROR, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true);

  // IsError
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISERROR(0/1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0,
    RPNNumber(1,
    RPNFunc(fekDiv,
    RPNFunc(fekISERROR, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(false);

  // VALUE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=VALUE("100")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('100',
    RPNFunc(fekVALUE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumber(100);

{------------------------------------------------------------------------------}
{                                Error cases                                   }
{------------------------------------------------------------------------------}
{$IFDEF ENABLE_DEFECTIVE_FORMULAS }
  // Using less parameters than specified
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNFunc(fekIF,3, nil)))));            // <-- we have only 2 parameters!
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errArgError);

  // Using more parameters than specified
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNString('A',
    RPNFunc(fekIF,2, nil))))));            // <-- we have 3 parameters, not 2
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateError(errWrongType);
  { The first idea was that this should report an ArgError, but in fact it is
    a WrongType error because popping two values from the stack finds a number,
    but a bool is expected }
{$ENDIF}
