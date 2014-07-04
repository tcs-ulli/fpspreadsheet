{ include file for "formulatests.pas", containing the test cases for the
  calcrpnformula test. }

  // Setting up some test numbers

  // Test range  M1:N3
  MyWorksheet.WriteUTF8Text(0, 12, 'A');         // M1
  MyWorksheet.WriteUTF8Text(0, 13, 'A');         // N1
  MyWorksheet.WriteNumber  (1, 12,  1);          // M2
  MyWorksheet.WriteNumber  (1, 13,  2);          // N2

{------------------------------------------------------------------------------}
{                            Basic operations                                  }
{------------------------------------------------------------------------------}
  // Addition
  Row := 0;

  MyWorksheet.WriteUTF8Text(Row, 0, '=1+1');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNFunc(fekAdd, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1.0 + 1.0);                   // B1 = 2
  // don't refer to the cell contents here because they have not yet been calculated!
{!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!!  DO NOT CHANGE THIS FORMULA - ITS RESULT (-9) IS HARD-CODED IN OTHER TESTS !!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}

  // Subtraction
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=1-10');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(10,
    RPNFunc(fekSub, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1-10);                        // B2 = -9
  // don't refer to the cell contents here because they have not yet been calculated!
{!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!!  DO NOT CHANGE THIS FORMULA - ITS RESULT (-9) IS HARD-CODED IN OTHER TESTS !!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}

  // Add cell values - relative addresses
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=B1+B2');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('B2',
    RPNFunc(fekAdd, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(-7);

  // Add cell values - absolute addresses
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=$B$1+$B$2');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$B$2',
    RPNFunc(fekAdd, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(-7);
// don't refer to the cell contents here because they have not yet been calculated!

  // Add cell values - mixed absolute and relative addresses
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=B$1+$B2');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B$1',
    RPNCellValue('$B2',
    RPNFunc(fekAdd, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(-7);
  // don't refer to the cell contents here because they have not yet been calculated!

  // Multiplication
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10*-3');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(-3,
    RPNFunc(fekMul, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(10*(-3));

  // Multiplication  w/Parenthesis
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10*(-3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(-3,
    RPNParenthesis(
    RPNFunc(fekMul, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(10*(-3));

  // Multiplication of cell values
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=B1*B2');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('B2',
    RPNFunc(fekMul, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2*(-9));

  // Division
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10/200');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(200,
    RPNFunc(fekDiv, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(10/200);

  // Division: Error case - divide by zero
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10/0');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNNumber(0,
    RPNFunc(fekDiv, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errDivideByZero);

  // Division of cell values
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=B1/B2');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNcellValue('B2',
    RPNFunc(fekDiv, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2/(-9));

  // Percentage
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=10%');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(10,
    RPNFunc(fekPercent, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(10*0.01);

  // Percentage of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=B1%');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekPercent, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2*0.01);

  // Power
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=power(2.0,0.5)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2.0,
    RPNNumber(0.5,
    RPNFunc(fekPower, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(power(2, 0.5));

  // Power of cell values
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=power(B1,B2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('B2',
    RPNFunc(fekPower, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(power(2, -9));

{$IFDEF ENABLE_CALC_RPN_EXCEPTIONS}
  // Power: Error case  "power( (negative number), (fractional number) )"
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=power(-2.0, 0.5)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-2.0,
    RPNNumber(0.5,
    RPNFunc(fekPower, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);
{$ENDIF}

  // Unary minus
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=-(-1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-1,
    RPNParenthesis(
    RPNFunc(fekUMinus, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

   // Unary minus of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=-B1');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekUMinus, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(-(2));

  // Unary plus
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=+(-1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-1,
    RPNParenthesis(
    RPNFunc(fekUPlus, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(-1);

  // Unary plus  of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=+(B2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellvalue('B2',
    RPNFunc(fekUPlus, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(+(-9));

  // String result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '="Hallo"');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo', nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('Hallo');

  // String concatenation
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '="Hallo"&" world"');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString(' world',
    RPNFunc(fekConcat, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('Hallo' + ' world');

  // Equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true=false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true = false);

  // Equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"="world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg('Hallo' = 'world');

  // Equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1=1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(1=1);

  // Greater (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true>false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekGreater, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true > false);

  // Greater (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo">"world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekGreater, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg('Hallo' > 'world');

  // Greater (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1>1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekGreater, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(1>1);

  // Greater equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true>=false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekGreaterEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true >= false);

  // Greater equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo">="world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekGreaterEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg('Hallo' >= 'world');

  // Greater equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1>=1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekGreaterEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(1>=1);

  // Less (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true<false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekLess, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true < false);

  // Less (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"<"world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekLess, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg('Hallo' < 'world');

  // Less (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1<1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekLess, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(1<1);

  // Less equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true<=false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekLessEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true <= false);

  // Less equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"<="world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekLessEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg('Hallo' <= 'world');

  // Less equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1<=1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekLessEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(1<=1);

  // Not equal (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(true<>false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNParenthesis(
    RPNFunc(fekNotEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true <> false);

  // Not equal (strings)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=("Hallo"<>"world")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo',
    RPNString('world',
    RPNParenthesis(
    RPNFunc(fekNotEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg('Hallo' <> 'world');

  // Not equal (numbers)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=(1<>1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNNumber(1.0,
    RPNParenthesis(
    RPNFunc(fekNotEqual, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(1<>1);

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
  sollValues[Row] := CreateNumberArg(abs(-1));

  // ARCCOS - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acos(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekACOS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(arccos(0.1));

  // ACOS - error result (arccos is not defined outside the interval [-1..1]
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acos(-2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-2,
    RPNFunc(fekACOS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverFlow);

   // ARCCOSH - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acosh(1.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.1,
    RPNFunc(fekACOSH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(arccosh(1.1));

  // ACOSH - error result (arccos is not defined below 1
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=acosh(0.9)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.9,
    RPNFunc(fekACOSH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverFlow);

  // ASIN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=asin(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekASIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(arcsin(0.1));

  // ASIN - error result (arcsin is not defined outside the interval [-1..1]
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=asin(-2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-2,
    RPNFunc(fekASIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverFlow);

   // ARCSINH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=asinh(1.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.1,
    RPNFunc(fekASINH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(arcsinh(1.1));

  // ATAN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=atan(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekATAN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(arctan(0.1));

  // ATANH - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=atanh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekATANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(arctanh(0.1));

  // ATANH - error result (arctan is only defined within ]-1,1[
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=atanh(1.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNFunc(fekATANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverFlow);

  // COS
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=cos(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekCOS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(cos(0.1));

  // COSH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=cosh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekCOSH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(cosh(0.1));

  // DEGREES
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=degrees(1.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1.0,
    RPNFunc(fekDEGREES, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(radToDeg(1.0));

  // EXP
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=exp(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekEXP, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(exp(0.1));

  // INT (positive argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=int(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekINT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(floor(0.1));

  // INT (negative argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=int(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekINT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(floor(-0.1));

  // LN - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ln(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekLN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(ln(0.1));

  // LN - error due to argument = 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ln(0.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.0,
    RPNFunc(fekLN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LN - error due to argument < 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ln(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekLN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LOG10 - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log10(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekLOG10, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(log10(0.1));

  // LOG10 - error due to argument = 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log10(0.0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.0,
    RPNFunc(fekLOG10, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LOG10 - error due to argument < 0
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log10(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekLOG10, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LOG - valid result (2 arguments)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1, 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNNumber(2,
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(logn(2, 0.1));

  // LOG - valid result (2 arguments, base missing)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1, )');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNMissingArg(
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LOG - valid result (1 argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekLOG, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(logn(10, 0.1));

  // LOG - negative base
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(0.1, -2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNNumber(-2,
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LOG - negative value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(-0.1, 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNNumber(2.0,
    RPNFunc(fekLOG, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // LOG of cell
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=log(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekLOG, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(logn(10, 2));

  // PI
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=pi()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(pi);

  // RADIANS
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=radians(60)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(60,
    RPNFunc(fekRADIANS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(degtorad(60));

  // RADIANS of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=radians(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekRADIANS, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(degtorad(2));

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
  sollValues[Row] := CreateNumberArg(RoundTo(pi, 2));

  // SIGN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sign(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekSIGN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sign(-0.1));

  // SIGN of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sign(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekSIGN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sign(2));

  // SIN
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sin(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekSIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sin(0.1));

  // SIN   of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sin(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekSIN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sin(2));

  // SINH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sinh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekSINH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sinh(0.1));

  // SINH of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sinh(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekSINH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sinh(2));

  // SQRT - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sqrt(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekSQRT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sqrt(0.1));

  // SQRT - error (negative argument)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sqrt(-0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(-0.1,
    RPNFunc(fekSQRT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errOverflow);

  // SQRT of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=sqrt(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellvalue('B1',
    RPNFunc(fekSQRT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(sqrt(2));

  // TAN - valid result
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=tan(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekTAN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(tan(0.1));

  // TAN - error (argument = pi/2)
  // This test is omitted because it is difficult to exactly hit pi/2.

  // TAN of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=tan(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekTAN, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(tan(2));

  // TANH
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=tanh(0.1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0.1,
    RPNFunc(fekTANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(tanh(0.1));

  // TANH of cell value
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=tanh(B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekTANH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(tanh(2));


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
  sollValues[Row] := CreateNumberArg(EncodeDate(2014,7,1));

  // DATEVALUE
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=DATEVALUE("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekDATEVALUE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(EncodeDate(2014,7,1));

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
  sollValues[Row] := CreateNumberArg(1);

  // DAY  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=DAY("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekDAY, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

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
  sollValues[Row] := CreateNumberArg(9);

  // HOUR  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=HOUR("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekHOUR, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(9);

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
  sollValues[Row] := CreateNumberArg(59);

  // MINUTE  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=MINUTE("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekMINUTE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(59);

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
  sollValues[Row] := CreateNumberArg(7);

  // MONTH  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=MONTH("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekMONTH, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(7);

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
  sollValues[Row] := CreateNumberArg(t);

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
  sollValues[Row] := CreateNumberArg(20);

  // SECOND  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=SECOND("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekSECOND, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(20);

  // TIME
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TIME(9, 59, 20)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(9,
    RPNNumber(59,
    RPNNumber(20,
    RPNFunc(fekTIME, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(EncodeTime(9, 59, 20, 0));

  // TIMEVALUE
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));  // Localization!
  MyWorksheet.WriteUTF8Text(Row, 0, '=TIMEVALUE("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekTIMEVALUE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(EncodeTime(9, 59, 20, 0));

  // TODAY
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TODAY()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekTODAY, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(Date());

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
  sollValues[Row] := CreateNumberArg(DayOfWeek(EncodeDate(2014,7,1)));

  // WEEKDAY  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=WEEKDAY("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekWEEKDAY, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(DayOfWeek(EncodeDate(2014,7,1)));

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
  sollValues[Row] := CreateNumberArg(2014);

  // YEAR  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  MyWorksheet.WriteUTF8Text(Row, 0, '=YEAR("'+s+'")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(s,
    RPNFunc(fekYEAR, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2014);


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
  sollValues[Row] := CreateNumberArg(mean([0.0, 0.1, 0.2, 0.1, 0.2]));
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
  sollValues[Row] := CreateNumberArg(mean([1.0, 1.1, 1.2, 0.9, 0.8]));

  // AVERAGE of cell block
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AVERAGE(B1:B2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('B1:B2',
    RPNFunc(fekAVERAGE, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(mean([2.0, -9.0]));

  // COUNT
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
  sollValues[Row] := CreateNumberArg(5);

  // COUNT  of cell range (no empty cells)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=COUNT(B1:B2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellrange('B1:B2',
    RPNFunc(fekCOUNT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  // COUNT of cell range (no empty cells, but with non-numeric cells)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=COUNT(A1:B2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellrange('A1:B2',
    RPNFunc(fekCOUNT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  // COUNT  of cell range (with empty cells)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=COUNT(B1:C2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellrange('B1:C2',
    RPNFunc(fekCOUNT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  if AFormat <> sfExcel2 then begin
    // COUNTBLANK
    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=COUNTBLANK(A1:D2)');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('A1:D2',
      RPNFunc(fekCOUNTBLANK, nil))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(4);
  end;

  if AFormat <> sfExcel2 then begin
    // COUNTIF
    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=COUNTIF(M1:N3,1)');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNNumber(1,                    // "1" is in M2
      RPNFunc(fekCOUNTIF, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=COUNTIF(M1:N3,">=1")');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNString('>=1',                    // M2=1, N2=2 --> 2
      RPNFunc(fekCOUNTIF, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(2);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=COUNTIF(M1:N3,"<2")');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNString('<2',                    // M2=1, N2=2 --> 1
      RPNFunc(fekCOUNTIF, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=COUNTIF(M1:N3,"<>2")');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNString('<>2',                    // N2=2, 6 other cells --> 5
      RPNFunc(fekCOUNTIF, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(5);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=COUNTIF(M1:N3,M1)');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNCellValue('M1',                    // M1="A", N1="A" -> 2
      RPNFunc(fekCOUNTIF, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(2);
  end;

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
  sollValues[Row] := CreateNumberArg(MaxValue([1.0, 1.1, 1.2, 0.9, 0.8]));

  // MAX of cell range (no empty cells)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MAX(B1:B2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('B1:B2',
    RPNFunc(fekMAX, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(MaxValue([2.0, -9.0]));

  // MAX of cell range  (incl empty cells)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MAX(B1:C2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('B1:C2',
    RPNFunc(fekMAX, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(MaxValue([2.0, -9.0]));

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
  sollValues[Row] := CreateNumberArg(MinValue([1.0, 1.1, 1.2, 0.9, 0.8]));

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
  sollValues[Row] := CreateNumberArg(1.0*1.1*1.2*0.9*0.8);

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
  sollValues[Row] := CreateNumberArg(StdDev([1.0, 1.1, 1.2, 0.9, 0.8]));

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
  sollValues[Row] := CreateNumberArg(PopnStdDev([1.0, 1.1, 1.2, 0.9, 0.8]));

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
  sollValues[Row] := CreateNumberArg(Sum([1.0, 1.1, 1.2, 0.9, 0.8]));

  if AFormat <> sfExcel2 then begin
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
    sollValues[Row] := CreateNumberArg(SumOfSquares([1.0, 1.1, 1.2, 0.9, 0.8]));

    // SUMIF
    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=SUMIF(M1:N3,1)');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNNumber(1,                    // "1" is in M2
      RPNFunc(fekSUMIF, 2, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=SUMIF(M1:N3,">=1")');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNString('>=1',                    // M2=1, N2=2 --> 2
      RPNFunc(fekSUMIF, 2, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1+2);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=SUMIF(M1:N3,"<2")');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNString('<2',                    // M2=1, N2=2 --> 1
      RPNFunc(fekSUMIF, 2, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=SUMIF(M1:N3,"<>2")');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNString('<>2',                    // N2=2, 6 other cells --> 5
      RPNFunc(fekSUMIF, 2, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1);

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=SUMIF(M1:N3,M1)');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNCellValue('M1',                    // M1="A", N1="A" -> 2
      RPNFunc(fekSUMIF, 2, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(0);  // no numbers!

    inc(Row);
    MyWorksheet.WriteUTF8Text(Row, 0, '=SUMIF(M1:N3,M2)');
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('M1:N3',
      RPNCellValue('M2',                    // M2=1
      RPNFunc(fekSUMIF, 2, nil)))));
    SetLength(sollValues, Row+1);
    sollValues[Row] := CreateNumberArg(1);
  end;

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
  sollValues[Row] := CreateNumberArg(Variance([1.0, 1.1, 1.2, 0.9, 0.8]));

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
  sollValues[Row] := CreateNumberArg(PopnVariance([1.0, 1.1, 1.2, 0.9, 0.8]));


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
  sollValues[Row] := CreateBoolArg(true);

  // AND of two values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AND(true,false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNFunc(fekAND, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true and false);

  // AND of three values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AND(true,false,true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNBool(true,
    RPNFunc(fekAND, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true and false and true);

  // OR of one values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=OR(true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNFunc(fekOR, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true);

  // OR of two values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=OR(true,false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNFunc(fekOR, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true or false);

  // OR of three values (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=OR(true,false,true)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNBool(true,
    RPNFunc(fekOR, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true or false or true);

  // function =FALSE()
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=FALSE()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekFALSE, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(false);

  // function =TRUE()
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TRUE()');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekTRUE, nil)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true);

  // NOT
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=NOT(false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNFunc(fekNOT, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(not false);

  // IF (2 parameters)/strings/case true
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNString('A',
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('A');

  // IF (2 parameters) /floats/case true
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  // IF (2 parameters)/strings/case falsee
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(false,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNString('A',
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(false);

  // IF (2 parameters) /floats/case tfalse
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(false,1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNNumber(1.0,
    RPNFunc(fekIF, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(false);

  // IF (3 parameters)/strings
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,"A","B")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNString('A',
    RPNString('B',
    RPNFunc(fekIF, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('A');

  // IF (3 parameters) /floats
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1,2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNNumber(2.0,
    RPNFunc(fekIF,3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  // IF (3 parameters) /floats / mixed types, case true
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNString('A',
    RPNFunc(fekIF,3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  // IF (3 parameters) /floats / mixed types, case false
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(false,1,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(false,
    RPNNumber(1.0,
    RPNString('A',
    RPNFunc(fekIF, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('A');


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
  sollValues[Row] := CreateStringArg(char(72));

   // CODE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=CODE("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekCODE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(ord('H'));

  // LEFT (2 parameters)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Hallo word", 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNNumber(2,
    RPNFunc(fekLEFT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('Ha');

  // LEFT (2 parameters, utf8)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Ändern", 3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Ändern',
    RPNNumber(3,
    RPNFunc(fekLEFT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('Änd');

  // LEFT (2 parameters, 1 of them missing)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Hallo word", )');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNMissingArg(
    RPNFunc(fekLEFT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('H');

  // LEFT (1 parameter)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LEFT("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekLEFT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('H');

  // Lower case
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LOWER("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekLOWER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(LowerCase('Hallo world'));

  // Lower case / utf8
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=LOWER("Viele Grüße")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Viele Grüße',
    RPNFunc(fekLOWER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(UTF8LowerCase('Viele Grüße'));

  // MID
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=MID("Hallo word", 3, 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNNumber(3,
    RPNNumber(2,
    RPNFunc(fekMID, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('ll');

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
  sollValues[Row] := CreateStringArg('whether');

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
  sollValues[Row] := CreateStringArg('wurde');

// RIGHT (2 parameters)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=RIGHT("Hallo word", 2)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNNumber(2,
    RPNFunc(fekRIGHT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('ld');

  // RIGHT (2 parameters, one of them missing)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=RIGHT("Hallo word", )');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNMissingArg(
    RPNFunc(fekRIGHT, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('d');

  // RIGHT (1 parameter)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=RIGHT("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekRIGHT, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('d');

  // SUBSTITUTE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=SUBSTITUTE("lAzArus", "A", "a")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('lAzArus',
    RPNString('A',
    RPNString('a',
    RPNFunc(fekSUBSTITUTE, 3, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('lazarus');

  // Trim
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=TRIM("   Hallo word   ")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('   Hallo world   ',
    RPNFunc(fekTRIM, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(UTF8Trim('   Hallo world   '));

  // Upper case
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=UPPER("Hallo word")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hallo world',
    RPNFunc(fekUPPER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(UpperCase('Hallo world'));

  // Upper case / utf8
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=UPPER("Viele Grüße")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Viele Grüße',
    RPNFunc(fekUPPER, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(UTF8UpperCase('Viele Grüße'));


{------------------------------------------------------------------------------}
{                        Lookup / referece functions                           }
{------------------------------------------------------------------------------}

  // COLUMN
  MyWorksheet.WriteUTF8Text(Row, 0, '=COLUMN(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekCOLUMN, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, '=COLUMN(C1:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C1:D3',
    RPNFunc(fekCOLUMN, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(3);

  // COLUMNS
  MyWorksheet.WriteUTF8Text(Row, 0, '=COLUMNS(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekCOLUMNS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, '=COLUMNS(C1:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C1:D3',
    RPNFunc(fekCOLUMNS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  // ROW
  MyWorksheet.WriteUTF8Text(Row, 0, '=ROW(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekROW, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, '=ROW(C2:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C2:D3',
    RPNFunc(fekROW, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  // ROWS
  MyWorksheet.WriteUTF8Text(Row, 0, '=ROWS(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekROWS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, '=ROWS(C2:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C2:D3',
    RPNFunc(fekROWS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);


{------------------------------------------------------------------------------}
{                           Information functions                              }
{------------------------------------------------------------------------------}

  // CELL
  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("address", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('address',
    RPNCellRef('A1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('$A$1');

  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("col", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('col',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);      // Excel starts counting at 1

  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("format", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('format',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('G');

  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("prefix", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('prefix',
    RPNCellRef('A1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('''');

  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("row", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('row',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);      // Excel starts counting at 1

  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("type", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('type',
    RPNCellRef('A1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('l');

  MyWorksheet.WriteUTF8Text(Row, 0, '=CELL("type", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('type',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('v');

  // INFO
  MyWorksheet.WriteUTF8Text(Row, 0, '=INFO("directory", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('directory',
    RPNFunc(fekINFO, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(ExtractFilePath(TempFile));

  MyWorksheet.WriteUTF8Text(Row, 0, '=INFO("numfile")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('numfile',
    RPNFunc(fekINFO, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(MyWorkbook.GetWorksheetCount);

  // IsBlank
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISBLANK(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekISBLANK, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(false);  // cell contains text --> not blank

  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISBLANK(G1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('G1',
    RPNFunc(fekISBLANK, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true);  // the cell does not exist --> blank

  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISBLANK(H1)');
  MyWorksheet.WriteBlank(0, 7);          // A11 is an empty cell
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('H1',
    RPNFunc(fekISBLANK, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true);  // the cell exists, but it is empty

  // IsError
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISERROR(1/0)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekDiv,
    RPNFunc(fekISERROR, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true);

  // IsError
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISERROR(0/1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0,
    RPNNumber(1,
    RPNFunc(fekDiv,
    RPNFunc(fekISERROR, nil))))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(false);

  // IsRef
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISREF(1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNFunc(fekISREF, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(false);

  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISREF(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekISREF, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBoolArg(true);

  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=ISREF(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('A1',                 // we use a cell value here !
    RPNFunc(fekISREF, nil))));
  SetLength(sollValues, Row+1);
  // The correct result would be "false" because a cell value is not the same as
  // a cell reference. But Excel seems to ignore this difference here and
  // accepts only a "true".
  sollValues[Row] := CreateBoolArg(true);

  // VALUE
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=VALUE("100")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('100',
    RPNFunc(fekVALUE, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(100);

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
  sollValues[Row] := CreateErrorArg(errArgError);

  // Using more parameters than specified
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=IF(true,1,"A")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNString('A',
    RPNFunc(fekIF,2, nil))))));         // <-- but we pushed 3 parameters, not 2
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errWrongType);
  { The first idea was that this should report an ArgError, but in fact it is
    a WrongType error because popping two values from the stack finds a number,
    but a bool is expected }
{$ENDIF}

