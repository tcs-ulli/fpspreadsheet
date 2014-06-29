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



                (*  variable param count !!!!!!!!!!!!!!!!
  // AND (bool)
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, '=AND(true,false)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNBool(false,
    RPNFunc(fekAND, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateBool(true and false);
                  *)


