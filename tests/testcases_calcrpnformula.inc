{ include file for "formulatests.pas", containing the test cases for the
  calcrpnformula test. }

  // Setting up some test numbers

  // Test range  M1:N3
  MyWorksheet.WriteUTF8Text(0, 12, 'A');         // M1
  MyWorksheet.WriteUTF8Text(0, 13, 'A');         // N1
  MyWorksheet.WriteNumber  (1, 12,  1);          // M2
  MyWorksheet.WriteNumber  (1, 13,  2);          // N2

  // Create an error in H2
  MyWorksheet.WriteFormula (1, 7, '1/0');
{------------------------------------------------------------------------------}
{                            Basic operations                                  }
{------------------------------------------------------------------------------}
  // Addition
  Row := 0;

  formula := '1+1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.0,
      RPNFunc(fekAdd, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(1 + 1);                           // B1 = 2
  cellB1 := 1 + 1;
  // don't refer to the cell contents here because they have not yet been calculated!
{!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!!  DO NOT CHANGE THIS FORMULA - ITS RESULT (-9) IS HARD-CODED IN OTHER TESTS !!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}

  // Subtraction
  inc(Row);
  formula := '1-10';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(10.0,
      RPNFunc(fekSub, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, -9);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(1 - 10);      // B2 = -9
  cellB2 := 1 - 10;
  // don't refer to the cell contents here because they have not yet been calculated!
{!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!!  DO NOT CHANGE THIS FORMULA - ITS RESULT (-9) IS HARD-CODED IN OTHER TESTS !!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}

  // Add cell values - relative addresses
  inc(Row);
  formula := 'B1+B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNCellValue('B2',
      RPNFunc(fekAdd, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 + cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 + cellB2);

  // Add cell values - absolute addresses
  inc(Row);
  formula := '$B$1+$B$2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('$B$1',
      RPNCellValue('$B$2',
      RPNFunc(fekAdd, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 + cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 + cellB2);
// don't refer to the cell contents here because they have not yet been calculated!

  // Add cell values - mixed absolute and relative addresses
  inc(Row);
  formula := 'B$1+$B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B$1',
      RPNCellValue('$B2',
      RPNFunc(fekAdd, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 + cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 + cellB2);
  // don't refer to the cell contents here because they have not yet been calculated!

  // Subtract cell values - relative addresses
  inc(Row);
  formula := 'B1-B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNCellValue('B2',
      RPNFunc(fekSub, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 - cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 - cellB2);

  // Subtract cell values - absolute addresses
  inc(Row);
  formula := '$B$1-$B$2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('$B$1',
      RPNCellValue('$B$2',
      RPNFunc(fekSub, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 - cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 - cellB2);

  // Subtract cell values - mixed absolute and relative addresses
  inc(Row);
  formula := 'B$1-$B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B$1',
      RPNCellValue('$B2',
      RPNFunc(fekSub, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 - cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 - cellB2);
  // don't refer to the cell contents here because they have not yet been calculated!

  // Multiplication w/o Parenthesis
  inc(Row);
  formula := '10*-3';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(10.0,
      RPNNumber(-3.0,
      RPNFunc(fekMul, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 10*(-3));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(10*(-3));

  // Multiplication  w/Parenthesis
  inc(Row);
  formula := '10*(-3)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(10.0,
      RPNNumber(-3.0,
      RPNParenthesis(
      RPNFunc(fekMul, nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 10*(-3));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(10*(-3));

  // Multiplication of cell values
  inc(Row);
  formula := 'B1*B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNCellValue('B2',
      RPNFunc(fekMul, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 * cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 * cellB2);

  // Division
  inc(Row);
  formula := '10/200';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(10.0,
      RPNNumber(200.0,
      RPNFunc(fekDiv, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 10/200);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(10/200);

  // Division: Error case - divide by zero
  inc(Row);
  formula := '10/0';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(10.0,
      RPNNumber(0,
      RPNFunc(fekDiv, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteErrorValue(Row, 2, errDivideByZero);
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errDivideByZero);

  // Division of cell values
  inc(Row);
  formula := 'B1/B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNcellValue('B2',
      RPNFunc(fekDiv, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, cellB1 / cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1 / cellB2);

  // Percentage
  inc(Row);
  formula := '10%';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(10.0,
      RPNFunc(fekPercent, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, 10*0.01);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(10*0.01);

  // Percentage of cell value
  inc(Row);
  formula := 'B1%';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc(fekPercent, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, cellB1*0.01);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cellB1*0.01);

  // Exotic test case from OpenOffice test file
  inc(Row);
  formula := '200%%';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(200,
      RPNFunc(fekPercent,
      RPNFunc(fekPercent, nil)))))
  else
    MyWorksheet.WriteFormula(row, 1, formula);
  SetLength(sollvalues, Row+1);
  sollValues[Row] := FloatResult(0.02);

  // Power symbol
  inc(Row);
  formula := '2^0.5';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2.0,
      RPNNumber(0.5,
      RPNFunc(fekPower, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, power(2, 0.5));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(power(2, 0.5));

  // Power symbol - precedence of operations
  inc(Row);
  formula := '2^0.5*4';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2.0,
      RPNNumber(0.5,
      RPNFunc(fekPower,
      RPNNumber(4,
      RPNFunc(fekMul, nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, power(2, 0.5)*4);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(power(2, 0.5)*4);

  // Power symbol - precedence of operators
  inc(Row);
  formula := '4*2^0.5';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(4.0,
      RPNNumber(2.0,
      RPNNumber(0.5,
      RPNFunc(fekPower,
      RPNFunc(fekMul, nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 4 * power(2, 0.5));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(4 * power(2, 0.5));

  // Exotic sample from OpenOffice test files
  inc(Row);
  formula := '50%^200%';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(50.0,
      RPNFunc(fekPercent,
      RPNNumber(200.0,
      RPNFunc(fekPercent,
      RPNFunc(fekPower, nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, power(0.5, 2));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(power(0.5, 2));

  // Power function
  if AFormat <> sfExcel2 then begin
    // Power of constant
    inc(Row);
    formula := 'POWER(2,0.5)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(2.0,
        RPNNumber(0.5,
        RPNFunc('POWER', nil)))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, power(2, 0.5));
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(power(2, 0.5));

    // Power of cell values
    inc(Row);
    formula := 'POWER(B1,B2)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellValue('B1',
        RPNCellValue('B2',
        RPNFunc('POWER', nil)))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, power(cellB1, cellB2));
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(power(cellB1, cellB2));
  end;

{$IFDEF ENABLE_CALC_RPN_EXCEPTIONS}
  // Power: Error case  "power( (negative number), (fractional number) )"
  inc(Row);
  formula := 'POWER(-2.0, 0.5)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-2.0,
      RPNNumber(0.5,
      RPNFunc('POWER', nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  with sollValues[Row] do begin
    ResultType := rtError;
    ResError := errOverflow;
  end;
{$ENDIF}

  // Unary minus
  inc(Row);
  formula := '-(-1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNformula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-1.0,
      RPNParenthesis(
      RPNFunc(fekUMinus, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, -(-1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(-(-1));

   // Unary minus of cell value
  inc(Row);
  formula := '-B1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc(fekUMinus, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, -cellB1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(-cellB1);

  // Unary plus
  inc(Row);
  formula := '+(-1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-1.0,
      RPNParenthesis(
      RPNFunc(fekUPlus, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, -1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(-1);

  // Unary plus of cell value
  inc(Row);
  formula := '+(B2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellvalue('B2',
      RPNParenthesis((
      RPNFunc(fekUPlus, nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, +cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(+cellB2);

  // Exotic case from OpenOffice test files
  inc(Row);
  formula := '+-B2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellvalue('B2',
      RPNFunc(fekUMinus,
      RPNFunc(fekUPlus, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, -cellB2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(-cellB2);

  // String result
  inc(Row);
  formula := '"Hallo"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo', nil)))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'Hallo');
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('Hallo');

  // String concatenation
  inc(Row);
  formula := '"Hallo"&" world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString(' world',
      RPNFunc(fekConcat, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'Hallo' + ' world');
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('Hallo' + ' world');

  // Equal (bool)
  inc(Row);
  formula := 'TRUE=FALSE';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc(fekEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true = false);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Equal (strings)
  inc(Row);
  formula := '"Hallo"="world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString('world',
      RPNFunc(fekEqual, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult('Hallo' = 'world');
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Equal (numbers)
  inc(Row);
  formula := '1=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(1,
      RPNFunc(fekEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(1 = 1);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Equal (cell)
  inc(Row);
  formula := 'M1=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M1',
      RPNNumber(1,
      RPNFunc(fekEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // M1 is "A"
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  inc(Row);
  formula := 'M2=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNNumber(1.0,
      RPNFunc(fekEqual, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // M2 is 1
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  inc(Row);
  formula := 'M2=N2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNCellValue('N2',
      RPNFunc(fekEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // M2 is 1, N2 is 2
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater (bool)
  inc(Row);
  formula := 'TRUE>FALSE';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc(fekGreater, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true > false);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater (strings)
  inc(Row);
  formula := '"Hallo">"world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString('world',
      RPNFunc(fekGreater, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult('Hallo' > 'world');
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater (numbers)
  inc(Row);
  formula := '1>1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(1,
      RPNFunc(fekGreater, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(1 > 1);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater (cell)
  inc(Row);
  formula := 'M2>1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNNumber(1,
      RPNFunc(fekGreater, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  Myworksheet.WriteBoolvalue(Row, 2, SollValues[Row].ResBoolean);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  inc(Row);
  formula := 'M2>N2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNCellValue('N2',
      RPNFunc(fekGreater, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // M2 is 1, N2 is 2
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater equal (bool)
  inc(Row);
  formula := 'TRUE>=FALSE';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc(fekGreaterEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true >= false);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater equal (strings)
  inc(Row);
  formula := '"Hallo">="world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString('world',
      RPNFunc(fekGreaterEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult('Hallo' >= 'world');
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater equal (numbers)
  inc(Row);
  formula := '1>=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(1,
      RPNFunc(fekGreaterEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(1 >= 1);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Greater equal(cell)
  inc(Row);
  formula := 'M2>=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNNumber(1,
      RPNFunc(fekGreaterEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // M2 is 1
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  inc(Row);
  formula := 'M2>=N2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNCellValue('N2',
      RPNFunc(fekGreaterEqual, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // M2 is 1, N2 is 2
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less (bool)
  inc(Row);
  formula := 'TRUE<FALSE';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc(fekLess, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true < false);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less (strings)
  inc(Row);
  formula := '"Hallo"<"world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString('world',
      RPNFunc(fekLess, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult('Hallo' < 'world');
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less (numbers)
  inc(Row);
  formula := '1<1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(1,
      RPNFunc(fekLess, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(1 < 1);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less (cell)
  inc(Row);
  formula := 'M2<1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNNumber(1,
      RPNFunc(fekLess, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // M2 is 1
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  inc(Row);
  formula := 'M2<N2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNCellValue('N2',
      RPNFunc(fekLess, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // M2 is 1, N2 is 2
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less equal (bool)
  inc(Row);
  formula := 'TRUE<=FALSE';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc(fekLessEqual, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true <= false);
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less equal (strings)
  inc(Row);
  formula := '"Hallo"<="world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString('world',
      RPNFunc(fekLessEqual, nil)))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult('Hallo' <= 'world');
  Myworksheet.WriteUTF8Text(Row, 2, BoolToStr(SollValues[Row].ResBoolean, 'TRUE','FALSE'));

  // Less equal (numbers)
  inc(Row);
  formula := '1<=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(1,
      RPNFunc(fekLessEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(1<=1);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Less equal (cell)
  inc(Row);
  formula := 'M2<=1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNNumber(1,
      RPNFunc(fekLessEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // M2 is 1
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'M2<=N2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNCellValue('N2',
      RPNFunc(fekLessEqual, nil)))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // M2 is 1, N2 is 2
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Not equal (bool)
  inc(Row);
  formula := 'TRUE<>FALSE';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc(fekNotEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true <> false);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Not equal (strings)
  inc(Row);
  formula := '"Hallo"<>"world"';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo',
      RPNString('world',
      RPNFunc(fekNotEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult('Hallo' <> 'world');
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Not equal (numbers)
  inc(Row);
  formula := '1<>1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(1,
      RPNFunc(fekNotEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(1 <> 1);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // Not equal (cell)
  inc(Row);
  formula := 'M2<>1';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNNumber(1,
      RPNFunc(fekNotEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // M2 is 1
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'M2<>N2';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('M2',
      RPNCellValue('N2',
      RPNFunc(fekNotEqual, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // M2 is 1, N2 is 2
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');


{------------------------------------------------------------------------------}
{                                     Math                                     }
{------------------------------------------------------------------------------}
// ABS
  inc(Row);
  formula := 'ABS(-1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-1,
      RPNFunc('ABS', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(abs(-1));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // ARCCOS - valid result
  inc(Row);
  number := 0.1;
  formula := 'ACOS(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('ACOS', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(arccos(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);
  // Don't explicitly use the constant here because this will be "extended"
  // which is slightly different from the double value used in the formula

  // ACOS - error result (arccos is not defined outside the interval [-1..1]
  inc(Row);
  formula := 'ACOS(-2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-2,
      RPNFunc('ACOS', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteErrorValue(Row, 2, errOverflow);;
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverFlow);

   // ARCCOSH - valid result
  inc(Row);
  number := 1.1;
  formula := 'ACOSH(1.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(number,
      RPNFunc('ACOSH', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(arccosh(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // ACOSH - error result (arccos is not defined below 1
  inc(Row);
  formula := 'ACOSH(0.9)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.9,
      RPNFunc('ACOSH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteErrorValue(Row, 2, errOverflow);
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverFlow);

  // ASIN
  inc(Row);
  number := 0.1;
  formula := 'ASIN(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('ASIN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(arcsin(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // ASIN - error result (arcsin is not defined outside the interval [-1..1]
  inc(Row);
  formula := 'ASIN(-2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-2,
      RPNFunc('ASIN', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteErrorValue(Row, 2, errOverflow);
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverFlow);

   // ARCSINH
  inc(Row);
  number := 1.1;
  formula := 'ASINH(1.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.1,
      RPNFunc('ASINH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(arcsinh(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // ATAN
  inc(Row);
  number := 0.1;
  formula := 'ATAN(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('ATAN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(arctan(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // ATANH - valid result
  inc(Row);
  number := 0.1;
  formula := 'ATANH(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('ATANH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(arctanh(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // ATANH - error result (arctan is only defined within ]-1,1[
  inc(Row);
  formula := 'ATANH(1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNFunc('ATANH', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteErrorValue(Row, 2, errOverFlow);
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverFlow);

  // CEILING
  if AFormat <> sfExcel2 then begin
    inc(Row);
    number := 0.1;
    formula := 'CEILING(5.02,0.25)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(5.02,
        RPNNumber(0.25,
        RPNFunc('CEILING', nil)))))
    else
      myWorksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(ceil(5.02/0.25)*0.25);
    MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);
  end;

  // COS
  inc(Row);
  number := 0.1;
  formula := 'COS(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('COS', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cos(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // COSH
  inc(Row);
  number := 0.1;
  formula := 'COSH(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('COSH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(cosh(number));
  MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);

  // DEGREES
  if AFormat <> sfExcel2 then
  begin
    inc(Row);
    formula := 'DEGREES(1)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(1,
        RPNFunc('DEGREES', nil))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(radToDeg(1.0));
    myWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);
  end;

  // EVEN
  if AFormat <> sfExcel2 then
  begin
    inc(Row);
    formula := 'EVEN(11.2)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(11.2,
        RPNFunc('EVEN', nil))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(12);
    myWorksheet.WriteNumber(Row, 2, 12);
  end;

  // EXP
  inc(Row);
  number := 0.1;
  formula := 'EXP(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('EXP', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, exp(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(exp(number));

  // FACT
  inc(Row);
  formula := 'FACT(5)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNInteger(5,
      RPNFunc('FACT', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, 1*2*3*4*5);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(1*2*3*4*5);

  // FLOOR
  if AFormat <> sfExcel2 then begin
    inc(Row);
    number := 0.1;
    formula := 'FLOOR(5.02,0.25)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(5.02,
        RPNNumber(0.25,
        RPNFunc('FLOOR', nil)))))
    else
      myWorksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(floor(5.02/0.25)*0.25);
    MyWorksheet.WriteNumber(Row, 2, sollValues[Row].ResFloat);
  end;

  // INT (positive argument)
  inc(Row);
  number := 0.1;
  formula := 'INT(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('INT', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, floor(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(floor(number));

  // INT (negative argument)
  inc(Row);
  formula := 'INT(-0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-0.1,
      RPNFunc('INT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, floor(-0.1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(floor(-0.1));

  // LN - valid result
  inc(Row);
  number := 0.1;
  formula := 'LN(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('LN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, ln(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(ln(number));

  // LN - error due to argument = 0
  inc(Row);
  formula := 'LN(0)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.0,
      RPNFunc('LN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LN - error due to argument < 0
  inc(Row);
  formula := 'LN(-0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-0.1,
      RPNFunc('LN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LOG10 - valid result
  inc(Row);
  number := 0.1;
  formula := 'LOG10(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('LOG10', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, log10(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(log10(number));

  // LOG10 - error due to argument = 0
  inc(Row);
  formula := 'LOG10(0)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.0,
      RPNFunc('LOG10', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LOG10 - error due to argument < 0
  inc(Row);
  formula := 'LOG10(-0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-0.1,
      RPNFunc('LOG10', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverFlow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LOG - valid result (2 arguments)
  inc(Row);
  number := 0.1;
  formula := 'LOG(0.1,2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNNumber(2,
      RPNFunc('LOG', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, logn(2, number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(logn(2, number));

  // LOG - error result (2 arguments, base missing)
  inc(Row);
  formula := 'LOG(0.1,)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNMissingArg(
      RPNFunc('LOG', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LOG - valid result (1 argument)
  inc(Row);
  number := 0.1;
  formula := 'LOG(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('LOG', 1, nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, logn(10, number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(logn(10, number));

  // LOG - negative base
  inc(Row);
  formula := 'LOG(0.1,-2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNNumber(-2,
      RPNFunc('LOG', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LOG - negative value
  inc(Row);
  formula := 'LOG(-0.1,2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-0.1,
      RPNNumber(2.0,
      RPNFunc('LOG', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // LOG of cell
  inc(Row);
  formula := 'LOG(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc('LOG', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, logn(10, 2));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(logn(10, 2));

  // MOD
  inc(Row);
  formula := 'MOD(12,5)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNInteger(12,
      RPNInteger(5,
      RPNFunc('MOD', nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 12 mod 5);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(12 mod 5);

  // ODD
  if AFormat <> sfExcel2 then
  begin
    inc(Row);
    formula := 'ODD(11.2)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(11.2,
        RPNFunc('ODD', nil))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(13);
    myWorksheet.WriteNumber(Row, 2, 13);
  end;

  // PI
  inc(Row);
  formula := 'PI()';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc('PI', nil)))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, pi);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(pi);

  if AFormat <> sfExcel2 then
  begin
    // RADIANS
    inc(Row);
    formula := 'RADIANS(60)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(60,
        RPNFunc('RADIANS', nil))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    MyWorksheet.WriteNumber(Row, 2, DegToRad(60));
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(degtorad(60));

    // RADIANS of cell value
    inc(Row);
    formula := 'RADIANS(B1)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellValue('B1',
        RPNFunc('RADIANS', nil))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, DegToRad(cellB1));
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(degtorad(cellB1));
  end;

  // RAND
  // Test omitted because we cannot enforce getting the same random number back
  // when we are reading the file.

  // ROUND
  inc(Row);
  formula := 'ROUND(PI(),2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc('PI',
      RPNNumber(2,
      RPNFunc('ROUND', nil)))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, RoundTo(pi, 2));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(RoundTo(pi, 2));

  // SIGN
  inc(Row);
  number := -0.1;
  formula := 'SIGN(-0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-0.1,
      RPNFunc('SIGN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, sign(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sign(number));

  // SIGN of cell value
  inc(Row);
  formula := 'SIGN(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc('SIGN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, sign(cellB1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sign(cellB1));

  // SIN
  inc(Row);
  number := 0.1;
  formula := 'SIN(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('SIN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, sin(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sin(number));

  // SIN of cell value
  inc(Row);
  formula := 'SIN(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc('SIN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, sin(cellB1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sin(cellB1));

  // SINH
  inc(Row);
  number := 0.1;
  formula := 'SINH(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('SINH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, sinh(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sinh(number));

  // SINH of cell value
  inc(Row);
  formula := 'SINH(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc('SINH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, sinh(cellB1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sinh(cellB1));

  // SQRT - valid result
  inc(Row);
  number := 0.1;
  formula := 'SQRT(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('SQRT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, sqrt(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sqrt(number));

  // SQRT - error (negative argument)
  inc(Row);
  formula := 'SQRT(-0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(-0.1,
      RPNFunc('SQRT', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, GetErrorValueStr(errOverflow));
  SetLength(sollValues, Row+1);
  if AFormat = sfOpenDocument then
    sollValues[Row] := FloatResult(0)
  else
    sollValues[Row] := ErrorResult(errOverflow);

  // SQRT of cell value
  inc(Row);
  formula := 'SQRT(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellvalue('B1',
      RPNFunc('SQRT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, sqrt(cellB1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(sqrt(cellB1));

  // TAN - valid result
  inc(Row);
  number := 0.1;
  formula := 'TAN(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('TAN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, tan(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(tan(number));

  // TAN - error (argument = pi/2)
  // This test is omitted because it is difficult to exactly hit pi/2.

  // TAN of cell value
  inc(Row);
  formula := 'TAN(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc('TAN', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, tan(cellB1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(tan(cellB1));

  // TANH
  inc(Row);
  number := 0.1;
  formula := 'TANH(0.1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.1,
      RPNFunc('TANH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, tanh(number));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(tanh(number));

  // TANH of cell value
  inc(Row);
  formula := 'TANH(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('B1',
      RPNFunc('TANH', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, tanh(cellB1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(tanh(cellB1));


{------------------------------------------------------------------------------}
{                             Date/time functions                              }
{------------------------------------------------------------------------------}

  // DATE
  inc(Row);
  formula := 'DATE(2014,7,1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2014,
      RPNNumber(7,
      RPNNumber(1,
      RPNFunc('DATE', nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteDateTime(Row, 2, EncodeDate(2014,7,1));
  SetLength(sollValues, Row+1);
  sollValues[Row] := DateTimeResult(EncodeDate(2014,7,1));

  // DATEDIF
  if AFormat <> sfExcel2 then begin
    inc(Row);
    formula := 'DATEDIF(DATE(2014,7,1),DATE(2014,1,1),"D")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(2014,
        RPNNumber(7,
        RPNNumber(1,
        RPNFunc('DATE',
        RPNNumber(2014,
        RPNNumber(1,
        RPNNumber(1,
        RPNFunc('DATE',
        RPNString('D',
        RPNFUNC('DATEDIF',
        nil))))))))))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    MyWorksheet.WriteNumber(Row, 2, EncodeDate(2014,7,1)-EncodeDate(2014,1,1));
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(round(EncodeDate(2014,7,1)-EncodeDate(2014,1,1)));
  end;

  // DATEVALUE
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  formula := 'DATEVALUE("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('DATEVALUE', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteDateTimeFormat(Row, 1, nfShortDate);
  MyWorksheet.WriteDateTime(Row, 2, StrToDate(s), nfShortDate);
  SetLength(sollValues, Row+1);
  sollValues[Row] := DateTimeResult(EncodeDate(2014,7,1));

  // DAY  / argument number
  inc(Row);
  formula := 'DAY(DATE(2014,7,1))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2014,
      RPNNumber(7,
      RPNNumber(1,
      RPNFunc('DATE',
      RPNFunc('DAY', nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(1);

  // DAY  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  formula := 'DAY("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('DAY', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(1);

  // HOUR  / argument number
  inc(Row);
  formula := 'HOUR(TIME(9,59,20))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(9,
      RPNNumber(59,
      RPNNumber(20,
      RPNFunc('TIME',
      RPNFunc('HOUR', nil)))))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 9);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(9);

  // HOUR  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  formula := 'HOUR("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('HOUR', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, 9);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(9);

  // MINUTE  / argument number
  inc(Row);
  formula := 'MINUTE(TIME(9,59,20))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(9,
      RPNNumber(59,
      RPNNumber(20,
      RPNFunc('TIME',
      RPNFunc('MINUTE', nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, 59);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(59);

  // MINUTE  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  formula := 'MINUTE("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('MINUTE', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 59);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(59);

  // MONTH  / argument number
  inc(Row);
  formula := 'MONTH(DATE(2014,7,1))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2014,
      RPNNumber(7,
      RPNNumber(1,
      RPNFunc('DATE',
      RPNFunc('MONTH', nil)))))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 7);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(7);

  // MONTH  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  formula := 'MONTH("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('MONTH', nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 7);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(7);

  // NOW
  inc(Row);
  formula := 'NOW()';
  // Make sure that, when the file is read we still have the same second
  // Otherwise there would be a mismatch.
  repeat
    number := now();
    DecodeTime(number, hr, min, sec, msec);
  until msec < 500;
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc('NOW', nil)))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteDateTimeFormat(Row, 1, nfShortDateTime);
  MyWorksheet.WriteDateTime(Row, 2, number, nfShortDateTime);
  SetLength(sollValues, Row+1);
  sollValues[Row] := DateTimeResult(number);

  // SECOND  / argument number
  inc(Row);
  formula := 'SECOND(TIME(9,59,20))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(9,
      RPNNumber(59,
      RPNNumber(20,
      RPNFunc('TIME',
      RPNFunc('SECOND', nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 1, 20);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(20);

  // SECOND  / argument string
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));   // Localization of the test
  formula := 'SECOND("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('SECOND', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 1, 20);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(20);

  // TIME
  inc(Row);
  formula := 'TIME(9,59,20)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(9,
      RPNNumber(59,
      RPNNumber(20,
      RPNFunc('TIME', nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteDateTimeFormat(Row, 1, nfLongTime);
  MyWorksheet.WriteDateTime(Row, 2, EncodeTime(9, 50, 20, 0), nfLongTime);
  SetLength(sollValues, Row+1);
  sollValues[Row] := DateTimeResult(EncodeTime(9, 59, 20, 0));

  // TIMEVALUE
  inc(Row);
  s := TimeToStr(EncodeTime(9, 59, 20, 0));  // Localization!
  formula := 'TIMEVALUE("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('TIMEVALUE', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteDateTimeFormat(Row, 1, nfLongTime);
  MyWorksheet.WriteDateTime(Row, 2, EncodeTime(9, 59, 20, 0), nfLongTime);
  SetLength(sollValues, Row+1);
  sollValues[Row] := DateTimeResult(EncodeTime(9, 59, 20, 0));

  // TODAY
  inc(Row);
  formula := 'TODAY()';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc('TODAY', nil)))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteDateTimeFormat(Row, 1, nfShortDate);
  Myworksheet.WriteDateTime(Row, 2, Date(), nfShortDate);
  SetLength(sollValues, Row+1);
  sollValues[Row] := DateTimeResult(Date());

  // WEEKDAY  / argument number
  inc(Row);
  formula := 'WEEKDAY(DATE(2014,7,1))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2014,
      RPNNumber(7,
      RPNNumber(1,
      RPNFunc('DATE',
      RPNFunc('WEEKDAY', 1, nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, DayOfWeek(EncodeDate(2014,7,1)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(DayOfWeek(EncodeDate(2014,7,1)));

  // WEEKDAY  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  formula := 'WEEKDAY("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('WEEKDAY', 1, nil))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, DayOfWeek(EncodeDate(2014,7,1)));
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(DayOfWeek(EncodeDate(2014,7,1)));

  // YEAR  / argument number
  inc(Row);
  formula := 'YEAR(DATE(2014,7,1))';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(2014,
      RPNNumber(7,
      RPNNumber(1,
      RPNFunc('DATE',
      RPNFunc('YEAR', nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 2014);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(2014);

  // YEAR  / argument string
  inc(Row);
  s := DateToStr(EncodeDate(2014, 7, 1));   // Localization of the test
  formula := 'YEAR("'+s+'")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString(s,
      RPNFunc('YEAR', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 2014);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(2014);


{------------------------------------------------------------------------------}
{                          Statistical functions                               }
{------------------------------------------------------------------------------}

  // AVEDEV
  if AFormat <> sfExcel2 then begin
    inc(Row);
    formula := 'AVEDEV(1,1.1,1.2,0.9,0.8)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(1.0,
        RPNNumber(1.1,
        RPNNumber(1.2,
        RPNNumber(0.9,
        RPNNumber(0.8,
        RPNFunc('AVEDEV', 5, nil))))))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    number := mean(STATS_NUMBERS);
    for k := 0 to High(STATS_NUMBERS) do numberArray[k] := abs(STATS_NUMBERS[k] - number);
      // these values are the absolute deviations from mean (1.0)
    sollValues[Row] := FloatResult(mean(numberArray));
    MyWorksheet.WriteNumber(Row, 2, mean(numberArray));
  end;

// AVERAGE
  inc(Row);
  formula := 'AVERAGE(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('AVERAGE', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, mean(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(mean(STATS_NUMBERS));

  // AVERAGE of cell block
  inc(Row);
  formula := 'AVERAGE(B1:B2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('B1:B2',
      RPNFunc('AVERAGE', 1, nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, mean([cellB1, cellB2]));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(mean([cellB1, cellB2]));

  // COUNT
  inc(Row);
  formula := 'COUNT(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('COUNT', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, 5);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(5);

  // COUNT  of cell range (no empty cells)
  inc(Row);
  formula := 'COUNT(B1:B2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellrange('B1:B2',
      RPNFunc('COUNT', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(2);

  // COUNT of cell range (no empty cells, but with non-numeric cells)
  inc(Row);
  formula := 'COUNT(A1:B2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellrange('A1:B2',
      RPNFunc('COUNT', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, 2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(2);

  // COUNT  of cell range (with empty cells)
  inc(Row);
  formula := 'COUNT(B1:C2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellrange('B1:C2',
      RPNFunc('COUNT', 1, nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 2);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(2);

  // COUNTA
  formula := 'COUNTA(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('COUNTA', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 5);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(5);

  formula := 'COUNTA(1,1.1,1.2,0.9,0.8,"A","B")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNString('A',
      RPNString('B',
      RPNFunc('COUNTA', 7, nil))))))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, 7);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(7);

  inc(Row);
  formula := 'COUNTA(A1:D2,1,2,"A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('A1:D2',
      RPNNumber(1,
      RPNNumber(2,
      RPNString('A',
      RPNFunc('COUNTA', 4, nil)))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, 9);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(9);

  // COUNTBLANK
  if AFormat <> sfExcel2 then begin
    inc(Row);
    formula := 'COUNTBLANK(A1:D2)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('A1:D2',
        RPNFunc('COUNTBLANK', nil))))
    else
      myworksheet.WriteFormula(row, 1, formula);
    MyWorksheet.WriteNumber(Row, 2, 2);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(2);
  end;

  // COUNTIF
{                       currently no support for COUNTIF

  if AFormat <> sfExcel2 then begin
    inc(Row);
    formula := 'COUNTIF(M1:N3,1)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNNumber(1,                    // "1" is in M2
        RPNFunc('COUNTIF', nil)))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 1);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(1);

    inc(Row);
    formula := 'COUNTIF(M1:N3,">=1")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNString('>=1',                    // M2=1, N2=2 --> 2
        RPNFunc('COUNTIF', nil)))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 2);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(2);

    inc(Row);
    formula := 'COUNTIF(M1:N3,"<2")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNString('<2',                    // M2=1, N2=2 --> 1
        RPNFunc('COUNTIF', nil)))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 1);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(1);

    inc(Row);
    formula := 'COUNTIF(M1:N3,"<>2")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNString('<>2',                    // N2=2, 6 other cells --> 5
        RPNFunc('COUNTIF', nil)))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 5);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(5);

    inc(Row);
    formula := 'COUNTIF(M1:N3,M1)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNCellValue('M1',                    // M1="A", N1="A" -> 2
        RPNFunc('COUNTIF', nil)))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    MyWorksheet.WriteNubmer(Row, 2, 2);
    SetLength(sollValues, Row+1);
    sollValues[Row] := IntegerResult(2);
  end;
                                    }
  // MAX
  inc(Row);
  formula := 'MAX(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('MAX', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, MaxValue(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(MaxValue(STATS_NUMBERS));

  // MAX of cell range (no empty cells)
  inc(Row);
  formula := 'MAX(B1:B2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('B1:B2',
      RPNFunc('MAX', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, MaxValue([cellB1, cellB2]));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(MaxValue([cellB1, cellB2]));

  // MAX of cell range  (incl empty cells)
  inc(Row);
  formula := 'MAX(B1:C2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('B1:C2',
      RPNFunc('MAX', 1, nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, MaxValue([cellB1, cellB2]));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(MaxValue([cellB1, cellB2]));

  // MIN
  inc(Row);
  formula := 'MIN(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('MIN', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, MinValue(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(MinValue(STATS_NUMBERS));

  // MIN of cell range (no empty cells)
  inc(Row);
  formula := 'MIN(B1:B2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('B1:B2',
      RPNFunc('MIN', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, MinValue([cellB1, cellB2]));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(MinValue([cellB1, cellB2]));

  // MIN of cell range  (incl empty cells)
  inc(Row);
  formula := 'MIN(B1:C2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('B1:C2',
      RPNFunc('MIN', 1, nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, MinValue([cellB1, cellB2]));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(MinValue([cellB1, cellB2]));

  // PRODUCT
  inc(Row);
  formula := 'PRODUCT(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('PRODUCT', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(STATS_NUMBERS[0]*STATS_NUMBERS[1]*STATS_NUMBERS[2]*STATS_NUMBERS[3]*STATS_NUMBERS[4]);
  MyWorksheet.WriteNumber(Row, 2, sollvalues[Row].ResFloat);

  // STDEV
  inc(Row);
  formula := 'STDEV(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('STDEV', 5, nil))))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, StdDev(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(StdDev(STATS_NUMBERS));

  // STDEVP
  inc(Row);
  formula := 'STDEVP(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('STDEVP', 5, nil))))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, PopnStdDev(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(PopnStdDev(STATS_NUMBERS));

  // SUM
  inc(Row);
  formula := 'SUM(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('SUM', 5, nil))))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, Sum(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(Sum(STATS_NUMBERS));

  // SUMSQ
  if AFormat <> sfExcel2 then begin
    inc(Row);
    formula := 'SUMSQ(1,1.1,1.2,0.9,0.8)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNNumber(1.0,
        RPNNumber(1.1,
        RPNNumber(1.2,
        RPNNumber(0.9,
        RPNNumber(0.8,
        RPNFunc('SUMSQ', 5, nil))))))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    MyWorksheet.WriteNumber(Row, 2, SumOfSquares(STATS_NUMBERS));
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(SumOfSquares(STATS_NUMBERS));

                      {   ---- SUMIF currently not supported
    // SUMIF
    inc(Row);
    formula = 'SUMIF(M1:N3,1)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNNumber(1,                    // "1" is in M2
        RPNFunc('SUMIF', 2, nil)))))
    else
      MyWorksheet.WriteRPNFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 1);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(1);

    inc(Row);
    formula := 'SUMIF(M1:N3,">=1")';
    MyWorksheet.WriteUTF8Text(Row, 0, formul);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNString('>=1',                    // M2=1, N2=2 --> 2
        RPNFunc('SUMIF', 2, nil)))))
    else
      MyWorksheet.WriteFormula(Row, 1 formula);
    MyWorksheet.WriteNumber(Row, 2, 1.0+2.0);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(1.0 + 2.0);

    inc(Row);
    formula := 'SUMIF(M1:N3,"<2")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNString('<2',                    // M2=1, N2=2 --> 1
        RPNFunc('SUMIF', 2, nil)))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    MyWorksheet.WriteFormula(Row, 2, 1);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(1);

    inc(Row);
    formula := 'SUMIF(M1:N3,"<>2")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNString('<>2',                    // N2=2, 6 other cells --> 5
        RPNFunc('SUMIF', 2, nil)))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 1);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(1);

    inc(Row);
    formula := 'SUMIF(M1:N3,M1)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNCellValue('M1',                    // M1="A", N1="A" -> 2
        RPNFunc('SUMIF', 2, nil)))))
    else
      MyWorksheet.WriteFormula(Row, 1, formula);
    Myworksheet.WriteNumber(Row, 2, 0);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(0);  // no numbers!

    inc(Row);
    formula := 'SUMIF(M1:N3,M2)';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNCellRange('M1:N3',
        RPNCellValue('M2',                    // M2=1
        RPNFunc('SUMIF', 2, nil)))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    MyWorksheet.WriteNumber(Row, 2, 1);
    SetLength(sollValues, Row+1);
    sollValues[Row] := FloatResult(1);
    }
  end;

  // VAR
  inc(Row);
  formula := 'VAR(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('VAR', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, Variance(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(Variance(STATS_NUMBERS));

  // VARP
  inc(Row);
  formula := 'VARP(1,1.1,1.2,0.9,0.8)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1.0,
      RPNNumber(1.1,
      RPNNumber(1.2,
      RPNNumber(0.9,
      RPNNumber(0.8,
      RPNFunc('VARP', 5, nil))))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, PopnVariance(STATS_NUMBERS));
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(PopnVariance(STATS_NUMBERS));


{------------------------------------------------------------------------------}
{                            Logical functions                                 }
{------------------------------------------------------------------------------}

  // AND of one values (bool)
  inc(Row);
  formula := 'AND(TRUE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNFunc('AND', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'TRUE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);

  // AND of two values (bool)
  inc(Row);
  formula := 'AND(TRUE,FALSE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc('AND', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'FALSE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true and false);

  // AND of three values (bool)
  inc(Row);
  formula := 'AND(TRUE,FALSE,TRUE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNBool(true,
      RPNFunc('AND', 3, nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true and false and true);

  // function =FALSE()
  inc(Row);
  formula := 'FALSE()';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc('FALSE', nil)))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'FALSE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);

  // IF (2 parameters)/strings/case true
  inc(Row);
  formula := 'IF(TRUE,"A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNString('A',
      RPNFunc('IF', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'A');
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('A');

  // IF (2 parameters) /floats/case true
  inc(Row);
  formula := 'IF(TRUE,1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNNumber(1.0,
      RPNFunc('IF', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(1);

  // IF (2 parameters)/strings/case false
  inc(Row);
  formula := 'IF(FALSE,"A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(false,
      RPNString('A',
      RPNFunc('IF', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);

  // IF (2 parameters) /floats/case false
  inc(Row);
  formula := 'IF(FALSE,1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(false,
      RPNNumber(1.0,
      RPNFunc('IF', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteUTF8Text(Row, 2, 'FALSE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);

  // IF (3 parameters)/strings
  inc(Row);
  formula := 'IF(TRUE,"A","B")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNString('A',
      RPNString('B',
      RPNFunc('IF', 3, nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'A');
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('A');

  // IF (3 parameters) /floats
  inc(Row);
  formula := 'IF(TRUE,1,2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNNumber(1.0,
      RPNNumber(2.0,
      RPNFunc('IF',3, nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteNumber(Row, 2, 1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(1);

  // IF (3 parameters) /floats / mixed types, case true
  inc(Row);
  formula := 'IF(TRUE,1,"A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNNumber(1.0,
      RPNString('A',
      RPNFunc('IF',3, nil))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteNumber(Row, 2, 1);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(1);

  // IF (3 parameters) /floats / mixed types, case false
  inc(Row);
  formula := 'IF(FALSE,1,"A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(false,
      RPNNumber(1.0,
      RPNString('A',
      RPNFunc('IF', 3, nil))))))
  else
    myWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'A');
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('A');

  // NOT
  inc(Row);
  formula := 'NOT(FALSE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(false,
      RPNFunc('NOT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'TRUE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(not false);

  // OR of one values (bool)
  inc(Row);
  formula := 'OR(TRUE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNFunc('OR', 1, nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteUTF8Text(Row, 2, 'TRUE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);

  // OR of two values (bool)
  inc(Row);
  formula := 'OR(TRUE,FALSE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNFunc('OR', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := Booleanresult(true or false);

  // OR of three values (bool)
  inc(Row);
  formula := 'OR(TRUE,FALSE,TRUE)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNBool(true,
      RPNBool(false,
      RPNBool(true,
      RPNFunc('OR', 3, nil))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true or false or true);

  // function =TRUE()
  inc(Row);
  formula := 'TRUE()';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
     RPNFunc('TRUE', nil)))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  MyWorksheet.WriteUTF8Text(Row, 2, 'TRUE');
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);


{------------------------------------------------------------------------------}
{                             String functions                                 }
{------------------------------------------------------------------------------}

  // CHAR
  inc(Row);
  formula := 'CHAR(72)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNInteger(72,
      RPNFunc('CHAR', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult(char(72));
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

   // CODE
  inc(Row);
  formula := 'CODE("Hallo world")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNFunc('CODE', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(ord('H'));
  Myworksheet.WriteNumber(Row, 2, sollValues[Row].ResInteger);

  // CONCATENATE
  if AFormat <> sfExcel2 then
  begin
    inc(Row);
    formula := 'CONCATENATE("A","B","C")';
    MyWorksheet.WriteUTF8Text(Row, 0, formula);
    if UseRPNFormula then
      MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
        RPNString('A',
        RPNString('B',
        RPNString('C',
        RPNFunc('CONCATENATE', 3, nil))))))
    else
      Myworksheet.WriteFormula(Row, 1, formula);
    SetLength(sollValues, Row+1);
    sollValues[Row] := StringResult('ABC');
    Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);
  end;

  // EXACT
  inc(Row);
  formula := 'EXACT("abc","ABC")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('abc',
      RPNString('ABC',
      RPNFunc('EXACT', nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // LEFT (2 parameters)
  inc(Row);
  formula := 'LEFT("Hallo world",2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNInteger(2,
      RPNFunc('LEFT', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('Ha');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // LEFT (2 parameters, utf8)
  inc(Row);
  formula := 'LEFT("Ändern",3)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Ändern',
      RPNInteger(3,
      RPNFunc('LEFT', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('Änd');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // LEFT (2 parameters, 1 of them missing)
  inc(Row);
  formula := 'LEFT("Hallo world",)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNMissingArg(
      RPNFunc('LEFT', 2, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('H');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // LEFT (1 parameter)
  inc(Row);
  formula := 'LEFT("Hallo world")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNFunc('LEFT', 1, nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('H');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // Len
  inc(Row);
  formula := 'LEN("Hallo world")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNFunc('LEN', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := IntegerResult(Length('Hallo world'));
  Myworksheet.WriteNumber(Row, 2, sollValues[Row].ResInteger);

  // Lower case
  inc(Row);
  formula := 'LOWER("Hallo world")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNFunc('LOWER', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult(LowerCase('Hallo world'));
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // Lower case / utf8
  inc(Row);
  formula := 'LOWER("Viele Grüße")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Viele Grüße',
      RPNFunc('LOWER', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult(UTF8LowerCase('Viele Grüße'));
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // MID
  inc(Row);
  formula := 'MID("Hallo world",3,2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNInteger(3,
      RPNInteger(2,
      RPNFunc('MID', nil))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('ll');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // REPLACE
  inc(Row);
  formula := 'REPLACE("weather",2,2,"he")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('weather',
      RPNInteger(2,
      RPNInteger(2,
      RPNString('he',
      RPNFunc('REPLACE', nil)))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('whether');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // REPLACE / utf8
  inc(Row);
  formula := 'REPLACE("würde",2,1,"u")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('würde',
      RPNInteger(2,
      RPNInteger(1,
      RPNString('u',
      RPNFunc('REPLACE', nil)))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('wurde');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // REPT
  inc(Row);
  formula := 'REPT("ABC",3)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('ABC',
      RPNInteger(3,
      RPNFunc('REPT', nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('ABCABCABC');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // RIGHT (2 parameters)
  inc(Row);
  formula := 'RIGHT("Hallo world",2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNInteger(2,
      RPNFunc('RIGHT', 2, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('ld');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // RIGHT (2 parameters, one of them missing)
  inc(Row);
  formula := 'RIGHT("Hallo world",)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNMissingArg(
      RPNFunc('RIGHT', 2, nil)))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('d');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // RIGHT (1 parameter)
  inc(Row);
  formula := 'RIGHT("Hallo world")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNFunc('RIGHT', 1, nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('d');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // SUBSTITUTE
  inc(Row);
  formula := 'SUBSTITUTE("lAzArus","A","a")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('lAzArus',
      RPNString('A',
      RPNString('a',
      RPNFunc('SUBSTITUTE', 3, nil))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('lazarus');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // SUBSTITUTE /nth appearance
  inc(Row);
  formula := 'SUBSTITUTE("lazarus","a","A",2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('lazarus',
      RPNString('a',
      RPNString('A',
      RPNInteger(2,
      RPNFunc('SUBSTITUTE', 4, nil)))))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('lazArus');
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // Trim
  inc(Row);
  formula := 'TRIM("   Hallo world   ")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('   Hallo world   ',
      RPNFunc('TRIM', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult(UTF8Trim('   Hallo world   '));
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // Upper case
  inc(Row);
  formula := 'UPPER("Hallo world")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Hallo world',
      RPNFunc('UPPER', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult(UpperCase('Hallo world'));
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // Upper case / utf8
  inc(Row);
  formula := 'UPPER("Viele Grüße")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('Viele Grüße',
      RPNFunc('UPPER', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult(UTF8UpperCase('Viele Grüße'));
  Myworksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  // VALUE
  inc(Row);
  formula := 'VALUE("100")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('100',
      RPNFunc('VALUE', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  myWorksheet.WriteNumber(Row, 2, 100);
  SetLength(sollValues, Row+1);
  sollValues[Row] := FloatResult(100);








{------------------------------------------------------------------------------}
{                        Lookup / reference functions                           }
{------------------------------------------------------------------------------}
            (*
  // COLUMN
  MyWorksheet.WriteUTF8Text(Row, 0, 'COLUMN(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekCOLUMN, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, 'COLUMN(C1:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C1:D3',
    RPNFunc(fekCOLUMN, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(3);

  // COLUMNS
  MyWorksheet.WriteUTF8Text(Row, 0, 'COLUMNS(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekCOLUMNS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, 'COLUMNS(C1:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C1:D3',
    RPNFunc(fekCOLUMNS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  // ROW
  MyWorksheet.WriteUTF8Text(Row, 0, 'ROW(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekROW, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, 'ROW(C2:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C2:D3',
    RPNFunc(fekROW, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

  // ROWS
  MyWorksheet.WriteUTF8Text(Row, 0, 'ROWS(A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef('A1',
    RPNFunc(fekROWS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);

  MyWorksheet.WriteUTF8Text(Row, 0, 'ROWS(C2:D3)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('C2:D3',
    RPNFunc(fekROWS, 1, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);

                *)
{------------------------------------------------------------------------------}
{                           Information functions                              }
{------------------------------------------------------------------------------}
                   { cell function will be removed because it requires localized string parameters
  // CELL
  formula := 'CELL("address",A1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('address',
      RPNCellRef('A1',
      RPNFunc('CELL', 2, nil)))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := StringResult('$A$1');
  MyWorksheet.WriteUTF8Text(Row, 2, sollValues[Row].ResString);

  MyWorksheet.WriteUTF8Text(Row, 0, 'CELL("col", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('col',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(2);      // Excel starts counting at 1

  MyWorksheet.WriteUTF8Text(Row, 0, 'CELL("format", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('format',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('G');

  MyWorksheet.WriteUTF8Text(Row, 0, 'CELL("prefix", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('prefix',
    RPNCellRef('A1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('''');

  MyWorksheet.WriteUTF8Text(Row, 0, 'CELL("row", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('row',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(1);      // Excel starts counting at 1

  MyWorksheet.WriteUTF8Text(Row, 0, 'CELL("type", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('type',
    RPNCellRef('A1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('l');

  MyWorksheet.WriteUTF8Text(Row, 0, 'CELL("type", B1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('type',
    RPNCellRef('B1',
    RPNFunc(fekCELLINFO, 2, nil)))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg('v');
}

(*
  // INFO
  MyWorksheet.WriteUTF8Text(Row, 0, 'INFO("directory", A1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('directory',
    RPNFunc(fekINFO, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateStringArg(ExtractFilePath(TempFile));

  MyWorksheet.WriteUTF8Text(Row, 0, 'INFO("numfile")');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('numfile',
    RPNFunc(fekINFO, nil))));
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateNumberArg(MyWorkbook.GetWorksheetCount);
           *)

  // ERROR.TYPE
(*               ---- fails for ODS ---> to be checked
  inc(Row);
  formula := 'ERROR.TYPE(A1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('A1',
      RPNFUNC('ERROR.TYPE', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollvalues, Row+1);
  sollValues[Row] := ErrorResult(errArgError);
  MyWorksheet.WriteUTF8Text(Row, 2, '#N/A');

  inc(Row);
  formula := 'ERROR.TYPE(H2)';   // this cells contains "1/0" --> #DIV/0!
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('H2',
      RPNFUNC('ERROR.TYPE', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollvalues, Row+1);
  sollValues[Row] := IntegerResult(ord(errDivideByZero));
  MyWorksheet.WriteUTF8Text(Row, 2, IntToStr(ord(errDivideByZero)));
  *)

  // IsBlank
  inc(Row);
  formula := 'ISBLANK(A1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('A1',
      RPNFunc('ISBLANK', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // cell contains text --> not blank
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISBLANK(G1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('G1',
      RPNFunc('ISBLANK', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // the cell does not exist --> blank
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'ISBLANK(H1)';
  MyWorksheet.WriteBlank(0, 7);          // A11 is an empty cell
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('H1',
      RPNFunc('ISBLANK', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // the cell exists, but it is empty
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // IsErr
  inc(Row);
  formula := 'ISERR(H2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    myWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('H2',
      RPNFunc('ISERR', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);  // there is an error in H2
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  formula := 'ISERR(H3)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    myWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('H3',
      RPNFunc('ISERR', nil))))
  else
    Myworksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);  // no error in H3
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // IsError
  inc(Row);
  formula := 'ISERROR(1/0)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNNumber(0,
      RPNFunc(fekDiv,
      RPNFunc('ISERROR', nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'ISERROR(0/1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0,
      RPNNumber(1,
      RPNFunc(fekDiv,
      RPNFunc('ISERROR', nil))))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISERROR(H2)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('H2',
      RPNFunc('ISERROR', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'ISERROR(H3)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('H3',
      RPNFunc('ISERROR', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // IsLogical
  inc(Row);
  formula := 'ISLOGICAL(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('B1',
      RPNFunc('ISLOGICAL', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // IsNONTEXT
  inc(Row);
  formula := 'ISNONTEXT(1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNFunc('ISNONTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'ISNONTEXT("A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('A',
      RPNFunc('ISNONTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISNONTEXT(A1)';             // A1 contains a text
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('A1',
      RPNFunc('ISNONTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISNONTEXT(B1)';              // B1 contains a number
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('B1',
      RPNFunc('ISNONTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // IsNumber
  inc(Row);
  formula := 'ISNUMBER(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('B1',
      RPNFunc('ISNUMBER', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // IsRef
  inc(Row);
  formula := 'ISREF(1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNFunc('ISREF', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISREF(A1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('A1',
      RPNFunc('ISREF', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISREF(A1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('A1',                 // we use a cell value here !
      RPNFunc('ISREF', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  // The correct result would be "false" because a cell value is not the same as
  // a cell reference. But Excel seems to ignore this difference here and
  // accepts only a "true".
  sollValues[Row] := BooleanResult(true);
  MyWorksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // IsText
  inc(Row);
  formula := 'ISTEXT(1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1,
      RPNFunc('ISTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  formula := 'ISTEXT("A")';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('A',
      RPNFunc('ISTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'ISTEXT(A1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('A1',
      RPNFunc('ISTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(true);
  Myworksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  formula := 'ISTEXT(B1)';
  MyWorksheet.WriteUTF8Text(Row, 0, formula);
  if UseRPNFormula then
    MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRef('B1',
      RPNFunc('ISTEXT', nil))))
  else
    MyWorksheet.WriteFormula(Row, 1, formula);
  SetLength(sollValues, Row+1);
  sollValues[Row] := BooleanResult(false);
  Myworksheet.WriteUTF8Text(Row, 2, 'FALSE');


{------------------------------------------------------------------------------}
{                                Error cases                                   }
{------------------------------------------------------------------------------}
{$IFDEF ENABLE_DEFECTIVE_FORMULAS }
  // Using less parameters than specified
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, 'IF(true,1)');
  MyWorksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    RPNNumber(1.0,
    RPNFunc(fekIF,3, nil)))));            // <-- we have only 2 parameters!
  SetLength(sollValues, Row+1);
  sollValues[Row] := CreateErrorArg(errArgError);

  // Using more parameters than specified
  inc(Row);
  MyWorksheet.WriteUTF8Text(Row, 0, 'IF(true,1,"A")');
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

