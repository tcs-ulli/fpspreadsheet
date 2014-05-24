unit rpnFormulaUnit;

interface

uses
  SysUtils, fpspreadsheet,fpsutils;

procedure WriteRPNFormulaSamples(Worksheet: TsWorksheet; AFormat: TsSpreadsheetFormat);

implementation

uses
  Math, StrUtils;

const
  FALSE_TRUE: array[Boolean] of String = ('FALSE', 'TRUE');

procedure WriteRPNFormulaSamples(Worksheet: TsWorksheet; AFormat: TsSpreadSheetFormat);
const
  cellB1 = 1.0;
  cellC1 = 2.0;
  cellD1 = 3.0;
  cellE1 = -1.0;
  cellF1 = 1.4567;
  cellG1 = -1.4567;
  SBaseCells = 'Data cells:';
  SHelloWorld = 'Hello world!';
var
  Cell: PCell;
  Row: Integer;
  value: Double;
  r,c: integer;
  celladdr: String;
begin
  if Worksheet = nil then
    exit;

  Worksheet.WriteUTF8Text(0, 0, SBaseCells);
  Worksheet.WriteUsedFormatting(0, 0, [uffBold]);
  Worksheet.WriteNumber(0,1, cellB1);
  Worksheet.WriteNumber(0,2, cellC1);
  Worksheet.WriteNumber(0,3, cellD1);
  Worksheet.WriteNumber(0,4, cellE1);
  Worksheet.WriteNumber(0,5, cellF1);
  Worksheet.WriteNumber(0,6, cellG1);

  Row := 2;
  Worksheet.WriteUTF8Text(Row, 1, 'read value');
  Worksheet.WriteUsedFormatting(Row, 1, [uffBold]);
  Worksheet.WriteUTF8Text(Row, 2, 'expected value');
  Worksheet.WriteUsedFormatting(Row, 2, [uffBold]);

  { ---------- }

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Constants');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Numbers
  inc(Row);
  value := 1.2345;
  Worksheet.WriteUTF8Text(Row, 0, '=1.2345');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    nil))
  );
  Worksheet.WriteNumber(Row, 2, value);

  // Strings
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('="%s"', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    nil)));
  Worksheet.WriteUTF8Text(Row, 2, SHelloWorld);

  // Boolean
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=TRUE');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNBool(true,
    nil)));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Cell references - please check formula in editing line');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Absolute col and row references
  inc(Row);
  cellAddr := '$B$1';
  Worksheet.WriteUTF8Text(Row, 0, '='+cellAddr);
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue(cellAddr,
    nil
  )));
  Worksheet.WriteNumber(Row, 2, cellB1);

  // Relative col and row references
  inc(Row);
  cellAddr := 'B1';
  Worksheet.WriteUTF8Text(Row, 0, '='+cellAddr);
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue(cellAddr,
    nil
  )));
  Worksheet.WriteNumber(Row, 2, cellB1);

  // Relative row reference
  inc(Row);
  cellAddr := '$B1';
  Worksheet.WriteUTF8Text(Row, 0, '='+cellAddr);
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B1',
    nil
  )));
  Worksheet.WriteNumber(Row, 2, cellB1);

  // Relative col reference
  inc(Row);
  cellAddr := 'B$1';
  Worksheet.WriteUTF8Text(Row, 0, '='+cellAddr);
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B$1',
    nil
  )));
  Worksheet.WriteNumber(Row, 2, cellB1);

  // Relative block reference
  inc(Row);
  cellAddr := 'A1:G2';
  Worksheet.WriteUTF8Text(Row, 0, Format('=COUNT(%s)', [cellAddr]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange(cellAddr,
    RPNFunc(fekCOUNT, 1,     // 1 parameter used in COUNT
    nil
  ))));
  Worksheet.WriteNumber(Row, 2, 6);   // 7 cells, but 1 is alpha-numerical!

  // Relative block cols reference
  inc(Row);
  cellAddr := 'A$1:G$2';
  Worksheet.WriteUTF8Text(Row, 0, Format('=COUNT(%s)', [cellAddr]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange(cellAddr,
    RPNFunc(fekCOUNT, 1,
    nil
  ))));
  Worksheet.WriteNumber(Row, 2, 6);   // 7 cells, but 1 is alph-numerical!

  // Relative block rows reference
  inc(Row);
  cellAddr := '$A1:$G2';
  Worksheet.WriteUTF8Text(Row, 0, Format('=COUNT(%s)', [cellAddr]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange(cellAddr,
    RPNFunc(fekCOUNT, 1,
    nil
  ))));
  Worksheet.WriteNumber(Row, 2, 6);   // 7 cells, but 1 is alpha-numerical!

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Basic operations');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Add two cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=B1+C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('C1',
    RPNFunc(fekAdd,
    nil
  )))));
  Worksheet.WriteNumber(Row, 2, cellB1 + cellC1);

  // Add three cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=B1+C1+$D$1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('C1',
    RPNCellvalue('$D$1',
    RPNFunc(fekAdd,
    RPNFunc(fekAdd,
    nil
  )))))));
  Worksheet.WriteNumber(Row, 2, cellB1 + cellC1 + cellD1);

  // Subtract two cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=B1-C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('C1',
    RPNFunc(fekSub,
    nil
  )))));
  Worksheet.WriteNumber(Row, 2, cellB1 - cellC1);

  // Multiply two (absolute) cells = $C$1*$D$1
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=$C$1*$D$1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$C$1',
    RPNCellValue('$D$1',
    RPNFunc(fekMul,
    nil
  )))));
  Worksheet.WriteNumber(Row, 2, cellC1 * cellD1);

  // Divide two (relative) cells = C1/D1
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1/D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekDiv,
    nil
  )))));
  Worksheet.WriteNumber(Row, 2, cellC1 / cellD1);

  // Power of two cells = C1^D1
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1^D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekPower,
    nil
  )))));
  Worksheet.WriteNumber(Row, 2, Power(cellC1, cellD1));

  // Unary plus =+C1
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=+C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNFunc(fekUPlus,
    nil
  ))));
  Worksheet.WriteNumber(Row, 2, +cellC1);

  // Unary minus =-$C$1
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=-$C$1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$C$1',
    RPNFunc(fekUMinus,
    nil
  ))));
  Worksheet.WriteNumber(Row, 2, -cellC1);

  // Percent
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=B1*$C$1%');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNCellValue('$C$1',
    RPNFunc(fekPercent,
    RPNFunc(fekMul,
    nil
  ))))));
  Worksheet.WriteUTF8Text(Row, 2, FloatToStr(cellB1*cellC1/100));

  // String concatenation
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '="Hello " & "world"');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('Hello ',
    RPNString('world',
    RPNFunc(fekConcat,
    nil
  )))));
  Worksheet.WriteUTF8Text(Row, 2, 'Hello ' + 'world');

  // Less Than - case a
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1<D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekLess,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1<cellD1]);

  // Less Than - case b
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=D1<C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('D1',
    RPNCellValue('C1',
    RPNFunc(fekLess,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellD1<cellC1]);

  // Less Than - case c (equal)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1<C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekLess,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1<cellC1]);

  // Less Than or Equal - case a
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1<=D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekLessEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1<=cellD1]);

  // LessThan or Equal - case b
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=D1<=C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('D1',
    RPNCellValue('C1',
    RPNFunc(fekLessEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellD1<=cellC1]);

  // LessThan or Equal - case c
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1<=C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekLessEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1<=cellC1]);

  // Greater Than - case a
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1>D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekGreater,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1>cellD1]);

  // Greater Than - case b
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=D1>C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('D1',
    RPNCellValue('C1',
    RPNFunc(fekGreater,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellD1>cellC1]);

  // Greater Than - case c
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1>C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekGreater,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1>cellC1]);

  // Greater Than or Equal - case a
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1>=D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekGreaterEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1>=cellD1]);

  // Greater Than or Equal - case b
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=D1>=C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('D1',
    RPNCellValue('C1',
    RPNFunc(fekGreaterEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellD1>=cellC1]);

  // Greater Than or Equal - case c
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1>=C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekGreaterEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1>=cellC1]);

  // Equal - case a
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1=D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekEqual,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1=cellD1]);

  // Equal - case b
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=D1=C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('D1',
    RPNCellValue('C1',
    RPNFunc(fekEQUAL,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellD1=cellC1]);

  // Equal - case c
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1=C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekEQUAL,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1=cellC1]);

  // Not equal - case a
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1<>D1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('D1',
    RPNFunc(fekNotEQUAL,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1<>cellD1]);

  // Not equal - case b
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=D1<>C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('D1',
    RPNCellValue('C1',
    RPNFunc(fekNotEQUAL,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellD1<>cellC1]);

  // Not equal - case c
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=C1<>C1');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekNotEQUAL,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[cellC1<>cellC1]);

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Logical functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // TRUE()
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=TRUE()');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekTRUE,
    nil)));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[true]);

  // FALSE()
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=FALSE()');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekFALSE,
    nil)));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[false]);


  // Logical NOT
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=NOT(C1=C1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('C1',
    RPNCellValue('C1',
    RPNFunc(fekEQUAL,
    RPNFunc(fekNOT,
    nil))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[not (cellC1=cellC1)]);

  // Logical AND - case false/false
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=AND(1=0,1=2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekEQUAL,
    RPNNumber(1,
    RPNNumber(2,
    RPNFunc(fekEQUAL,
    RPNFunc(fekAND, 2,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[(1=0) and (1=2)]);

  // Logical AND - case false/true
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=AND(1=0,2=2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekEQUAL,
    RPNNumber(2,
    RPNNumber(2,
    RPNFunc(fekEQUAL,
    RPNFunc(fekAND, 2,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[(1=0) and (2=2)]);

  // Logical AND - case true/true
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=AND(1=1,2=2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(1,
    RPNFunc(fekEQUAL,
    RPNNumber(2,
    RPNNumber(2,
    RPNFunc(fekEQUAL,
    RPNFunc(fekAND, 2,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[(1=1) and (2=2)]);

  // Logical OR - case false/false
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=OR(1=0,1=2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekEQUAL,
    RPNNumber(1,
    RPNNumber(2,
    RPNFunc(fekEQUAL,
    RPNFunc(fekOR, 2,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[(1=0) or (1=2)]);

  // Logical OR - case false/true
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=OR(1=0,2=2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekEQUAL,
    RPNNumber(2,
    RPNNumber(2,
    RPNFunc(fekEQUAL,
    RPNFunc(fekOR, 2,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[(1=0) or (2=2)]);

  // Logical OR - case true/true
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=OR(1=1,2=2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(1,
    RPNFunc(fekEQUAL,
    RPNNumber(2,
    RPNNumber(2,
    RPNFunc(fekEQUAL,
    RPNFunc(fekOR, 2,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, FALSE_TRUE[(1=1) or (2=2)]);

  // IF - case true
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=IF(B1=1,"correct","wrong")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNNumber(1,
    RPNFunc(fekEQUAL,
    RPNString('correct',
    RPNString('wrong',
    RPNFunc(fekIF, 3,
    nil))))))));
  Worksheet.WriteUTF8Text(Row, 2, IfThen(cellB1=1.0, 'correct', 'wrong'));

  // IF - case false
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=IF(B1<>1,"correct","wrong")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNNumber(1,
    RPNFunc(fekNotEQUAL,
    RPNString('correct',
    RPNString('wrong',
    RPNFunc(fekIF, 3,
    nil))))))));
  Worksheet.WriteUTF8Text(Row, 2, IfThen(cellB1<>1.0, 'correct', 'wrong'));

  // IF - case true (2 params)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=IF(B1=1,"correct")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNNumber(1,
    RPNFunc(fekEQUAL,
    RPNString('correct',
    RPNFunc(fekIF, 2,
    nil)))))));
  Worksheet.WriteUTF8Text(Row, 2, IfThen(cellB1=1.0, 'correct', 'FALSE'));

  // IF - case false (2 params)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=IF(B1<>1,"correct")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNNumber(1,
    RPNFunc(fekNotEQUAL,
    RPNString('correct',
    RPNFunc(fekIF, 2,
    nil)))))));
  Worksheet.WriteUTF8Text(Row, 2, IfThen(cellB1<>1.0, 'correct', 'FALSE'));

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Math functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // absolute of positive number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ABS($B1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B1',
    RPNFunc(fekABS,
    nil))));
  Worksheet.WriteNumber(Row, 2, abs(cellB1));

  // absolute of negative number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ABS(E$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('E$1',
    RPNFunc(fekABS,
    nil))));
  Worksheet.WriteNumber(Row, 2, abs(cellE1));

  // sign of positive number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SIGN(F1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('F1',
    RPNFunc(fekSIGN,
    nil))));
  Worksheet.WriteNumber(Row, 2, sign(cellF1));

  // sign of zero
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SIGN(0)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(0,
    RPNFunc(fekSIGN,
    nil))));
  Worksheet.WriteNumber(Row, 2, sign(0));

  // sign of negative number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SIGN(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekSIGN,
    nil))));
  Worksheet.WriteNumber(Row, 2, sign(cellG1));

  // Random number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=RAND()');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekRAND,
    nil)));
  Worksheet.WriteUTF8Text(Row, 2, '(random number - cannot compare)');

  // pi
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=PI()');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI,
    nil)));
  Worksheet.WriteNumber(Row, 2, pi);

  if AFormat <> sfExcel2 then begin
    // Degrees
    inc(Row);
    value := pi/2;
    Worksheet.WriteUTF8Text(Row, 0, '=DEGREES(PI()/2)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc(fekPI,
      RPNNumber(2,
      RPNFunc(fekDIV,
      RPNFunc(fekDEGREES,
      nil))))));
    Worksheet.WriteNumber(Row, 2, value/pi*180);

    // Radians
    inc(Row);
    value := 90;
    Worksheet.WriteUTF8Text(Row, 0, '=RADIANS(90)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekRADIANS,
      nil))));
    Worksheet.WriteNumber(Row, 2, value/180*pi);
  end;

  // sin(pi/2)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SIN(PI()/2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI,
    RPNNumber(2,
    RPNFunc(fekDIV,
    RPNFunc(fekSIN,
    nil))))));
  Worksheet.WriteNumber(Row, 2, sin(pi/2));

  // arcsin(0.5)
  inc(Row);
  value := 0.5;
  Worksheet.WriteUTF8Text(Row, 0, '=ASIN(0.5)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekASIN,
    nil))));
  Worksheet.WriteNumber(Row, 2, arcsin(value));

  // cos(pi)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=COS(PI())');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI,
    RPNFunc(fekCOS,
    nil))));
  Worksheet.WriteNumber(Row, 2, cos(pi));

  // arccos(0.5)
  inc(Row);
  value := 0.5;
  Worksheet.WriteUTF8Text(Row, 0, '=ACOS(0.5)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekACOS,
    nil))));
  Worksheet.WriteNumber(Row, 2, arccos(value));

  // tan(pi/4)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=TAN(PI()/4)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekPI,
    RPNNumber(4,
    RPNFunc(fekDiv,
    RPNFunc(fekTAN,
    nil))))));
  Worksheet.WriteNumber(Row, 2, tan(pi/4));

  // arctan(1)
  inc(Row);
  value := 1.0;
  Worksheet.WriteUTF8Text(Row, 0, '=ATAN(1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekATAN,
    nil))));
  Worksheet.WriteNumber(Row, 2, arctan(1.0));

  if AFormat <> sfExcel2 then begin
    // Next functions are not available in BIFF2

    // sinh
    inc(Row);
    value := 3;
    Worksheet.WriteUTF8Text(Row, 0, '=SINH(3)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekSINH,
      nil))));
    Worksheet.WriteNumber(Row, 2, sinh(value));

    // arcsinh
    inc(Row);
    value := 0.5;
    Worksheet.WriteUTF8Text(Row, 0, '=ASINH(0.5)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekASINH,
      nil))));
    Worksheet.WriteNumber(Row, 2, arcsinh(value));

    // cosh
    inc(Row);
    value := 3;
    Worksheet.WriteUTF8Text(Row, 0, '=COSH(3)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekCOSH,
      nil))));
    Worksheet.WriteNumber(Row, 2, cosh(value));

    // arccosh
    inc(Row);
    value := 10.0;
    Worksheet.WriteUTF8Text(Row, 0, '=ACOSH(10.0)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekACOSH,
      nil))));
    Worksheet.WriteNumber(Row, 2, arccosh(value));

    // tanh
    inc(Row);
    value := 3;
    Worksheet.WriteUTF8Text(Row, 0, '=TANH(3)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekTANH,
      nil))));
    Worksheet.WriteNumber(Row, 2, tanh(value));

    // arctanh
    inc(Row);
    value := 0.5;
    Worksheet.WriteUTF8Text(Row, 0, '=ATANH(0.5)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNFunc(fekATANH,
      nil))));
    Worksheet.WriteNumber(Row, 2, arctanh(value));
  end;

  // sqrt(2.0);
  inc(Row);
  value := 2.0;
  Worksheet.WriteUTF8Text(Row, 0, '=SQRT(2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekSQRT,
    nil))));
  Worksheet.WriteNumber(Row, 2, sqrt(value));

  // exp(2)
  inc(Row);
  value := 2.0;
  Worksheet.WriteUTF8Text(Row, 0, '=EXP(2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekEXP,
    nil))));
  Worksheet.WriteNumber(Row, 2, exp(value));

  // ln(2)
  inc(Row);
  value := 2.0;
  Worksheet.WriteUTF8Text(Row, 0, '=LN(2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekLN,
    nil))));
  Worksheet.WriteNumber(Row, 2, ln(value));

  // log to any basis
  if AFormat <> sfExcel2 then begin
    // This test is not working in Excel 2.
    // Not clear if this is correct, need to debug later
    inc(Row);
    value := 256;
    Worksheet.WriteUTF8Text(Row, 0, '=LOG(256,2)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(value,
      RPNNumber(2,
      RPNFunc(fekLOG, 2,
      nil)))));
    Worksheet.WriteNumber(Row, 2, logn(2.0, value));
  end;

  // log10(100)
  inc(Row);
  value := 100;
  Worksheet.WriteUTF8Text(Row, 0, '=LOG10(100)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekLOG10,
    nil))));
  Worksheet.WriteNumber(Row, 2, log10(value));

  // log10(0.01)
  inc(Row);
  value := 0.01;
  Worksheet.WriteUTF8Text(Row, 0, '=LOG10(0.01)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(value,
    RPNFunc(fekLOG10,
    nil))));
  Worksheet.WriteNumber(Row, 2, log10(value));

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Rounding');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Round positive number to 1 decimal
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ROUND($F$1,1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$F$1',
    RPNNumber(1,
    RPNFunc(fekROUND,
    nil)))));
  Worksheet.WriteNumber(Row, 2, Round(cellF1*10)/10); //RoundTo(cellF1, 1));

  // Round negative number to 1 decimal
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ROUND(G1,1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNNumber(1,
    RPNFunc(fekROUND,
    nil)))));
  Worksheet.WriteNumber(Row, 2, Round(cellG1*10)/10); //RoundTo(cellG1, 1));

  // integer part of positive number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=INT(F1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('F1',
    RPNFunc(fekINT,
    nil))));
  Worksheet.WriteNumber(Row, 2, trunc(cellF1));

  // integer part of negative number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=INT(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekINT,
    nil))));
  Worksheet.WriteNumber(Row, 2, floor(cellG1));  // is this true?

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Lookup/reference functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Column of cell
  inc(Row);
  cellAddr := 'AB100';
  Worksheet.WriteUTF8Text(Row, 0, Format('=COLUMN(%s)',[cellAddr]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef(cellAddr,           // note: CellRef instead of CellValue!
    RPNFunc(fekCOLUMN, 1,
    nil))));
  ParseCellString(cellAddr, r,c);
  Worksheet.WriteNumber(Row, 2, c+1);     // +1 because Excel index is 1-based

  // Column count of cell block
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=COLUMNS(A1:C5)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('A1:C5',
    RPNFunc(fekCOLUMNS,
    nil))));
  ParseCellString(cellAddr, r,c);
  Worksheet.WriteNumber(Row, 2, 3);

  // Row count of cell block
  // Row of cell
  inc(Row);
  cellAddr := 'AB100';
  Worksheet.WriteUTF8Text(Row, 0, Format('=ROW(%s)', [cellAddr]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRef(cellAddr,         // note: CellRef instead of CellValue!
    RPNFunc(fekROW, 1,
    nil))));
  ParseCellString(cellAddr, r,c);
  Worksheet.WriteNumber(Row, 2, r+1);     // +1 because Excel index is 1-based

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ROWS(A1:C5)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellRange('A1:C5',
    RPNFunc(fekROWS,
    nil))));
  ParseCellString(cellAddr, r,c);
  Worksheet.WriteNumber(Row, 2, 5);

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Info functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Note: looks as if some of these functions are not updated when loading.
  // Press F2 and ENTER for each formula cell

  // Cell is blank
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISBLANK(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsBLANK,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISBLANK(G2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G2',
    RPNFunc(fekIsBLANK,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Cell is in error
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISERR(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G2',
    RPNFunc(fekIsERR,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISERR(G2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G2',
    RPNFunc(fekIsERR,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // Cell is in error
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISERROR(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsERROR,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISERROR(G2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G2',
    RPNFunc(fekIsERROR,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // Cell is logical
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISLOGICAL(B30)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B30',
    RPNFunc(fekIsLOGICAL,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISLOGICAL(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsLOGICAL,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // Cell has #N/A error
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISNA(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsNA,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISNA(G2)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G2',
    RPNFunc(fekIsNA,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // Cell is number
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISNUMBER(A1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('A1',
    RPNFunc(fekIsNUMBER,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISNUMBER(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsNUMBER,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Cell is refence
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISREF(B9)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B9',
    RPNFunc(fekIsREF,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Cell is text
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISTEXT(A1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('A1',
    RPNFunc(fekIsTEXT,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISTEXT(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsTEXT,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  // Cell is non-text
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISNONTEXT(A1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('A1',
    RPNFunc(fekIsNONTEXT,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'FALSE');

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=ISNONTEXT(G1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('G1',
    RPNFunc(fekIsNONTEXT,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'TRUE');

  // Cell information
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=CELL("Address",B80)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('address',
    RPNCellRef('B80',            // note: CellRef instead of CellValue!
    RPNFunc(fekCELLINFO, 2,
    nil)))));

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=CELL("Filename",B80)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('filename',
    RPNCellRef('B80',            // note: CellRef instead of CellValue!
    RPNFunc(fekCELLINFO, 2,
    nil)))));

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=CELL("Row",B80)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('row',
    RPNCellRef('B80',            // note: CellRef instead of CellValue!
    RPNFunc(fekCELLINFO, 2,
    nil)))));

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=CELL("format",B80)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('format',
    RPNCellRef('B80',            // note: CellRef instead of CellValue!
    RPNFunc(fekCELLINFO, 2,
    nil)))));

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=CELL("color",B80)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('color',
    RPNCellRef('B80',            // note: CellRef instead of CellValue!
    RPNFunc(fekCELLINFO, 2,
    nil)))));

  // Value of cell - only for strings which can be converted to numbers
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=VALUE(B1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('B1',
    RPNFunc(fekVALUE,
    nil))));
  Worksheet.WriteNumber(Row, 2, cellB1);

  // General info
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=INFO("osversion")');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('osversion',
      RPNFunc(fekINFO,
      nil))));

    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=INFO("recalc")');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('recalc',
      RPNFunc(fekINFO,
      nil))));

    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=INFO("release")');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('release',
      RPNFunc(fekINFO,
      nil))));

    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=INFO("system")');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('system',
      RPNFunc(fekINFO,
      nil))));
  end;

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Date/time functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Now
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=NOW()');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNFunc(fekNOW,
    nil)));
  Worksheet.WriteNumber(Row, 2, now);

  // Today
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=TODAY()');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNFunc(fekTODAY,
      nil)));
    Worksheet.WriteNumber(Row, 2, date);
  end;

  // Date  (build date from parts)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=DATE(2014,1,25)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(1,
    RPNNumber(25,
    RPNFunc(fekDATE,
    nil))))));
  Worksheet.WriteNumber(Row, 2, EncodeDate(2014,1,25));

  // DateValue (string to date number)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=DATEVALUE("25.01.2014")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('2014-01-25',
    RPNFunc(fekDATEVALUE,
    nil))));
  Worksheet.WriteNumber(Row, 2, EncodeDate(2014,1,25));

  // DateDifference
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=DATEDIF("2010-01-01",DATE(2014;1;25),"M")');
    // Note: Dates must be ordered: Date1 < Date2 !!!
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNString('2010-01-01',
      RPNNumber(2014,
      RPNNumber(1,
      RPNNumber(25,
      RPNFunc(fekDATE,
      RPNString('M',
      RPNFunc(fekDATEDIF,
      nil)))))))));
    Worksheet.WriteUTF8Text(Row, 2, '(not available in FPC)');
  end;

  // Year
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=YEAR(DATE(2014,1,25))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(1,
    RPNNumber(25,
    RPNFunc(fekDATE,
    RPNFunc(fekYEAR,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 2014);

  // Month
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=MONTH(DATE(2014,1,25))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(1,
    RPNNumber(25,
    RPNFunc(fekDATE,
    RPNFunc(fekMONTH,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 1);

  // Day
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=DAY(DATE(2014,1,25))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(1,
    RPNNumber(25,
    RPNFunc(fekDATE,
    RPNFunc(fekDAY,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 25);

  // Weekday     - 2 params can be used, but not in BIFF2
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=WEEKDAY(DATE(2014,1,25))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2014,
    RPNNumber(1,
    RPNNumber(25,
    RPNFunc(fekDATE,
    RPNFunc(fekWEEKDAY, 1,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, DayOfWeek(EncodeDate(2014,1,25)));

  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=WEEKDAY("2014-01-25")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('2014-01-25',
    RPNFunc(fekWEEKDAY, 1,
    nil))));
  Worksheet.WriteNumber(Row, 2, DayOfWeek(EncodeDate(2014,1,25)));

  // Time  (build time from parts)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=TIME(21, 10, 5)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(21,
    RPNNumber(10,
    RPNNumber(5,
    RPNFunc(fekTIME,
    nil))))));
  Worksheet.WriteNumber(Row, 2, EncodeTime(21, 10, 5, 0));

  // Time value (convert string to time number)
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=TIMEVALUE("21:10:05")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('21:10:05',
    RPNFunc(fekTIMEVALUE,
    nil))));
  Worksheet.WriteNumber(Row, 2, EncodeTime(21, 10, 5, 0));

  // Hour
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=HOUR(TIME(21, 10, 5)))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(21,
    RPNNumber(10,
    RPNNumber(5,
    RPNFunc(fekTIME,
    RPNFunc(fekHOUR,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 21);

  // Minute
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=MINUTE(TIME(21, 10, 5)))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(21,
    RPNNumber(10,
    RPNNumber(5,
    RPNFunc(fekTIME,
    RPNFunc(fekMINUTE,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 10);

  // Second
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SECOND(TIME(21, 10, 5)))');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(21,
    RPNNumber(10,
    RPNNumber(5,
    RPNFunc(fekTIME,
    RPNFunc(fekSECOND,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 5);

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Statistical functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Count - non-empty cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=COUNT($B$1,$C$1,$D$1:F1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:F1',
    RPNFunc(fekCOUNT, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, 5);  // 5 cells in total, all not empty

  // Count - with empty cells & alpha-numeric cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=COUNT($B$1,$C$1,$D$1:$F$2,"ABC")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$2',
    RPNString('ABC',
    RPNFunc(fekCOUNT, 4,
    nil)))))));
  Worksheet.WriteNumber(Row, 2, 5);  // 5 non-empty, 3 empty

  // CountA - empty cells and constants
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=COUNTA($B$1,$C$1,$D$1:$F$2,"ABC","DEF")');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$2', //0,3, 1,5, false,false, false,false,
    RPNString('ABC',
    RPNString('DEF',
    RPNFunc(fekCOUNTA, 5,
    nil))))))));
  Worksheet.WriteNumber(Row, 2, 7);  // 7 non-empty values

  if AFormat <> sfExcel2 then begin
    // CountIF
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=COUNTIF(A1:G1,"<=1")');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('A1:G1',
      RPNString('<=1',
      RPNFunc(fekCOUNTIF, 2,
      nil)))));
    Worksheet.WriteNumber(Row, 2,
      IfThen(cellB1<=1, 1, 0) + IfThen(cellC1<=1, 1, 0) +
      IfThen(cellD1<=1, 1, 0) + IfThen(cellE1<=1, 1, 0) +
      IfThen(cellF1<=1, 1, 0) + IfThen(cellG1<=1, 1, 0)
    );

    // Count blank cells
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=COUNTBLANK(A1:H1)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('A1:H1',
      RPNFunc(fekCOUNTBLANK,
      nil))));
    Worksheet.WriteNumber(Row, 2, 1);
  end;

  // Sum - non-empty cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SUM($B$1,$C$1,D$1:F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:F$1',
    RPNFunc(fekSUM, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, cellB1+cellC1+cellD1+cellE1+cellF1);

  if AFormat <> sfExcel2 then begin
    // SumIF
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=SUMIF(A1:G1,"<=1")');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellRange('A1:G1',
      RPNString('<=1',
      RPNFunc(fekSUMIF, 2,
      nil)))));
    Worksheet.WriteNumber(Row, 2,
      IfThen(cellB1<=1, cellB1, 0) + IfThen(cellC1<=1, cellC1, 0) +
      IfThen(cellD1<=1, cellD1, 0) + IfThen(cellE1<=1, cellE1, 0) +
      IfThen(cellF1<=1, cellF1, 0) + IfThen(cellG1<=1, cellG1, 0)
    );

    // Sum of squares
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=SUMSQ($B$1,$C$1,$D$1:$F$1)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('$B$1',
      RPNCellValue('$C$1',
      RPNCellRange('$D$1:$F$1',
      RPNFunc(fekSUMSQ, 3,
      nil))))));
    Worksheet.WriteNumber(Row, 2, sumofsquares([cellB1, cellC1, cellD1, cellE1, cellF1]));
  end;

  // Product
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=PRODUCT($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekPRODUCT, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, cellB1*cellC1*cellD1*cellE1*cellF1);

  // Average - non-empty cells
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=AVERAGE($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekAVERAGE, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, (cellB1+cellC1+cellD1+cellE1+cellF1)/5);

  // StdDev
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=STDEV($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekSTDEV, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, stddev([cellB1,cellC1,cellD1,cellE1,cellF1]));

  // Population StdDev
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=STDEVP($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekSTDEVP, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, popnstddev([cellB1,cellC1,cellD1,cellE1,cellF1]));

  // Average deviation
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=AVEDEV($B$1,$C$1,$D$1:$F$1)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('$B$1',
      RPNCellValue('$C$1',
      RPNCellRange('$D$1:$F$1',
      RPNFunc(fekAVEDEV, 3,
      nil))))));
    value := mean([cellB1, cellC1, cellD1, cellE1, cellF1]);
    Worksheet.WriteNumber(Row, 2, mean([abs(cellB1-value),abs(cellC1-value),
      abs(cellD1-value),abs(cellE1-value),abs(cellF1-value)]));
  end;

  // Variance
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=VAR($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekVAR, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, variance([cellB1,cellC1,cellD1,cellE1,cellF1]));

  // Population variance
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=VARP($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekVARP, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, popnvariance([cellB1,cellC1,cellD1,cellE1,cellF1]));

  // Max
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=MAX($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekMAX, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, MaxValue([cellB1,cellC1,cellD1,cellE1,cellF1]));

  // Min
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=MIN($B$1,$C$1,$D$1:$F$1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNCellValue('$B$1',
    RPNCellValue('$C$1',
    RPNCellRange('$D$1:$F$1',
    RPNFunc(fekMIN, 3,
    nil))))));
  Worksheet.WriteNumber(Row, 2, MinValue([cellB1,cellC1,cellD1,cellE1,cellF1]));

  // Median
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=MEDIAN($B$1,$C$1,$D$1:$F$1)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNCellValue('$B$1',
      RPNCellValue('$C$1',
      RPNCellRange('$D$1:$F$1',
      RPNFunc(fekMEDIAN, 3,
      nil))))));
    Worksheet.WriteNumber(Row, 2, cellF1);  // manually calculated median...
  end;

  // Beta distribution
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=BETADIST(3,7.5,9,1,4)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(3,
      RPNNumber(7.5,
      RPNNumber(9,
      RPNNumber(1,
      RPNNumber(4,
      RPNFunc(fekBETADIST, 5,
      nil))))))));
    Worksheet.WriteNumber(Row, 2, 0.960370937);  // result according to http://www.techonthenet.com/excel/formulas/betadist.php
  end;

  // Inverse beta function
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=BETAINV(0.3,7.5,9,1,4)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.3,
      RPNNumber(7.5,
      RPNNumber(9,
      RPNNumber(1,
      RPNNumber(4,
      RPNFunc(fekBETAINV, 5,
      nil))))))));
    Worksheet.WriteNumber(Row, 2, 2.164759636);  // result according to http://www.techonthenet.com/excel/formulas/betainv.php
  end;

  // Binomial distribution
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=BINOMDIST(3,8,0.35,TRUE)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(3,
      RPNNumber(8,
      RPNNumber(0.35,
      RPNBool(true,
      RPNFunc(fekBINOMDIST,
      nil)))))));
    Worksheet.WriteNumber(Row, 2, 0.706399436);  // result according to http://www.techonthenet.com/excel/formulas/binomdist.php
  end;

  // Chi2 distribution
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=CHIDIST(3,9)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(3,
      RPNNumber(9,
      RPNFunc(fekCHIDIST,
      nil)))));
    Worksheet.WriteNumber(Row, 2,  0.964294973);  // result according to http://www.techonthenet.com/excel/formulas/chidist.php
  end;

  // Inverse of Chi2 distribution
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=CHIINV(0.3,7)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(0.3,
      RPNNumber(7,
      RPNFunc(fekCHIINV,
      nil)))));
    Worksheet.WriteNumber(Row, 2,  8.38343064);  // result according to http://www.techonthenet.com/excel/formulas/chiinv.php
  end;

  // Permutations
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=PERMUT(21,5)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(21,
      RPNNumber(5,
      RPNFunc(fekPERMUT,
      nil)))));
    Worksheet.WriteNumber(Row, 2, 2441880);  // result according to http://www.techonthenet.com/excel/formulas/permut.php
  end;

  // Poisson distribution
  if AFormat <> sfExcel2 then begin
    inc(Row);
    Worksheet.WriteUTF8Text(Row, 0, '=POISSON(1400,1500,TRUE)');
    Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
      RPNNumber(1400,
      RPNNumber(1500,
      RPNBool(true,
      RPNFunc(fekPOISSON,
      nil))))));
    Worksheet.WriteNumber(Row, 2, 0.004744099);  // result according to http://support.microsoft.com/kb/828130/de
  end;

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'Financial');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Future value of an investment where you deposit $5,000 into a savings account
  // that earns 3.5% annually. You are going to deposit $250 at the beginning of
  // the month, each month, for 2 years.
  // according to: www.techonthenet.com/excel/formulas/fv.php
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=FV(3.5%/12,2*12,-250,-5000,1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(3.5,
    RPNFunc(fekPERCENT,
    RPNNumber(12,
    RPNFunc(fekDIV,
    RPNNumber(2,
    RPNNumber(12,
    RPNFunc(fekMUL,
    RPNNumber(-250,
    RPNNumber(-5000,
    RPNNumber(1,
    RPNFunc(fekFV, 5,
    nil)))))))))))));
  Worksheet.WriteUTF8Text(Row, 2, '(not available in FPC)');

  // Present value of an investment that pays $250 at the end of every month for
  // 2 years. The money paid out will earn 3.5% annually.
  // according to: www.techonthenet.com/excel/formulas/pv.php
  // Note the missing argument!
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=PV(3.5%/12,2*12,250,,0)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(3.5,
    RPNFunc(fekPercent,
    RPNNumber(12,
    RPNFunc(fekDIV,
    RPNNumber(2,
    RPNNumber(12,
    RPNFunc(fekMUL,
    RPNNumber(250,
    RPNMissingArg(
    RPNNumber(0,
    RPNFunc(fekPV, 5,
    nil)))))))))))));
  Worksheet.WriteUTF8Text(Row, 2, '(not available in FPC)');

  // Interest rate on a $5,000 loan where monthly payments of $250 are made for
  // 2 years. All payments are made at the end of the period.
  // Adapted from //www.techonthenet.com/excel/formulas/rate.php
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=RATE(2*12,-250,5000)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(2,
    RPNNumber(12,
    RPNFunc(fekMUL,
    RPNNumber(-250,
    RPNNumber(5000,
    RPNFunc(fekRATE, 3,
    nil))))))));
  {
  Worksheet.WriteUsedFormatting(Row, 1, [uffNumberFormat]);
  Worksheet.WriteNumberFormat(Row, 1, nfPercentage);
  }
  Worksheet.WriteUTF8Text(Row, 2, '(not available in FPC)');

  // Number of monthly payments of $150 for a $5,000 investment that earns
  // 3.5% annually. Payments are due at the end of the period.
  // Adapted from //www.techonthenet.com/excel/formulas/nper.php
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=NPER(3.5%/12,-150,5000)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(3.5,
    RPNFunc(fekPERCENT,
    RPNNumber(12,
    RPNFunc(fekDIV,
    RPNNumber(-150,
    RPNNumber(5000,
    RPNFunc(fekNPER, 3,
    nil)))))))));
  Worksheet.WriteUTF8Text(Row, 2, '(not available in FPC)');

  // Monthly payment on a $5,000 loan at an annual rate of 3.5%. The loan is
  // paid off in 2 years (ie: 2 x 12). All payments are made at the beginning
  // of the period.
  // Adapted from //www.techonthenet.com/excel/formulas/pmt.php
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=PMT(3.5%/12,2*12,5000,0,1)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(3.5,
    RPNFunc(fekPERCENT,
    RPNNumber(12,
    RPNFunc(fekDIV,
    RPNNumber(2,
    RPNNumber(12,
    RPNFunc(fekMUL,
    RPNNumber(5000,
    RPNNumber(0,
    RPNNumber(1,
    RPNFunc(fekPMT, 5,
    nil)))))))))))));
  Worksheet.WriteUTF8Text(Row, 2, '(not available in FPC)');

  { ---------- }

  inc(Row);
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, 'String functions');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Character conversion
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=CHAR(64)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(64,
    RPNFunc(fekCHAR,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, chr(64));

  // Character conversion
  // Note: uses only first character of string
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=CODE("%s")', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNFunc(fekCODE,
    nil))));
  Worksheet.WriteNumber(Row, 2, ord(SHelloWorld[1]));

  // Left part of string
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=LEFT("%s",3)', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNNumber(3,
    RPNFunc(fekLEFT, 2,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, Copy(SHelloWorld, 1, 3));

  // Mid part of string
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=MID("%s",4,5)', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNNumber(4,
    RPNNumber(5,
    RPNFunc(fekMID,
    nil))))));
  Worksheet.WriteUTF8Text(Row, 2, Copy(SHelloWorld, 4, 5));

  // Right part of string
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=RIGHT("%s",3)', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNNumber(3,
    RPNFunc(fekRIGHT, 2,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, Copy(SHelloWorld, Length(SHelloWorld)-3+1, 3));

  // Trimming spaces
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=TRIM("   %s   ")', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString('   '+SHelloWorld+'   ',
    RPNFunc(fekTRIM,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, SHelloWorld);

  // Lower case
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=LOWER("%s")', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNFunc(fekLOWER,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, lowercase(SHelloWorld));

  // Upper case
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=UPPER("%s")', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNFunc(fekUPPER,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, uppercase(SHelloWorld));

  // 1st char upper case
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=PROPER("%s")', [uppercase(SHelloWorld)]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNFunc(fekPROPER,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'Hello World!');

  // replace
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=REPLACE("%s",7,5,"Friend")', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNNumber(7,
    RPNNumber(5,
    RPNString('Friend',
    RPNFunc(fekREPLACE,
    nil)))))));
  Worksheet.WriteUTF8Text(Row, 2, 'Hello Friend!');

  // substitute
  // Note: the function can have an optional parameter. Therefore, you have
  // to specify the actual parameter count.
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, Format('=SUBSTITUTE("%s","l",".")', [SHelloWorld]));
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNString(SHelloWorld,
    RPNString('l',
    RPNString('.',
    RPNFunc(fekSUBSTITUTE, 3,
    nil))))));
  Worksheet.WriteUTF8Text(Row, 2, 'He..o Wor.d!');

  inc(Row, 2);
  Worksheet.WriteUTF8Text(Row, 0, 'Errors');
  Worksheet.WriteUsedFormatting(Row, 0, [uffBold]);

  // Division by 0

  // These tests partly produce an error messsage when the file is read by Excel.
  // In order to avoid confusion they are deactivated by default.
  // Remove the comment below to see these tests.

                    (*
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=1/0');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(0,
    RPNFunc(fekDiv,
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, 'Error #DIV/0!');

  // Not enough operands  for operation
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=1/2 ("2" forgotten from formula)" ');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    // here we'd normally put "RPNFormula(2," - but we "forget" it...
    RPNFunc(fekDiv,
    nil))));
  Worksheet.WriteUTF8Text(Row, 2, 'Error #N/A"');

  // Too many operands given
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=1/2 (too many operands)');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(2,
    RPNNumber(3,       // This line is too much
    RPNFunc(fekDiv,
    nil))))));
  Worksheet.WriteUTF8Text(Row, 2, 'Error #N/A!');

  // Parameter count not specified
  inc(Row);
  Worksheet.WriteUTF8Text(Row, 0, '=SUM(1, 2) (parameter count not specified');
  Worksheet.WriteRPNFormula(Row, 1, CreateRPNFormula(
    RPNNumber(1,
    RPNNumber(2,
    RPNFunc(fekSum,  // We "forget" to specify the number of arguments
    nil)))));
  Worksheet.WriteUTF8Text(Row, 2, 'Error #N/A');
                       *)
end;

end.
