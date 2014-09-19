program demo_expression_parser;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  SysUtils, Classes
  { you can add units after this },
  TypInfo, fpSpreadsheet, fpsUtils, fpsExprParser;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  cell: PCell;
  parser: TsExpressionParser;
  res: TsExpressionResult;
  formula: TsRPNFormula;
  i: Integer;
  s: String;

begin
  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Test');

    worksheet.WriteNumber(0, 0, 1);           // A1
    worksheet.WriteNumber(0, 1, 2.5);         // B1

    {
    worksheet.WriteUTF8Text(0, 0, 'Hallo');         // A1
    worksheet.WriteUTF8Text(0, 1, 'World');         // B1
     }
    //cell := worksheet.WriteFormula(1, 0, '=4+5');    // A2
    //cell := worksheet.WriteFormula(1, 0, 'AND(TRUE(), TRUE(), TRUE())');
    //cell := worksheet.WriteFormula(1, 0, 'SIN(A1+B1)');
    //cell := worksheet.WriteFormula(1, 0, '=TRUE()');
    //cell := worksheet.WriteFormula(1, 0, '=1-(4/2)^2*2-1');    // A2
    //cell := Worksheet.WriteFormula(1, 0, 'datedif(today(),Date(2014,1,1),"D")');
    //cell := Worksheet.WriteFormula(1, 0, 'Day(Date(2014, 1, 12))');
    //cell := Worksheet.WriteFormula(1, 0, 'SUM(1,2,3)');
    //cell := Worksheet.WriteFormula(1, 0, 'CELL("address",A1)');
//    cell := Worksheet.WriteFormula(1, 0, 'REPT("Hallo", 3)');
    cell := Worksheet.WriteFormula(1, 0, '#REF!');

    WriteLn('A1: ', worksheet.ReadAsUTF8Text(0, 0));
    WriteLn('B1: ', worksheet.ReadAsUTF8Text(0, 1));

    parser := TsSpreadsheetParser.Create(worksheet);
    try
      try
        parser.Expression := cell^.FormulaValue;
        res := parser.Evaluate;

        WriteLn('A2: ', parser.Expression);
        Write('Result: ');
        case res.ResultType of
          rtEmpty    : WriteLn('--- empty ---');
          rtBoolean  : WriteLn(BoolToStr(res.ResBoolean, true));
          rtFloat    : WriteLn(FloatToStr(res.ResFloat));
          rtInteger  : WriteLn(IntToStr(res.ResInteger));
          rtDateTime : WriteLn(FormatDateTime('c', res.ResDateTime));
          rtString   : WriteLn(res.ResString);
          rtError    : WriteLn(GetErrorValueStr(res.ResError));
        end;

        WriteLn('Reconstructed string formula: ', parser.Expression);
        WriteLn('Reconstructed localized formula: ', parser.LocalizedExpression[DefaultFormatSettings]);
        formula := parser.RPNFormula;

        for i:=0 to Length(formula)-1 do begin
          Write('  Item ', i, ': token ', GetEnumName(TypeInfo(TFEKind), ord(formula[i].ElementKind)), ' ', formula[i].FuncName);
          case formula[i].ElementKind of
            fekCell    : Write(' / cell: ' +GetCellString(formula[i].Row, formula[i].Col, formula[i].RelFlags));
            fekNum     : Write(' / float value: ', FloatToStr(formula[i].DoubleValue));
            fekInteger : Write(' / integer value: ', IntToStr(formula[i].IntValue));
            fekString  : Write(' / string value: "', formula[i].StringValue, '"');
            fekBool    : Write(' / boolean value: ', BoolToStr(formula[i].DoubleValue <> 0, true));
            fekErr     : Write(' / error value: ', GetErrorValueStr(TsErrorValue(formula[i].IntValue)));
          end;
          WriteLn;
        end;
      finally
        parser.Free;
      end;

    except on E:Exception do
      begin
        WriteLn('Parser/calculation error: ', E.Message);
        raise;
      end;
    end;

    parser := TsSpreadsheetParser.Create(worksheet);
    try
      try
        parser.RPNFormula := formula;
        s := parser.Expression;
        WriteLn('String formula, reconstructed from RPN formula: ', s);
      except on E:Exception do
        begin
          WriteLn('RPN/string formula conversion error: ', E.Message);
          raise;
        end;
      end;
    finally
      parser.Free;
    end;

  finally
    workbook.Free;
  end;

end.

