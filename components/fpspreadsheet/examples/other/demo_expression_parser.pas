program demo_expression_parser;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  SysUtils, Classes
  { you can add units after this },
  TypInfo, fpSpreadsheet, fpsUtils, fpsExprParser;

function Prepare(AFormula: String): String;
begin
  if (AFormula <> '') and (AFormula[1] = '=') then
    Result := Copy(AFormula, 2, Length(AFormula)-1)
  else
    Result := AFormula;
end;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  cell: PCell;
  parser: TsExpressionParser;
  res: TsExpressionResult;
  formula: TsRPNFormula;
  i: Integer;

begin
  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Test');
    {
    worksheet.WriteNumber(0, 0, 2);         // A1
    worksheet.WriteNumber(0, 1, 2.5);         // B1
    }
    worksheet.WriteUTF8Text(0, 0, 'Hallo');         // A1
    worksheet.WriteUTF8Text(0, 1, 'World');         // B1

    //cell := worksheet.WriteFormula(1, 0, '=(A1+2)*3');    // A2
    cell := worksheet.WriteFormula(1, 0, 'A1&" "&B1');

    WriteLn('A1 = ', worksheet.ReadAsUTF8Text(0, 0));
    WriteLn('B1 = ', worksheet.ReadAsUTF8Text(0, 1));

    parser := TsExpressionParser.Create(worksheet);
    try
      parser.Builtins := [bcStrings, bcDateTime, bcMath, bcBoolean, bcConversion, bcData,
        bcVaria, bcUser];
      parser.Expression := Prepare(cell^.FormulaValue.FormulaStr);
      res := parser.Evaluate;

      Write('A2 = ', Prepare(cell^.FormulaValue.FormulaStr), ' = ');
      case res.ResultType of
        rtBoolean  : WriteLn(BoolToStr(res.ResBoolean));
        rtFloat    : WriteLn(FloatToStr(res.ResFloat));
        rtInteger  : WriteLn(IntToStr(res.ResInteger));
        rtDateTime : WriteLn(FormatDateTime('c', res.ResDateTime));
        rtString   : WriteLn(res.ResString);
      end;

      WriteLn('Reconstructed string formula: ', parser.BuildFormula);

      WriteLn('RPN formula:');
      formula := parser.BuildRPNFormula;
      for i:=0 to Length(formula)-1 do begin
        Write('  Item ', i, ': token ', GetEnumName(TypeInfo(TFEKind), ord(formula[i].ElementKind)));
        case formula[i].ElementKind of
          fekCell   : Write(' / cell: ' +GetCellString(formula[i].Row, formula[i].Col, formula[i].RelFlags));
          fekNum    : Write(' / number value: ', FloatToStr(formula[i].DoubleValue));
          fekString : Write(' / string value: ', formula[i].StringValue);
          fekBool   : Write(' / boolean value: ', BoolToStr(formula[i].DoubleValue <> 0));
        end;
        WriteLn;
      end;

    finally
      parser.Free;
    end;

  finally
    workbook.Free;
  end;
end.

