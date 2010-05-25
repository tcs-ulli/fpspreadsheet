unit mainform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Grids, EditBtn, ExtCtrls, fpspreadsheetchart, fpspreadsheetgrid,
  TAGraph, TASeries;

type
  
  { TFPSChartForm }

  TFPSChartForm = class(TForm)
    btnCreateGraphic: TButton;
    btnLoadSpreadsheet: TButton;
    editSourceFile: TFileNameEdit;
    Label1: TLabel;
    Label2: TLabel;
    editXAxis: TLabeledEdit;
    EditYAxis: TLabeledEdit;
    MyChart: TChart;
    FPSChartSource: TsWorksheetChartSource;
    MyChartLineSeries: TLineSeries;
    WorksheetGrid: TsWorksheetGrid;
    procedure btnCreateGraphicClick(Sender: TObject);
    procedure btnLoadSpreadsheetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  FPSChartForm: TFPSChartForm; 

implementation

uses
  // FPSpreadsheet and supported formats
  fpspreadsheet, xlsbiff8, xlsbiff5, xlsbiff2, xlsxooxml, fpsopendocument;

{$R *.lfm}

{ TFPSChartForm }

procedure TFPSChartForm.btnCreateGraphicClick(Sender: TObject);
begin
  FPSChartSource.LoadPropertiesFromStrings(editXAxis.Text, editYAxis.Text, '', '', '');
  FPSChartSource.LoadFromWorksheetGrid(WorksheetGrid);
end;

procedure TFPSChartForm.btnLoadSpreadsheetClick(Sender: TObject);
var
  Format: TsSpreadsheetFormat;
  lExt: string;
begin
  // First some logic to detect the format from the extension
  lExt := ExtractFileExt(editSourceFile.Text);
  if lExt = STR_EXCEL_EXTENSION then Format := sfExcel2
  else if lExt = STR_OOXML_EXCEL_EXTENSION then Format := sfOOXML
  else if lExt = STR_OPENDOCUMENT_CALC_EXTENSION then Format := sfOpenDocument
  else raise Exception.Create('Invalid File Extension');

  // Now the actual loading
  WorksheetGrid.LoadFromSpreadsheetFile(editSourceFile.Text, Format);
end;

procedure TFPSChartForm.FormCreate(Sender: TObject);
begin
  editSourceFile.InitialDir := ExtractFilePath(ParamStr(0));
end;

end.

