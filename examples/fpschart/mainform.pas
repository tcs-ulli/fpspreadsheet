unit mainform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Grids, fpspreadsheetchart, fpspreadsheetgrid, TAGraph, TASeries;

type
  
  { TFPSChartForm }

  TFPSChartForm = class(TForm)
    btnCreateGraphic: TButton;
    MyChart: TChart;
    FPSChartSource: TsWorksheetChartSource;
    MyChartLineSeries: TLineSeries;
    WorksheetGrid: TsWorksheetGrid;
    procedure btnCreateGraphicClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  FPSChartForm: TFPSChartForm; 

implementation

{$R *.lfm}

{ TFPSChartForm }

procedure TFPSChartForm.btnCreateGraphicClick(Sender: TObject);
begin
  FPSChartSource.LoadFromWorksheetGrid(WorksheetGrid);
end;

end.

