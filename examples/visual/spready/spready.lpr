program spready;

{$mode objfpc}{$H+}

uses
  Interfaces, // this includes the LCL widgetset
  Forms, mainform, laz_fpspreadsheet_visual, sCtrls, fpsCurrency,
  scsvparamsform, sfcurrencyform, sformatsettingsform, ssortparamsform;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainFrm, MainFrm);
  MainFrm.BeforeRun;
  Application.Run;
end.

