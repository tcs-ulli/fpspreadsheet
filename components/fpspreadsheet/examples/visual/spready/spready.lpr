program spready;

{$mode objfpc}{$H+}

uses
  Interfaces, // this includes the LCL widgetset
  Forms, mainform,
  scsvparamsform, sformatsettingsform, ssortparamsform, sctrls, scurrencyform;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainFrm, MainFrm);
  MainFrm.BeforeRun;
  Application.Run;
end.

