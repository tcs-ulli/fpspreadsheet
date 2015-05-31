program spreadtestcli;

{$mode objfpc}
{$h+}

uses
  custapp, Classes, SysUtils, fpcunit,
  plaintestreport {results output to plain text},
  xmltestreport {used to get results into XML format},
  testregistry,
  testdbwriter {used to get results into db},

  {the actual tests}
  datetests, manualtests, stringtests, internaltests, testsutility, testutils,
  formattests, colortests, emptycelltests, insertdeletetests, errortests,
  numberstests, fonttests, formulatests, numformatparsertests, optiontests,
  virtualmodetests, dbexporttests, sortingtests, copytests, celltypetests,
  commenttests, enumeratortests, hyperlinktests, pagelayouttests;

const
  ShortOpts = 'ac:dhlpr:x';
  Longopts: Array[1..11] of String = (
    'all','comment:','db', 'database', 'help','list','revision:','revisionid:','suite:','plain','xml');
  Version = 'Version 1';

type
  { TTestRunner }
  TTestOutputFormat = (tDB, tXMLAdvanced, tPlainText);

  TTestRunner = Class(TCustomApplication)
  private
    FFormat: TTestOutputFormat;
    FDBResultsWriter: TDBResultsWriter;
    FPlainResultsWriter: TPlainResultsWriter;
    FXMLResultsWriter: TXMLResultsWriter;
    procedure WriteHelp;
  protected
    procedure DoRun ; Override;
    procedure doTestRun(aTest: TTest); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;


constructor TTestRunner.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFormat:=tPlainText;
  //FDBResultsWriter := TDBResultsWriter.Create; //done in procedures
  FPlainResultsWriter:=TPlainResultsWriter.Create(nil);
  //Don't write out timing info, makes it more difficult to run a diff.
  //If you want to use timing, use the XML or db output:
  {$if FPC_FULLVERSION>=20701}
  FPlainResultsWriter.SkipTiming:=true;
  {$endif}
  FXMLResultsWriter:=TXMLResultsWriter.Create(nil);
end;


destructor TTestRunner.Destroy;
begin
  //FDBResultsWriter.Free; //done in procedures
  FPlainResultsWriter.Free;
  FXMLResultsWriter.Free;
end;


procedure TTestRunner.doTestRun(aTest: TTest);
var
  RevisionID: string;
  testResult: TTestResult;
begin
  testResult := TTestResult.Create;
  try
    case FFormat of
      tDB:
      begin
        testResult.AddListener(FDBResultsWriter);

        RevisionID:=GetOptionValue('r','revisionid');
        if RevisionID='' then
          RevisionID:=GetOptionValue('revision');
        if RevisionID<>'' then
          FDBResultsWriter.RevisionID:=RevisionID;

        FDBResultsWriter.Comment:=GetOptionValue('c','comment');

				{
        // Depending on the application, you may want to add some fake test suite hierarchy
        // at the top of the test project.
        // Why? This makes it easier to avoid comparing apples to oranges when you have
        // various platforms, editions, configurations, database connectors etc of your program/test set.
        // Here, we demonstrate this with a Latin language edition of the code in its Enterprise edition:
        FDBResultsWriter.TestSuiteRoot.Add('Enterprise');
        FDBResultsWriter.TestSuiteRoot.Add('Latin');
				}

        {
        // Normally, we would edit the testdbwriter.ini file and select our db
        // where the tests are stored that way.... or omit any ini file and let it
        // fallback to a Firebird embedded database.
        // However, if needed, that can be overridden here:
        FDBResultsWriter.DatabaseType:=TDBW_POSTGRESQLCONN_NAME;
        FDBResultsWriter.DatabaseHostname:='dbserver';
        FDBResultsWriter.DatabaseName:='dbtests';
        FDBResultsWriter.DatabaseUser:='postgres';
        FDBResultsWriter.DatabasePassword:='password';
        FDBResultsWriter.DatabaseCharset:='UTF8';
        }
      end;
      tPlainText:
      begin
        testResult.AddListener(FPlainResultsWriter);
      end;
      tXMLAdvanced:
      begin
        testResult.AddListener(FXMLResultsWriter);
        // if filename='null', no console output is generated...
        //FXMLResultsWriter.FileName:='';
      end;
    end;
    aTest.Run(testResult);
    case FFormat of
      tDB: testResult.RemoveListener(FDBResultsWriter);
      tPlainText:
      begin
        // This actually generates the plain text output:
        FPlainResultsWriter.WriteResult(TestResult);
        testResult.RemoveListener(FPlainResultsWriter);
      end;
      tXMLAdvanced:
      begin
        // This actually generates the XML output:
        FXMLResultsWriter.WriteResult(TestResult);
        // You can use fcl-xml's xmlwrite.WriteXMLFile to write the results
        // to a stream or file...
        testResult.RemoveListener(FXMLResultsWriter);
      end;
    end;
  finally
    testResult.Free;
  end;
end;

procedure TTestRunner.WriteHelp;
begin
  writeln(Title);
  writeln(Version);
  writeln(ExeName+': console test runner for fpspreadsheet tests');
  writeln('Runs test set for fpspreadsheet and');
	writeln('- stores the results in a database, or');
	writeln('- outputs to screen');
  writeln('');
  writeln('Usage: ');
  writeln('-c <comment>, --comment=<comment>');
  writeln('   add comment to test run info.');
  writeln('   (if database output is used)');
  writeln('-d or --db or --database: run all tests, output to database');
  writeln('-l or --list to show a list of registered tests');
  writeln('-p or --plain: run all tests, output in plain text (default)');
  writeln('-r <id> --revision=<id>, --revisionid=<id>');
  writeln('   add revision id/application version ID to test run info.');
  writeln('   (if database output is used)');
  writeln('-x or --xml to run all tests and show the output in XML (new '
    +'DUnit style)');
  writeln('');
  writeln('--suite=MyTestSuiteName to run only the tests in a single test '
    +'suite class');
  writeln('Example: --suite=TSpreadWriteReadStringTests');
end;

procedure TTestRunner.DoRun;
var
  FoundTest: boolean;
  I : Integer;
  S : String;
begin
  S:=CheckOptions(ShortOpts,LongOpts);
  If (S<>'') then
  begin
    Writeln(StdErr,S);
    WriteHelp;
    halt(1);
  end;

  // Default to plain text output:
  FFormat:=tPlainText;

  if HasOption('d', 'database') or HasOption('db') then
    FFormat:=tDB;

  if HasOption('h', 'help') then
  begin
    WriteHelp;
    halt(0);
  end;

  if HasOption('l', 'list') then
  begin
    writeln(GetSuiteAsPlain(GetTestRegistry));
    halt(0);
  end;

  if HasOption('p', 'plain') then
    FFormat:=tPlainText;

  if HasOption('x', 'xml') then
    FFormat:=tXMLAdvanced;

  if HasOption('suite') then
  begin
    S := '';
    S:=GetOptionValue('suite');

    // For the db writer: recreate test objects so we get new runs each time
    FoundTest:=false;
    FDBResultsWriter:=TDBResultsWriter.Create;
    try
      if S = '' then
      begin
        writeln('Error');
        writeln('You have to specify a test(suite). Valid test suite names:');
        for I := 0 to GetTestRegistry.Tests.count - 1 do
          writeln(GetTestRegistry[i].TestName)
      end
      else
      begin
        for I := 0 to GetTestRegistry.Tests.count - 1 do
        begin
          if GetTestRegistry[i].TestName = S then
          begin
            doTestRun(GetTestRegistry[i]);
            FoundTest:=true;
          end;
        end;
        if not(FoundTest) then
        begin
          writeln('Error: the testsuite "',S,'" you specified does not exist.');
        end;
      end;
    finally
      FDBResultsWriter.Free;
    end;
  end
  else
  begin
    // No suite
    // For the db writer: recreate test objects so we get new runs each time
    FDBResultsWriter:=TDBResultsWriter.Create;
    try
      doTestRun(GetTestRegistry);
    finally
      FDBResultsWriter.Free;
    end;
  end;
  Terminate;
end;


var
  App: TTestRunner;
begin
  App := TTestRunner.Create(nil);
  App.Initialize;
  App.Title := 'spreadtestcli';
  App.Run;
  App.Free;
end.

