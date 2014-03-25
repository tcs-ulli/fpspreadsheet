{ testdbwriter

  Copyright (C) 2012-2014 Reinier Olislagers

  This library is free software; you can redistribute it and/or modify it
  under the terms of the GNU Library General Public License as published by
  the Free Software Foundation; either version 2 of the License, or (at your
  option) any later version with the following modification:

  As a special exception, the copyright holders of this library give you
  permission to link this library with independent modules to produce an
  executable, regardless of the license terms of these independent modules,and
  to copy and distribute the resulting executable under terms of your choice,
  provided that you also meet, for each linked independent module, the terms
  and conditions of the license of that module. An independent module is a
  module which is not derived from or based on this library. If you modify
  this library, you may extend this exception to your version of the library,
  but you are not obligated to do so. If you do not wish to do so, delete this
  exception statement from your version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
  FITNESS FOR A PARTICULAR PURPOSE. See the GNU Library General Public License
  for more details.

  You should have received a copy of the GNU Library General Public License
  along with this library; if not, write to the Free Software Foundation,
  Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
}

{
  Purpose:
  This unit contains a database TestListener for use with the fpcUnit testing
  framework. It puts test results into a database, which can be useful for
  regression testing, reporting etc.

  To set the database to be used, you can use a testdbwriter.ini file, The
  format of this file is similar to that used in the FPC db test framework.
  Please see testdbwriter.ini.txt for an example.

  If you don't use testdbwriter.ini, the code will try to create and use a
  Firebird embedded database.

  You can also override the selection from your test application by setting
  the DB* properties.

  This code uses simple insert/select SQL statements so it should be
  compatible with most RDBMS. Varchar size is a factor though; 800 is chosen
  to avoid Firebird "key size too big for index" errors.

  The database must support autoincrement/autonumber fields or equivalents
  (eg. sequences/generators and triggers). Firebird embedded database support
  is included; if the database does not exist, it can be created. Please see
  the testdbwriter_firebird.sql file for Firebird database statements you can
  adapt to your own database system.
  Firebird notes: the index sizes in the DDL requires page_size 16384.

  A PostgreSQL SQL file is also included.

  No autocreate support for sqlite yet; patches welcome.

}
// todo: change flat testsuite into testsuitepath; add normal column testsuitename
// todo: add regressionflat modification to postgresql (using testsuiteflat instead of regular testsuite)
// todo: verify/add views in firebird db to postgres schema as well
// todo: performance optimalizations => triggers instead of ctes etc

// Enable to get logging/writelns in order to troubleshoot
// database problems/program correctness. Note: uses writeln; only for console applications.
{.$DEFINE DEBUGCONSOLE}


{$IF FPC_FULLVERSION<20601}
// FPC < 2.6.1 doesn't yet have LocalTimeToUniversal.
{$ERROR "This version of FPC does not support getting UTC time!"}
// Use UTC time or what the local clock tells us? For consistency, UTC is best:
// It will eliminate inconsistency between *nix hosts (mostly UTC clocks)
// and Windows hosts (mostly local clocks)
{$ENDIF}

unit testdbwriter;

{$mode objfpc}{$H+}
{$IFDEF MSWINDOWS}
{$r *.rc} //include SQL for Firebird embedded
{$ELSE}
// unfortunately, the installers on e.g. Linux do not have windres as a install dependency
// so we make do with the (possibly stale) .res file
{$r *.res} //include SQL for Firebird embedded
{$ENDIF}

interface
uses
  Classes
  ,SysUtils
  ,fpcUnit
  ,TestUtils
  ,sqldb
  ,db
  { Even if we use sqlconnector we need to add the relevant db unit to the uses clause in the project.
  In order to make it simpler for the tester to share programs across db systems, we choose to add the units
  here; if needed, other units can be added by the test writer in his test program. }
  ,mysql40conn, mysql41conn, mysql50conn, mysql51conn, mysql55conn
  ,ibconnection
  ,mssqlconn
  ,odbcconn
  {$IFNDEF WIN64}
  ,oracleconnection
  {$ENDIF}
  ,pqconnection
  ,sqlite3conn
  ,inifiles
  ,ibase60dyn
  ,dateutils
  ,resource;

const
  TDBW_INVALID_DB_ID=0; //Used as a marker for an invalid database ID
  TDBW_DBVERSION='20121202'; //Update this whenever the database schema changes

  // Some known database connector names to be used for FDBType; the
  // values are the connectiondef name for the sqldb connector
  TDBW_FIREBIRDCONN_NAME='Firebird';
  TDBW_MSSQLCONN_NAME='MSSQLServer';
  TDBW_ODBCCONN_NAME='ODBC';
  TDBW_ORACLECONN_NAME='Oracle';
  TDBW_POSTGRESQLCONN_NAME='PostGreSQL';
  TDBW_SQLITECONN_NAME='SQLite3';
  TDBW_SYBASECONN_NAME='Sybase';

type

  { TDBResultsWriter }
  TDBResultsWriter = class(TNoRefCountObject, ITestListener)
  protected
    FApplicationName: string;
    FComment: string; //User-specified comment for test run
    FConn: TSQLConnector;
    FCPU: string; //Processor used (name uses FPC terminology)
    FDBProfile: string; //Name of the profile in the ini file the user chose. Only useful in error messages to user
    FDBSetup: boolean; //Track if database is set up including adding test run data
    FDBType: string; //Output database type chosen by user
    FDBHostname: string; //Host/IP address of database server. Empty string for embedded dbs
    FDBName: string; //Name/path for database
    FDBUser: string; //Username
    FDBPassword: string; //Password needed for username
    FDBCharset: string; //Character encoding for db connection
    FOS: string; //Operating system used (name uses FPC terminology)
    FRevisionID: string; //Unique identifier for source code being tested
    FRootSuite: TStringList; //This list of testsuite names will be prepended to the test suite hierarchy when tests are run
    FTestSuitePath: TStringList; //Stack of suite names (and IDs) up to the current test suite. Will include FRootSuite
    FTran: TSQLTransaction;

    // Prepared queries for performance as well as one general purpose query:
    FQuery: TSQLQuery; //General purpose query
    FAddTestRun: TSQLQuery;
    FGetApplicationName: TSQLQuery;
    FGetCPU: TSQLQuery;
    FGetExceptionMessage: TSQLQuery;
    FGetExceptionClass: TSQLQuery;
    FGetMethodName: TSQLQuery;
    FGetOS: TSQLQuery;
    FGetResultValue: TSQLQuery;
    FGetSourceLocation: TSQLQuery;
    FGetSourceUnit: TSQLQuery;
    FGetTestSuite: TSQLQuery;
    FGetTestSuiteParent: TSQLQuery;
    FGetTestName: TSQLQuery;
    FGetTestResultID: TSQLQuery;
    FInsertExceptionMessage: TSQLQuery;
    FInsertSourceLocation: TSQLQuery;
    FInsertTest: TSQLQuery;
    FInsertTestResult: TSQLQuery;
    FInsertTestSuiteParent: TSQLQuery;
    FUpdateTestResult: TSQLQuery;
    FUpdateTestResultException: TSQLQuery;
    FUpdateTestResultExceptionError: TSQLQuery;
    FUpdateTestResultOK: TSQLQuery;

    FResultErrorID: integer; //Cached result value ID
    FResultFailedID: integer; //Cached result value ID
    FResultIgnoredID: integer; //Cached result value ID
    FResultOKID: integer; //Cached result value ID
    FStartRun: TDateTime; //Start of running entire test run, in UTC if possible. Used to identify test run record.
    FStartTest: TDateTime; //Start of running test suite
    FTestResultID: integer; //DB identifier of current test
    FTestRunID: integer; //DB identifier of the test run - the list of tests we ran with their results and messages
    FTestResultOK: boolean; //Test framework doesn't seem to track succesful tests?!?! So we need to do it ourselves
    FTestSuiteID: integer; //DB identifier of current test
    { Looks up PK based on the parameters given; if not present, inserts using InsertSQL}
    function AddOrFindKey(GetQuery: TSQLQuery; InsertSQL: string; ParamValue: String): integer; virtual;
    { Creates a new database. Called automatically; only implemented for Firebird embedded/sqlite }
    function CreateDatabase: boolean; virtual;
    { Create database tables. Called automatically; only implemented for Firebird embedded }
    procedure CreateTables; virtual;
    { Sets up database connection, if necessary create embedded db, and set up queries }
    procedure DatabaseSetup; virtual;
    { Looks up exception class and message in db; if necessary adds them.
    Returns EXCEPTIONMESSAGES.ID }
    function DBGetExceptionMessageID(ExceptionClass, ExceptionMessage: string): integer; virtual;
    { Looks up method name in db; if necessary adds it.}
    function DBGetMethodNameID(Name: string): integer; virtual;
    { Looks up result value ID in db; if necessary adds it }
    function DBGetResultValueID(ResultText: string): integer; virtual;
    { Looks up source location ID; if necessary adds data }
    function DBGetSourceLocationID(SourceUnit: string; Line: integer):integer; virtual;
    { Looks up test suite name/test name combination in db; if necessary adds name}
    function DBGetTestID(TestName: string): integer; virtual;
    { Looks up test suite in db that has no parent; if necessary adds it}
    function DBGetTestSuiteID(TestSuiteName: string): integer;
    { Looks up test suite with parent in db; if necessary adds it with given hierarchical depth}
    function DBGetTestSuiteID(TestSuiteName: string; Depth, ParentSuiteID: integer): integer;
    { Generate a unique ID, used to identify newly inserted records in autonumber tables }
    function GetUniqueID: string; virtual;
    function GetTestRunID: string; virtual;
    {$IFDEF DEBUGCONSOLE}
    procedure GetLogEvent(Sender: TSQLConnection;
        EventType: TDBEventType; const Msg: String);
    {$ENDIF DEBUGCONSOLE}
    { Insert test run info at the beginning of a test run. Needs db setup; called when processing first test suite. }
    procedure InsertTestRun; virtual;
    { Reads testdbwriter.ini to get database connection details. Defaults to firebird embedded if no testdbwriter.ini found}
    procedure ReadSettings; virtual;
    { Configures an auxiliary Firebird connection for Firebird-specific functionality }
    procedure SetupIBConnection(var IBConn: TIBConnection); virtual;
    { Helper method to set up a class level query object }
    procedure SetupQuery(var Query: TSQLQuery; SQL: string);
  public
    { Name of the application that is running the tests. Optional: if nothing chosen, the current paramstr(0) will be used }
    property ApplicationName: string write FApplicationName;
    { Comment regarding the test run }
    property Comment: string write FComment;
    { CPU on which the tests are run. Note: default already set by testdbwriter based on compiler defines}
    property CPU: string read FCPU write FCPU;
    { Character set/encoding used by test storage database. Optional, overrides values set by ini file (if any) }
    property DatabaseCharset: string read FDBCharset write FDBCharset;
    { Host ip address/name for test storage database. Leave blank for embedded. Optional, overrides values set by ini file (if any)}
    property DatabaseHostname: string read FDBHostname write FDBHostname;
    { Name of the test storage database. Optional, overrides values set by ini file (if any)}
    property DatabaseName: string read FDBName write FDBName;
    { Connector type for test storage database. Use connector name. Optional, overrides values set by ini file (if any).
      If no type ends up being selected, a Firebird embedded database will be created/used }
    property DatabaseType: string read FDBType write FDBType;
    { Password for test storage database. Optional, overrides values set by ini file (if any)}
    property DatabasePassword: string read FDBPassword write FDBPassword;
    { Username for test storage database. Optional, overrides values set by ini file (if any)}
    property DatabaseUser: string read FDBUser write FDBUser;
    { Operating system on which the tests are run. Note: default already set by testdbwriter based on compiler defines }
    property OS: string read FOS write FOS;
    { Unique identifier for the revision/version for the source code under test }
    property RevisionID: string write FRevisionID;
    { Unique ID for test run in database }
    property TestRunID: string read GetTestRunID;
    { You can use this to specify a root hierarchy of testsuite names. The test suites will be added below this hierarchy.
    This is useful e.g. for later comparing the same tests against different environments/backends/databases/configurations }
    property TestSuiteRoot: TStringList read FRootSuite;
    constructor Create;
    destructor Destroy; override;

    { ITestListener interface requirements }
    { Called when a test results in a failure. Updates test record with failure info. }
    procedure AddFailure(ATest: TTest; AFailure: TTestFailure);
    { Called when a test results in an error. Updates test record with error info. }
    procedure AddError(ATest: TTest; AError: TTestFailure);
    { Called at the beginning of a test within a testsuite. Inserts new test record. }
    procedure StartTest(ATest: TTest);
    { Called at the end of a test within a testsuite. Updates test record with elapsed time etc. }
    procedure EndTest(ATest: TTest);
    { Called at the beginning of a testsuite. Inserts testsuite name if needed.}
    procedure StartTestSuite(ATestSuite: TTestSuite);
    { Called at the end of a testsuite. Virtually empty procedure here.}
    procedure EndTestSuite(ATestSuite: TTestSuite);
  end;


implementation

const
  RESULT_ERROR='Error'; //unknown if this should be used
  RESULT_FAILED='Failed'; //Used in RESULTVALUES table
  RESULT_IGNORED='Ignored'; //Used in RESULTVALUES table
  RESULT_OK='OK'; //Used in RESULTVALUES table

{ TDBResultsWriter }

function TDBResultsWriter.CreateDatabase: boolean;
var
  IBConn:TIBConnection;
begin
  result:=false;
  // For now, only Firebird embedded. SQLite creates missing dbs automatically.
  if (FConn.ConnectorType=TDBW_FIREBIRDCONN_NAME) and
    (FConn.HostName='') and
    (not FileExists(FDBName)) then
  begin
    if FConn.DatabaseName='' then Exception.Create('No database name specified. This is required for embedded databases.');
    IBConn:=TIBConnection.Create(nil);
    try
      SetupIBConnection(IBConn);
      {$IFDEF DEBUGCONSOLE}
      writeln('Creating Firebird db:');
      writeln('ibconn.username: '+ibconn.UserName);
      writeln('ibconn.dbname: '+ibconn.DatabaseName);
      writeln('ibconn.dialect: '+inttostr(ibconn.Dialect));
      {$ENDIF DEBUGCONSOLE}
      IBConn.CreateDB;
    finally
      IBConn.Free;
    end;
    result:=true;
  end;
end;


function TDBResultsWriter.AddOrFindKey(GetQuery: TSQLQuery; InsertSQL:string; ParamValue: string): integer;
begin
  result:=TDBW_INVALID_DB_ID;
  GetQuery.Params[0].AsString:=ParamValue;
  try
    GetQuery.Open;
    if GetQuery.RecordCount>0 then
      result:=GetQuery.Fields[0].AsInteger;
    GetQuery.Close;
  except
    //ignore
  end;
  if result=TDBW_INVALID_DB_ID then
  begin
    // Doesn't exist, so add and retrieve PK
    try
      FQuery.SQL.Text:=InsertSQL;
      FQuery.Params[0].AsString:=ParamValue;
      FQuery.ExecSQL;
    except
      //ignore duplicate key failure or violations due to NOT NULL fields
    end;
    GetQuery.Params[0].AsString:=ParamValue;
    try
      GetQuery.Open;
      if GetQuery.RecordCount>0 then
        result:=GetQuery.Fields[0].AsInteger;
      GetQuery.Close;
    except
      result:=TDBW_INVALID_DB_ID;
    end;
  end;
end;

procedure TDBResultsWriter.SetupQuery(var Query: TSQLQuery; SQL: string);
begin
  Query:=TSQLQuery.Create(nil);
  Query.SQL.Text:=SQL;
  Query.DataBase:=FConn;
end;

procedure TDBResultsWriter.CreateTables;
// Creates required tables in database.
// Uses include file generated from Flamerobin DDL dump
var
  FBScript:TSQLScript;
  ScriptResource: TResourceStream;
  ScriptText: TStringList;
  IBConn:TIBConnection;
  IBTran:TSQLTransaction;
  TranWasStarted: boolean;
  dummypchar:pchar;
begin
  // We currently only support firebird sql
  if (FConn.ConnectorType=TDBW_FIREBIRDCONN_NAME) then
  begin
    TranWasStarted:=FTran.Active;
    if TranWasStarted then
    begin
      // We want FTran to see our DDL later...
      FTran.Commit;
    end;
    FBScript:=TSQLScript.Create(nil);
    IBConn:=TIBConnection.Create(nil);
    IBTran:=TSQLTransaction.Create(nil);
    dummypchar:=(PChar(RT_RCDATA));
    ScriptResource:=TResourceStream.Create(hInstance,'SQLSCRIPT',PChar(RT_RCDATA));
    ScriptText:=TStringList.Create;
    try
      SetupIBConnection(IBConn);
      IBConn.Transaction:=IBTran;
      ScriptText.LoadFromStream(ScriptResource);
      FBScript.UseCommit:=true; //needed for Firebird DDL+DML statements
      FBScript.UseSetTerm:=true; //needed for Firebird SP definitions
      FBScript.DataBase:=IBConn;
      FBScript.Transaction:=IBTran;
      FBScript.Script:=ScriptText;
      IBTran.StartTransaction;
      try
        FBScript.Execute;
        //todo: investigate: fb is half created; no triggers; procedure recalc indexes exists but is empty
      except
        on E: EIBDatabaseError do
        begin
          // Only tell the user something is up in debug mode:
          {$IFDEF DEBUGCONSOLE}
          writeln('TestDBWriter: CreateTables: error running script: GDS error '+inttostr(E.GDSErrorCode)+'('+E.Message+')');
          {$ENDIF DEBUGCONSOLE}
          halt(16);
        end;
        on F: Exception do
        begin
          // Only tell the user something is up in debug mode:
          {$IFDEF DEBUGCONSOLE}
          writeln('TestDBWriter: CreateTables: error running script: '+F.ClassName+'/'+F.Message);
          {$ENDIF DEBUGCONSOLE}
          halt(17);
        end;
      end;

      IBTran.Commit;
    finally
      FBScript.Free;
      IBConn.Free;
      IBTran.Free;
      ScriptText.Free;
      ScriptResource.Free;
    end;
    FConn.Transaction:=FTran; //apparently this info gets lost, perhaps only on FPC 2.6.0!??
    FTran.StartTransaction;
    //Indicate database schema version, useful for upgrade scripts
    FConn.ExecuteDirect('INSERT INTO OPTIONS (OPTIONNAME, OPTIONVALUE,REMARKS) VALUES (''DBVERSION'','''+TDBW_DBVERSION+''',''Database schema version'')');
    FTran.Commit;
    if TranWasStarted then
    begin
      FTran.StartTransaction;
    end;
  end
  else
  begin
    raise Exception.CreateFmt('No support for autocreate tables for database type %s', [FDBType]);
  end;
end;

procedure TDBResultsWriter.DatabaseSetup;
var
  CreatedDB: boolean;
begin
  FDBSetup:=true; //try setup only once. If it fails, doing it again won't help.

  // Translate some connector names as a courtesy:
  FConn:=TSQLConnector.Create(nil);
  case lowercase(FDBType) of
    'interbase', 'firebird', 'firebirdembedded': FConn.ConnectorType:=
      TDBW_FIREBIRDCONN_NAME;
    'mssql', 'sql', 'sqlserver': FConn.ConnectorType:=TDBW_MSSQLCONN_NAME;
    'odbc': FConn.ConnectorType:=TDBW_ODBCCONN_NAME;
    'oracle': FConn.ConnectorType:=TDBW_ORACLECONN_NAME;
    'pq', 'postgres', 'postgresql': FConn.ConnectorType:=TDBW_POSTGRESQLCONN_NAME;
    'sqlite', 'sqlite3': FConn.ConnectorType:=TDBW_SQLITECONN_NAME;
    'sybase', 'sybasease': FConn.ConnectorType:=TDBW_SYBASECONN_NAME;
  else FConn.ConnectorType:=FDBType;
  end;

  // Make sure Firebird embedded gets loaded (on old FPC versions):
  if (FConn.ConnectorType=TDBW_FIREBIRDCONN_NAME) and (FDBHostname='') then
    UseEmbeddedFirebird:=true;

  {$IFDEF DEBUGCONSOLE}
  FConn.LogEvents:=LogAllEvents;
  FConn.OnLog:=@Self.GetLogEvent;
  {$ENDIF DEBUGCONSOLE}
  FConn.HostName:=FDBHostname;
  FConn.DatabaseName:=FDBName;
  FConn.UserName:=FDBUser;
  FConn.Password:=FDBPassword;
  FConn.Charset:=FDBCharset;

  FTran:=TSQLTransaction.Create(nil);
  FConn.Transaction:=FTran;

  // Check for db/tables and create if necessary
  // (only suppported on some systems):
  CreatedDB:=CreateDatabase;

  FConn.Open;
  if CreatedDB then CreateTables;

  //Apparently this info gets lost by the createtable calls!?
  FConn.Transaction:=FTran;
  FTran.StartTransaction;

  SetupQuery(FQuery, ''); //general purpose query

  SetupQuery(FGetApplicationName,
    'SELECT a.ID '+
    'FROM APPLICATIONS a '+
    'WHERE '+
    'a.NAME = :NAME ');

  SetupQuery(FGetCPU,
    'SELECT a.ID '+
    'FROM CPU a '+
    'WHERE '+
    'a.CPUNAME = :CPUNAME ');

  SetupQuery(FGetExceptionMessage,
    'SELECT a.ID ' +
    'FROM EXCEPTIONMESSAGES a ' +
    'WHERE ' +
    'a.EXCEPTIONCLASS = :EXCEPTIONCLASS AND  ' +
    'a.EXCEPTIONMESSAGE = :EXCEPTIONMESSAGE ');

  SetupQuery(FGetExceptionClass,
    'SELECT a.ID ' +
    'FROM EXCEPTIONCLASSES a  ' +
    'WHERE ' +
    'a.EXCEPTIONCLASS = :EXCEPTIONCLASS ');

  SetupQuery(FGetMethodName,
    'SELECT a.ID ' +
    'FROM METHODNAMES a  ' +
    'WHERE ' +
    'a.NAME = :NAME ');

  SetupQuery(FGetOS,
    'SELECT a.ID '+
    'FROM OS a '+
    'WHERE '+
    'a.OSNAME = :OSNAME ');

  SetupQuery(FGetResultValue,
    'SELECT a.ID ' +
    'FROM RESULTVALUES a ' +
    'WHERE ' +
    'a.NAME = :NAME ');

  SetupQuery(FGetSourceLocation,
    'SELECT a.ID ' +
    'FROM SOURCELOCATIONS a ' +
    'WHERE ' +
    'a.SOURCEUNIT = :SOURCEUNIT AND ' +
    'a.LINE = :LINE ');

  SetupQuery(FGetSourceUnit,
    'SELECT a.ID ' +
    'FROM SOURCEUNITS a ' +
    'WHERE ' +
    'a.NAME = :NAME ');

  // Gets only top-level test suites
  SetupQuery(FGetTestSuite,
    'SELECT a.ID ' +
    'FROM TESTSUITES a ' +
    'WHERE ' +
    'a.NAME = :NAME AND ' +
    'a.PARENTSUITE IS NULL ');

  // Gets test suites with specified parent suite
  SetupQuery(FGetTestSuiteParent,
    'SELECT a.ID ' +
    'FROM TESTSUITES a ' +
    'WHERE ' +
    'a.PARENTSUITE=:PARENTSUITE AND '+
    'a.NAME = :NAME ');

  SetupQuery(FGetTestName,
    'SELECT a.ID ' +
    'FROM TESTS a  ' +
    'WHERE ' +
    'a.TESTSUITE = :TESTSUITE AND  ' +
    'a.NAME = :NAME ');

  SetupQuery(FInsertExceptionMessage,
    'INSERT INTO EXCEPTIONMESSAGES (EXCEPTIONCLASS,EXCEPTIONMESSAGE) VALUES (:'
      +'EXCEPTIONCLASS,:EXCEPTIONMESSAGE) ');

  SetupQuery(FInsertSourceLocation,
    'INSERT INTO SOURCELOCATIONS (SOURCEUNIT,LINE) VALUES (:SOURCEUNIT,:LINE) '
      );

  SetupQuery(FGetTestResultID,
    'SELECT ID FROM TESTRESULTS WHERE TESTRUN=:TESTRUN AND TEST=:TEST '+
    'AND RESULTCOMMENT=:RESULTCOMMENT ');

  SetupQuery(FInsertTest,
    'INSERT INTO TESTS (TESTSUITE,NAME) VALUES (:TESTSUITE,:NAME) ');

  SetupQuery(FInsertTestResult,
    'INSERT INTO TESTRESULTS (TESTRUN,TEST,RESULTCOMMENT) VALUES (:TESTRUN,:TEST,:RESULTCOMMENT) ');

  SetupQuery(FInsertTestSuiteParent,
    'INSERT INTO TESTSUITES (PARENTSUITE,NAME,DEPTH) VALUES (:PARENTSUITE,:NAME,:DEPTH) ');

  SetupQuery(FUpdateTestResult,
    'UPDATE TESTRESULTS SET ELAPSEDTIME=:ELAPSEDTIME, RESULTCOMMENT=:RESULTCOMMENT '+
    'WHERE ID=:ID ');

  SetupQuery(FUpdateTestResultException,
    'UPDATE TESTRESULTS SET RESULTVALUE=:RESULTVALUE, ' +
    'EXCEPTIONMESSAGE=:EXCEPTIONMESSAGE WHERE ID=:ID ');

  SetupQuery(FUpdateTestResultExceptionError,
    'UPDATE TESTRESULTS SET RESULTVALUE=:RESULTVALUE, '+
    'EXCEPTIONMESSAGE=:EXCEPTIONMESSAGE, METHODNAME=:METHODNAME, '+
    'SOURCELOCATION=:SOURCELOCATION '+
    'WHERE ID=:ID ');

  SetupQuery(FUpdateTestResultOK,
    'UPDATE TESTRESULTS SET ELAPSEDTIME=:ELAPSEDTIME, '+
    'RESULTVALUE=:RESULTVALUE, RESULTCOMMENT=:RESULTCOMMENT WHERE ID=:ID ');
end;


function TDBResultsWriter.DBGetExceptionMessageID(ExceptionClass,
  ExceptionMessage: string): integer;
var
  ID: integer;
begin
  result:=TDBW_INVALID_DB_ID;
  ID:=AddOrFindKey(FGetExceptionClass,
    'INSERT INTO EXCEPTIONCLASSES (EXCEPTIONCLASS) VALUES (:EXCEPTIONCLASS) ',
    ExceptionClass);
  if ID<>TDBW_INVALID_DB_ID then
  begin
    // We can't call AddOrFindKey due to integer parameters, so do it manually:
    result:=TDBW_INVALID_DB_ID;
    FGetExceptionMessage.Params[0].AsInteger:=ID;
    FGetExceptionMessage.Params[1].AsString:=ExceptionMessage;
    try
      FGetExceptionMessage.Open;
      if FGetExceptionMessage.recordcount>0 then
        result:=FGetExceptionMessage.Fields[0].AsInteger;
      FGetExceptionMessage.Close;
    except
      result:=TDBW_INVALID_DB_ID;
    end;
    if result=TDBW_INVALID_DB_ID then
    begin
      // Doesn't exist yet
      try
        FInsertExceptionMessage.Params[0].AsInteger:=ID;
        FInsertExceptionMessage.Params[1].AsString:=ExceptionMessage;
        FInsertExceptionMessage.ExecSQL;
      except
        //ignore duplicate key failure
      end;

      FGetExceptionMessage.Params[0].AsInteger:=ID;
      FGetExceptionMessage.Params[1].AsString:=ExceptionMessage;
      try
        FGetExceptionMessage.Open;
        if FGetExceptionMessage.recordcount>0 then
          result:=FGetExceptionMessage.Fields[0].AsInteger;
        FGetExceptionMessage.Close;
      except
        result:=TDBW_INVALID_DB_ID;
      end;
    end;
  end;
end;

function TDBResultsWriter.DBGetMethodNameID(Name: string): integer;
var
  ID: integer;
begin
  result:=TDBW_INVALID_DB_ID;
  ID:=AddOrFindKey(FGetMethodName,
    'INSERT INTO METHODNAMES (NAME) VALUES (:NAME) ',
    Name);
  result:=ID;
end;

function TDBResultsWriter.DBGetResultValueID(ResultText: string): integer;
begin
  result:=TDBW_INVALID_DB_ID;
  case ResultText of
  RESULT_ERROR:
    begin
      if FResultErrorID=TDBW_INVALID_DB_ID then
      begin
        FResultErrorID:=AddOrFindKey(FGetResultValue,
          'INSERT INTO RESULTVALUES (NAME) VALUES (:NAME) ',
          ResultText);
      end;
    result:=FResultErrorID;
    end;
  RESULT_FAILED:
    begin
      if FResultFailedID=TDBW_INVALID_DB_ID then
      begin
        FResultFailedID:=AddOrFindKey(FGetResultValue,
          'INSERT INTO RESULTVALUES (NAME) VALUES (:NAME) ',
          ResultText);
      end;
    result:=FResultFailedID;
    end;
  RESULT_IGNORED:
    begin
      if FResultIgnoredID=TDBW_INVALID_DB_ID then
      begin
        FResultIgnoredID:=AddOrFindKey(FGetResultValue,
          'INSERT INTO RESULTVALUES (NAME) VALUES (:NAME) ',
          ResultText);
      end;
    result:=FResultIgnoredID;
    end;
  RESULT_OK:
    begin
      if FResultOKID=TDBW_INVALID_DB_ID then
      begin
        FResultOKID:=AddOrFindKey(FGetResultValue,
          'INSERT INTO RESULTVALUES (NAME) VALUES (:NAME) ',
          ResultText);
      end;
      result:=FResultOKID;
    end
  else
    // Not cached, but we can still look it up
    result:=AddOrFindKey(FGetResultValue,
      'INSERT INTO RESULTVALUES (NAME) VALUES (:NAME) ',
      ResultText);
  end;
end;

function TDBResultsWriter.DBGetSourceLocationID(SourceUnit: string;
  Line: integer): integer;
var
  ID: integer;
begin
  result:=TDBW_INVALID_DB_ID;
  ID:=AddOrFindKey(FGetSourceUnit,
    'INSERT INTO SOURCEUNITS (NAME) VALUES (:SOURCEUNIT) ',
    SourceUnit);
  if ID<>TDBW_INVALID_DB_ID then
  begin
    // We can't call AddOrFindKey due to integer parameters, so do it manually:
    result:=TDBW_INVALID_DB_ID;
    FGetSourceLocation.Params[0].AsInteger:=ID;
    FGetSourceLocation.Params[1].AsInteger:=Line;
    try
      FGetSourceLocation.Open;
      if FGetSourceLocation.recordcount>0 then
        result:=FGetSourceLocation.Fields[0].AsInteger;
      FGetSourceLocation.Close;
    except
      result:=TDBW_INVALID_DB_ID;
    end;
    if result=TDBW_INVALID_DB_ID then
    begin
      // Doesn't exist, so add it
      try
        FInsertSourceLocation.Params[0].AsInteger:=ID;
        FInsertSourceLocation.Params[1].AsInteger:=Line;
        FInsertSourceLocation.ExecSQL;
      except
        //ignore duplicate key failure
      end;

      FGetSourceLocation.Params[0].AsInteger:=ID;
      FGetSourceLocation.Params[1].AsInteger:=Line;
      try
        FGetSourceLocation.Open;
        if FGetSourceLocation.recordcount>0 then
          result:=FGetSourceLocation.Fields[0].AsInteger;
        FGetSourceLocation.Close;
      except
        result:=TDBW_INVALID_DB_ID;
      end;
    end;
  end;
end;

function TDBResultsWriter.DBGetTestID(TestName: string): integer;
// We already have FTestSuiteID
begin
  result:=TDBW_INVALID_DB_ID;
  assert(FTestSuiteID<>TDBW_INVALID_DB_ID,'Test suite ID must be valid.');
  // We can't call AddOrFindKey due to integer parameters, so do it manually:
  result:=TDBW_INVALID_DB_ID;
  {$IFDEF DEBUGCONSOLE}
  writeln('debug: Getting test id for test '+TestName+' in test suite id: '+inttostr(FTestSuiteID));
  {$ENDIF DEBUGCONSOLE}
  FGetTestName.Params[0].AsInteger:=FTestSuiteID;
  FGetTestName.Params[1].AsString:=TestName;
  try
    FGetTestName.Open;
    if FGetTestName.RecordCount>0 then
      result:=FGetTestName.Fields[0].AsInteger;
    FGetTestName.Close;
  except
    result:=TDBW_INVALID_DB_ID;
  end;

  if result=TDBW_INVALID_DB_ID then
  begin
    // Doesn't exist, so add it
    try
      FInsertTest.Params[0].AsInteger:=FTestSuiteID;
      FInsertTest.Params[1].AsString:=TestName;
      FInsertTest.ExecSQL;
    except
      //ignore duplicate key failure
    end;

    FGetTestName.Params[0].AsInteger:=FTestSuiteID;
    FGetTestName.Params[1].AsString:=TestName;
    try
      FGetTestName.Open;
      if FGetTestName.RecordCount>0 then
        result:=FGetTestName.Fields[0].AsInteger;
      FGetTestName.Close;
    except
      result:=TDBW_INVALID_DB_ID;
    end;
  end;
end;

function TDBResultsWriter.DBGetTestSuiteID(TestSuiteName: string): integer;
// Looks up top-level test suite (parentsuite is null)
var
  ID: integer;
begin
  result:=TDBW_INVALID_DB_ID;
  ID:=AddOrFindKey(FGetTestSuite,
    'INSERT INTO TESTSUITES (PARENTSUITE, NAME, DEPTH) VALUES (NULL, :NAME, 0) ',
    TestSuiteName);
  result:=ID;
end;

function TDBResultsWriter.DBGetTestSuiteID(TestSuiteName: string; Depth, ParentSuiteID: integer): integer;
// Looks up test suite (with given depth) that is a chiled of ParentSuiteID, returns test suite ID
begin
  result:=TDBW_INVALID_DB_ID;
  assert(ParentSuiteID<>TDBW_INVALID_DB_ID,'Test suite ID must be valid.');
  if ParentSuiteID<>TDBW_INVALID_DB_ID then
  begin
    // We can't call AddOrFindKey due to integer parameters, so do it manually:
    result:=TDBW_INVALID_DB_ID;
    FGetTestSuiteParent.Params[0].AsInteger:=ParentSuiteID;
    FGetTestSuiteParent.Params[1].AsString:=TestSuiteName;
    try
      FGetTestSuiteParent.Open;
      if FGetTestSuiteParent.RecordCount>0 then
        result:=FGetTestSuiteParent.Fields[0].AsInteger;
      FGetTestSuiteParent.Close;
    except
      result:=TDBW_INVALID_DB_ID;
    end;

    if result=TDBW_INVALID_DB_ID then
    begin
      // Doesn't exist, so add it
      try
        FInsertTestSuiteParent.Params[0].AsInteger:=ParentSuiteID;
        FInsertTestSuiteParent.Params[1].AsString:=TestSuiteName;
        FInsertTestSuiteParent.Params[2].AsInteger:=Depth;
        FInsertTestSuiteParent.ExecSQL;
      except
        //ignore duplicate key failure
      end;

      FGetTestSuiteParent.Params[0].AsInteger:=ParentSuiteID;
      FGetTestSuiteParent.Params[1].AsString:=TestSuiteName;
      try
        FGetTestSuiteParent.Open;
        if FGetTestSuiteParent.RecordCount>0 then
          result:=FGetTestSuiteParent.Fields[0].AsInteger;
        FGetTestSuiteParent.Close;
      except
        result:=TDBW_INVALID_DB_ID;
      end;
    end;
  end;
end;

function TDBResultsWriter.GetUniqueID: string;
begin
  Result:='In progress; ID: '+FormatDateTime('ssz',Now())+'/'+inttostr(random(32767));
end;

function TDBResultsWriter.GetTestRunID: string;
// Return a string in case we ever use GUIDs etc.
begin
  result:=inttostr(FTestRunID);
end;

{$IFDEF DEBUGCONSOLE}
procedure TDBResultsWriter.GetLogEvent(Sender: TSQLConnection;
  EventType: TDBEventType; const Msg: String);
  // The procedure is called by TSQLConnection and saves the received log messages
  // in the FConnectionLog stringlist
var
  Source: string;
begin
  // Nicely right aligned...
  Source:='';
  case EventType of
    detCustom:   Source:='Custom:  ';
    detExecute:  Source:='Execute: ';
    detCommit:   Source:='Commit:  ';
    detRollBack: Source:='Rollback:';
  end;
  if Source<>'' then
  begin
    writeln(Source + ' ' + Msg);
  end;
end;
{$ENDIF DEBUGCONSOLE}

procedure TDBResultsWriter.InsertTestRun;
var
  ApplicationID: integer; //Applications.ID needed for FK value in TestRuns.ApplicationID
  CPUID: integer; //CPU.ID needed for FK value in TestRuns.CPU
  i: integer;
  OSID: integer; //OS.ID needed for FK value in TestRuns.OS
  RandomID: string; //For increasing record uniqueness in db
  TestSuiteID: integer; //ID of current testsuite name
begin
  // We use this to insert a new run at the beginning of the test run
  // We identify the record by the DATETIMERAN field, as well as a temporary
  // random value in RUNCOMMENT (multiple writers could be updating the db
  // at the same time)
  RandomID:=GetUniqueID;
  // When the run is done, we update the other fields.
  if FApplicationName<>'' then
    ApplicationID:=AddOrFindKey(FGetApplicationName,
      'INSERT INTO APPLICATIONS (NAME) VALUES (:NAME) ',
      FApplicationName);
  if FCPU<>'' then
    CPUID:=AddOrFindKey(FGetCPU,
      'INSERT INTO CPU (CPUNAME) VALUES (:CPUNAME) ',
      FCPU);
  if FOS<>'' then
    OSID:=AddOrFindKey(FGetOS,
      'INSERT INTO OS (OSNAME) VALUES (:OSNAME) ',
      FOS);
  SetupQuery(FAddTestRun,
    'INSERT INTO TESTRUNS (DATETIMERAN,APPLICATIONID,RUNCOMMENT,CPU,OS) ' +
    'VALUES (:DATETIMERAN,:APPLICATIONID,:RUNCOMMENT,:CPU,:OS) ');
  // Now set up test run:
  FAddTestRun.Params[0].AsDateTime:=FStartRun;
  if FApplicationName='' then
  begin
    FAddTestRun.Params[1].Clear;
    FAddTestRun.Params[1].DataType:=ftInteger;
  end
  else
  begin
    FAddTestRun.Params[1].AsInteger:=ApplicationID;
  end;
  FAddTestRun.Params[2].AsString:=RandomID; //temporarily, in comment field
  if FCPU='' then
  begin
    FAddTestRun.Params[3].Clear;
    FAddTestRun.Params[3].DataType:=ftInteger;
  end
  else
  begin
    FAddTestRun.Params[3].AsInteger:=CPUID;
  end;
  if FOS='' then
  begin
    FAddTestRun.Params[4].Clear;
    FAddTestRun.Params[4].DataType:=ftInteger;
  end
  else
  begin
    FAddTestRun.Params[4].AsInteger:=OSID;
  end;
  FAddTestRun.ExecSQL;

  FTestRunID:=TDBW_INVALID_DB_ID;
  FQuery.SQL.Text:='SELECT ID FROM TESTRUNS '+
    'WHERE DATETIMERAN=:DATETIMERAN AND RUNCOMMENT=:RUNCOMMENT';
  FQuery.Params[0].AsDateTime:=FStartRun;
  FQuery.Params[1].AsString:=RandomID;
  FQuery.Open;
  if FQuery.RecordCount>0 then
    FTestRunID:=FQuery.Fields[0].AsInteger;
  FQuery.Close;
  assert(FTestRunID<>TDBW_INVALID_DB_ID, 'Constructor: test run ID must have been '
    +'created.');

  // If testsuite prefixes are specified, first clean them up a bit.
  // This only needs to be done once, so running it here makes sense.
  if FRootSuite.Count>0 then
  begin
    // Don't allow empty names:
    for i:=FRootSuite.Count-1 downto 0 do
    begin
      if FRootSuite[i]='' then
        FRootSuite.Delete(i);
    end;
  end;

  // Add all non-empty testsuite prefixes/names to the testsuite stack
  // that will be used in further processing
  // and look up their IDs
  if FRootSuite.Count>0 then
  begin
    TestSuiteID:=TDBW_INVALID_DB_ID;
    for i:=0 to FRootSuite.Count-1 do
    begin
      if TestSuiteID=TDBW_INVALID_DB_ID then
        TestSuiteID:=DBGetTestSuiteID(FRootSuite[i])
      else
        TestSuiteID:=DBGetTestSuiteID(FRootSuite[i],i,TestSuiteID);
      assert(TestSuiteID<>TDBW_INVALID_DB_ID,'InsertTestRun: TestSuiteID for '+FRootSuite[i]+' must not be invalid. Aborting.');
      {$IFDEF DEBUGCONSOLE}
      // Assertions won't stop executions but apparently restart the test?
      if TestSuiteID=TDBW_INVALID_DB_ID then
      begin
        writeln('Halt: InsertTestRun: TestSuiteID for '+FRootSuite[i]+' must not be invalid. Aborting.');
        halt(17);
      end;
      //todo: debug
      writeln('InsertTestRun: Adding test suite'+FRootSuite[i]+' with id '+inttostr(TestSuiteID)+' to FTestSuitePath');
      {$ENDIF}
      FTestSuitePath.AddObject(FRootSuite[i],TObject(PtrInt(TestSuiteID)));
    end;
  end;
end;

procedure TDBResultsWriter.ReadSettings;
var
  IniFile : TIniFile;
begin
  //Defaults:
  FDBType:=TDBW_FIREBIRDCONN_NAME;
  FDBHostname:='';
  FDBName:='testdbwriter.fdb';
  FDBUser:='SYSDBA';
  FDBPassword:='masterkey';
  FDBCharset:='UTF8';
  // Read ini file in program directory if possible
  IniFile := TIniFile.Create(ExtractFilePath(ParamStr(0))+'testdbwriter.ini');
  try
    FDBProfile:=IniFile.ReadString('Database','profile','firebirdembedded');
    FDBType:=LowerCase(IniFile.ReadString(FDBProfile,'type','firebird')); //Normalize to lowercase
    FDBHostname := IniFile.ReadString(FDBProfile,'hostname','');
    FDBName := IniFile.ReadString(FDBProfile,'name','testdbwriter.fdb');
    FDBUser := IniFile.ReadString(FDBProfile,'user','SYSDBA');
    FDBPassword := IniFile.ReadString(FDBProfile,'password','masterkey');
    FDBCharset := IniFile.ReadString(FDBProfile,'charset','UTF8');
  finally
    IniFile.Free;
  end;
end;

procedure TDBResultsWriter.SetupIBConnection(var IBConn: TIBConnection);
begin
  IBConn.UserName:=FConn.UserName;
  IBConn.Password:=FConn.Password;
  IBConn.DatabaseName:=FConn.DatabaseName;
  IBConn.CharSet:=FConn.Charset;
  IBConn.Params.Add('PAGE_SIZE=16384'); //important for max index size=>max column size
end;


constructor TDBResultsWriter.Create;
begin
  FRootSuite:=TStringList.Create;
  FTestSuitePath:=TStringList.Create;
  // Let's include setup in our run timings... only fair.
  FStartRun:=Now();
  FStartRun:=LocalTimeToUniversal(FStartRun);
  Randomize;

  // Invalidate cached values
  FResultErrorID:=TDBW_INVALID_DB_ID;
  FResultFailedID:=TDBW_INVALID_DB_ID;
  FResultIgnoredID:=TDBW_INVALID_DB_ID;
  FResultOKID:=TDBW_INVALID_DB_ID;

  FTestRunID:=TDBW_INVALID_DB_ID;
  FTestResultID:=TDBW_INVALID_DB_ID;
  FTestSuiteID:=TDBW_INVALID_DB_ID;

  FDBSetup:=false; //Flag to notify test run code to set up the db if required

  // Strip out extension so you get the same result on Windows, *nix etc:
  FApplicationName:=ChangeFileExt(ExtractFileName(Paramstr(0)),'');
  FCPU:={$I %FPCTARGETCPU%};
  FOS:={$I %FPCTARGETOS%};

  ReadSettings; //Read INI file settings (which can be overridden by DB* properties)
end;


destructor TDBResultsWriter.Destroy;
var
  ElapsedTime:TDateTime;
  EndTime: TDateTime;
begin
  EndTime:=Now();
  EndTime:=LocalTimeToUniversal(EndTime);
  ElapsedTime:=EndTime-FStartRun;

  // Update final test run data - but don't stop the destructor on errors
  if FDBSetup then
  begin
    try
      Assert(FTestRunID<>TDBW_INVALID_DB_ID,'Test run id must be valid due to assignment in constructor.');
      FQuery.SQL.Text:='UPDATE TESTRUNS SET '+
        'TOTALELAPSEDTIME=:TOTALELAPSEDTIME, RUNCOMMENT=:RUNCOMMENT, REVISIONID=:REVISIONID '+
        'WHERE ID=:ID ';
      FQuery.Params[0].AsDateTime:=ElapsedTime;
      if FComment='' then
      begin
        FQuery.Params[1].Clear; //null
        FQuery.Params[1].Datatype:=ftstring;
      end
      else
        FQuery.Params[1].AsString:=FComment;
      if FRevisionID='' then
      begin
        FQuery.Params[2].Clear; //null
        FQuery.Params[2].Datatype:=ftstring;
      end
      else
        FQuery.Params[2].AsString:=FRevisionID;
      FQuery.Params[3].AsInteger:=FTestRunID;
      FQuery.ExecSQL;
      FTran.Commit;

      FAddTestRun.Free;
      FGetApplicationName.Free;
      FGetCPU.Free;
      FGetExceptionMessage.Free;
      FGetExceptionClass.Free;
      FGetMethodName.Free;
      FGetOS.Free;
      FGetResultValue.Free;
      FGetSourceLocation.Free;
      FGetSourceUnit.Free;
      FGetTestSuite.Free;
      FGetTestSuiteParent.Free;
      FGetTestName.Free;
      FGetTestResultID.Free;
      FInsertExceptionMessage.Free;
      FInsertSourceLocation.Free;
      FInsertTest.Free;
      FInsertTestResult.Free;
      FInsertTestSuiteParent.Free;
      FUpdateTestResult.Free;
      FUpdateTestResultException.Free;
      FUpdateTestResultExceptionError.Free;
      FUpdateTestResultOK.Free;
      FTran.Free;
      FConn.Close;
      FConn.Free;
    except
      on E: Exception do begin
      // We need to swallow the exception because we're closing down.
      // Only tell the user something is up in debug mode:
      {$IFDEF DEBUGCONSOLE}
      writeln('TestDBWriter: ignored problem in destructor: '+E.ClassName+'/'+E.Message);
      {$ENDIF DEBUGCONSOLE}
      end;
    end;
  end;

  FRootSuite.Free;
  FTestSuitePath.Free;

  inherited Destroy;
end;


procedure TDBResultsWriter.AddFailure(ATest: TTest; AFailure: TTestFailure);
var
  MessageID: integer;
  ResultID: integer;
begin
  assert(FTestResultID<>TDBW_INVALID_DB_ID,'AddFailure: test id must be assigned.');
  FTestResultOK:=false;
  MessageID:=TDBW_INVALID_DB_ID;
  if (AFailure.ExceptionClassName<>'') or (AFailure.ExceptionMessage<>'') then
    MessageID:=DBGetExceptionMessageID(AFailure.ExceptionClassName,AFailure.ExceptionMessage);
  if AFailure.IsIgnoredTest then
  begin
    // Ignored
    ResultID:=DBGetResultValueID(RESULT_IGNORED);
    FUpdateTestResultException.Params[0].AsInteger:=ResultID;
    FUpdateTestResultException.Params[1].AsInteger:=MessageID;
    FUpdateTestResultException.Params[2].AsInteger:=FTestResultID;
    FUpdateTestResultException.ExecSQL;
  end
  else
  begin
    // Really failed
    ResultID:=DBGetResultValueID(RESULT_FAILED);
    FUpdateTestResultException.Params[0].AsInteger:=ResultID;
    if MessageID=TDBW_INVALID_DB_ID then
      FUpdateTestResultException.Params[1].Clear //null
    else
      FUpdateTestResultException.Params[1].AsInteger:=MessageID;
    FUpdateTestResultException.Params[2].AsInteger:=FTestResultID;
    FUpdateTestResultException.ExecSQL;
  end;
end;


procedure TDBResultsWriter.AddError(ATest: TTest; AError: TTestFailure);
var
  MethodNameID: integer;
  MessageID: integer;
  ResultID: integer;
  SourceLocationID: integer;
begin
  assert(FTestResultID<>TDBW_INVALID_DB_ID,'AddError: test id must be assigned.');
  FTestResultOK:=false;
  MessageID:=TDBW_INVALID_DB_ID;
  MethodNameID:=TDBW_INVALID_DB_ID;
  SourceLocationID:=TDBW_INVALID_DB_ID;

  if (AError.ExceptionClassName<>'') or (AError.ExceptionMessage<>'') then
    MessageID:=DBGetExceptionMessageID(AError.ExceptionClassName,AError.ExceptionMessage);
  if (AError.SourceUnitName<>'') and (AError.LineNumber<>0) then
    SourceLocationID:=DBGetSourceLocationID(AError.SourceUnitName,AError.LineNumber);
  if AError.FailedMethodName<>'' then
    MethodNameID:=DBGetMethodNameID(AError.FailedMethodName);

  ResultID:=DBGetResultValueID(RESULT_ERROR);
  FUpdateTestResultExceptionError.Params[0].AsInteger:=ResultID;
  if MessageID=TDBW_INVALID_DB_ID then
  begin
    FUpdateTestResultExceptionError.Params[1].Clear;
    FUpdateTestResultExceptionError.Params[1].DataType:=ftInteger;
  end
  else
    FUpdateTestResultExceptionError.Params[1].AsInteger:=MessageID;
  if MethodNameID=TDBW_INVALID_DB_ID then
  begin
    FUpdateTestResultExceptionError.Params[2].Clear;
    FUpdateTestResultExceptionError.Params[2].DataType:=ftInteger;
  end
  else
    FUpdateTestResultExceptionError.Params[2].AsInteger:=MethodNameID;
  if SourceLocationID=TDBW_INVALID_DB_ID then
  begin
    FUpdateTestResultExceptionError.Params[3].Clear;
    FUpdateTestResultExceptionError.Params[3].DataType:=ftInteger;
  end
  else
    FUpdateTestResultExceptionError.Params[3].AsInteger:=SourceLocationID;
  FUpdateTestResultExceptionError.Params[4].AsInteger:=FTestResultID;
  FUpdateTestResultExceptionError.ExecSQL;
end;


procedure TDBResultsWriter.StartTest(ATest: TTest);
var
  RandomID: string;
  TestID: integer;
begin
  // Empty etc test suite names may lead to invalid FTestSuiteIDs which should be ignored
  if FTestSuiteID<>TDBW_INVALID_DB_ID then
  begin
    {$IFDEF DEBUGCONSOLE}
    // Assertions won't stop executions but apparently restart the test?
    if FTestRunID=TDBW_INVALID_DB_ID then
    begin
      writeln('Halt: Valid FTestRunID must be retrieved from db before using.');
      halt(16);
    end;
    {$ENDIF}
    assert(FTestRunID<>TDBW_INVALID_DB_ID,'Valid testrun ID must be retrieved from db before using.');
    FTestResultOK:=true; //Keep track of whether test succeeded or not
    TestID:=TDBW_INVALID_DB_ID; //Notify db code we need to get a new one
    TestID:=DBGetTestID(ATest.TestName);
    assert(TestID<>TDBW_INVALID_DB_ID,'Valid test ID must be retrieved from db before using.');
    // If doing e.g. performance tests, the same test name may be run multiple
    // times per run, so keep track of what we inserted:
    RandomID:=GetUniqueID;

    try
      FInsertTestResult.Params[0].AsInteger:=FTestRunID;
      FInsertTestResult.Params[1].AsInteger:=TestID;
      FInsertTestResult.Params[2].AsString:=RandomID;
      FInsertTestResult.ExecSQL;

      FGetTestResultID.Params[0].AsInteger:=FTestRunID;
      FGetTestResultID.Params[1].AsInteger:=TestID;
      FGetTestResultID.Params[2].AsString:=RandomID;
      FGetTestResultID.Open;
      if FGetTestResultID.RecordCount>0 then
        FTestResultID:=FGetTestResultID.Fields[0].AsInteger;
      FGetTestResultID.Close;
    except
      on E: Exception do
      begin
        {$IFDEF DEBUGCONSOLE}
        writeln('testdbwriter: error inserting test start data.');
        writeln('Exception: '+E.ClassName+'/'+E.Message);
        writeln('Aborting.');
        {$ENDIF}
        halt(15);
      end;
    end;

    FStartTest:=Now();
    FStartTest:=LocalTimeToUniversal(FStartTest);
  end;
end;


procedure TDBResultsWriter.EndTest(ATest: TTest);
var
  ElapsedTime: TDateTime;
  EndTime: TDateTime;
  ResultID: integer;
begin
  EndTime:=Now();
  EndTime:=LocalTimeToUniversal(EndTime);
  ElapsedTime:=EndTime-FStartTest;

  // Update test record depending on result
  Assert(FTestResultID<>TDBW_INVALID_DB_ID,'We must have a valid test ID from the StartTest code.');
  if FTestResultOK then
  begin
    ResultID:=DBGetResultValueID(RESULT_OK);
    Assert(ResultID<>TDBW_INVALID_DB_ID,'Result ID for OK must be available.');
    FUpdateTestResultOK.Params[0].AsDateTime:=ElapsedTime;
    FUpdateTestResultOK.Params[1].AsInteger:=ResultID;
    FUpdateTestResultOK.Params[2].Clear; //we used this as temp ID
    FUpdateTestResultOK.Params[2].DataType:=ftString;
    FUpdateTestResultOK.Params[3].AsInteger:=FTestResultID;
    FUpdateTestResultOK.ExecSQL;
  end
  else
  begin
    FUpdateTestResult.Params[0].AsDateTime:=ElapsedTime;
    FUpdateTestResult.Params[1].Clear; //we used this as temp ID
    FUpdateTestResult.Params[1].DataType:=ftString;
    FUpdateTestResult.Params[2].AsInteger:=FTestResultID;
    FUpdateTestResult.ExecSQL;
  end;
  FTestResultID:=TDBW_INVALID_DB_ID; //Notify db code we need to add a new one
end;

procedure TDBResultsWriter.StartTestSuite(ATestSuite: TTestSuite);
begin
  if not FDBSetup then
  begin
    {$IFDEF DEBUGCONSOLE}
    writeln('*** debug: dbsetup begin');
    {$ENDIF DEBUGCONSOLE}
    // FDBSetup indicates first time a test suite is entered. Not only set up
    // db, but also insert test run information:

    // If something went wrong here we should just quite; it makes no sense
    // to run test results without getting them into the db.
    // Apparently the test suite ignores exceptions, so explicit testing here
    try
      // Wrong username/password etc will be picked up here:
      DatabaseSetup;
    except
      on E: Exception do
      begin
        {$IFDEF DEBUGCONSOLE}
        writeln('testdbwriter: error setting up db connection.');
        writeln('(Are your connection details correct?)');
        writeln('Exception: '+E.ClassName+'/'+E.Message);
        writeln('Aborting.');
        {$ENDIF}
        halt(13);
      end;
    end;
    try
      // This is the first time we try to insert data, so permissions errors,
      // missing table problems etc can occur here:
      InsertTestRun;
    except
      on E: Exception do
      begin
        {$IFDEF DEBUGCONSOLE}
        writeln('testdbwriter: error inserting test run info.');
        writeln('(Does the db have the right tables and are permissions OK?)');
        writeln('Exception: '+E.ClassName+'/'+E.Message);
        writeln('Aborting.');
        {$ENDIF}
        halt(14);
      end;
    end;
    {$IFDEF DEBUGCONSOLE}
    writeln('*** debug: dbsetup end');
    {$ENDIF}
  end;

  {$IFDEF DEBUGCONSOLE}
  writeln('*** debug: test suite begin');
  {$ENDIF}
  //todo: debug
  // Only honor non-empty test suite names. This means that
  // TDBW_INVALID_DB_ID may be passed on to other procedures
  // where it should be ignored.
  FTestSuiteID:=TDBW_INVALID_DB_ID;
  if ATestSuite.TestName<>'' then
  begin
    // Retrieve test suite based on parent - if any:
    if (FTestSuitePath.Count>0) then
      FTestSuiteID:=ptrint(FTestSuitePath.Objects[FTestSuitePath.Count-1]);
    if FTestSuiteID=TDBW_INVALID_DB_ID then
      FTestSuiteID:=DBGetTestSuiteID(ATestSuite.TestName)
    else
      FTestSuiteID:=DBGetTestSuiteID(ATestSuite.TestName,FTestSuitePath.Count,FTestSuiteID); //Depth is 0 based
    assert(FTestSuiteID<>TDBW_INVALID_DB_ID,'StartTestSuite: FTestSuiteID for testsuite '+ATestSuite.TestName+' must not be invalid.');
    FTestSuitePath.AddObject(ATestSuite.TestName,TObject(ptrint(FTestSuiteID)));
    {$IFDEF DEBUGCONSOLE}
    writeln('StartTestSuite: Added test suite'+ATestSuite.TestName+' with id '+inttostr(FTestSuiteID)+' to FTestSuitePath, now count: '+inttostr(FTestSuitePath.Count));
    {$ENDIF}
    if (FTestSuiteID=TDBW_INVALID_DB_ID) then
    begin
      writeln('Halt: StartTestSuite: FTestSuiteID for testsuite '+ATestSuite.TestName+' must not be invalid.');
      halt(18);
    end;
  end;
  {$IFDEF DEBUGCONSOLE}
  writeln('*** debug: test suite end');
  writeln('');
  {$ENDIF}
end;


procedure TDBResultsWriter.EndTestSuite(ATestSuite: TTestSuite);
begin
  // Skip back up one item in the stack:
  if FTestSuitePath.Count>0 then
    FTestSuitePath.Delete(FTestSuitePath.Count-1);
  FTestSuiteID:=TDBW_INVALID_DB_ID;
end;


end.

