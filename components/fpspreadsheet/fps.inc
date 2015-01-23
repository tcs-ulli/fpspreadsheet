{------------------------------------------------------------------------------}
{                 Central definition of conditional defines                    }
{------------------------------------------------------------------------------}

{ In Unix systems, the unit clocale is automatically added to the uses clause
  of fpspreadsheet.pas. This unit sets up localization settings needed for
  locale-dependent number and date/time formats. However, this adds a dependence
  on the C library to the package.
  If this is not wanted, define FPS_DONT_USE_CLOCALE. }
{.$DEFINE FPS_DONT_USE_CLOCALE}


{ In older versions of fpspreadsheet, the formatting fields had belonged to the
  cell record. This has been given up to reduce memory consumption.
  For fpc >2.6, however, record helpers allow to get this feature back. In case
  of older compilers activate the define FPS_NO_RECORD_HELPERS. Note that worksheet
  methods can only be used to change cell formatting then. }
{.$DEFINE FPS_NO_RECORD_HELPERS}

{ The next defines activate code duplicated from new compiler versions in case
  an old compiler is used. }

{ Numberformats require an extended version of FormatDateTime (in SysUtils)
  which is not available before FPC 3.0. Define FPS_FORMATDATETIME if the
  compiler used is older. }
{$DEFINE FPS_FORMATDATETIME}

{ fpspreadsheet requires the function VarIsBool which was introduced by
  fpc 2.6.4. If an older FPC versions is used define FPS_VARISBOOL. Keep
  undefined for the current FPC version. }
{.$DEFINE FPS_VARISBOOL}

{ fpspreadsheet requires some functions from LazUTF8 which were
  introduced by Lazarus 1.2. If an older Lazarus version is used define
  FPS_LAZUTF8. Keep undefined for the current Lazarus version. }
{.$DEFINE FPS_LAZUTF8}


