{ package-wide declaration of constants used by Excel which are not only needed
  by the biff readers and writers. }

unit xlsconst;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

const
  { Formula constants TokenID values }

  { Binary Operator Tokens 3.6}
  INT_EXCEL_TOKEN_TADD      = $03;
  INT_EXCEL_TOKEN_TSUB      = $04;
  INT_EXCEL_TOKEN_TMUL      = $05;
  INT_EXCEL_TOKEN_TDIV      = $06;
  INT_EXCEL_TOKEN_TPOWER    = $07; // Power Exponentiation ^
  INT_EXCEL_TOKEN_TCONCAT   = $08; // Concatenation &
  INT_EXCEL_TOKEN_TLT       = $09; // Less than <
  INT_EXCEL_TOKEN_TLE       = $0A; // Less than or equal <=
  INT_EXCEL_TOKEN_TEQ       = $0B; // Equal =
  INT_EXCEL_TOKEN_TGE       = $0C; // Greater than or equal >=
  INT_EXCEL_TOKEN_TGT       = $0D; // Greater than >
  INT_EXCEL_TOKEN_TNE       = $0E; // Not equal <>
  INT_EXCEL_TOKEN_TISECT    = $0F; // Cell range intersection
  INT_EXCEL_TOKEN_TLIST     = $10; // Cell range list
  INT_EXCEL_TOKEN_TRANGE    = $11; // Cell range
  INT_EXCEL_TOKEN_TUPLUS    = $12; // Unary plus  +
  INT_EXCEL_TOKEN_TUMINUS   = $13; // Unary minus +
  INT_EXCEL_TOKEN_TPERCENT  = $14; // Percent (%, divides operand by 100)
  INT_EXCEL_TOKEN_TPAREN    = $15; // Operator in parenthesis

  { Constant Operand Tokens, 3.8}
  INT_EXCEL_TOKEN_TMISSARG  = $16; //missing operand
  INT_EXCEL_TOKEN_TSTR      = $17; //string
  INT_EXCEL_TOKEN_TERR      = $1C; //error value
  INT_EXCEL_TOKEN_TBOOL     = $1D; //boolean
  INT_EXCEL_TOKEN_TINT      = $1E; //(unsigned) integer
  INT_EXCEL_TOKEN_TNUM      = $1F; //floating-point

  { Operand Tokens }
  // _R: reference; _V: value; _A: array
  INT_EXCEL_TOKEN_TREFR     = $24;
  INT_EXCEL_TOKEN_TREFV     = $44;
  INT_EXCEL_TOKEN_TREFA     = $64;
  INT_EXCEL_TOKEN_TAREA_R   = $25;
  INT_EXCEL_TOKEN_TAREA_V   = $45;
  INT_EXCEL_TOKEN_TAREA_A   = $65;
  INT_EXCEL_TOKEN_TREFN_R   = $2C;
  INT_EXCEL_TOKEN_TREFN_V   = $4C;
  INT_EXCEL_TOKEN_TREFN_A   = $6C;
  INT_EXCEL_TOKEN_TAREAN_R  = $2D;
  INT_EXCEL_TOKEN_TAREAN_V  = $4D;
  INT_EXCEL_TOKEN_TAREAN_A  = $6D;

  { Function Tokens }
  // _R: reference; _V: value; _A: array
  // Offset 0: token; offset 1: index to a built-in sheet function ( ➜ 3.111) )
  INT_EXCEL_TOKEN_FUNC_R    = $21;
  INT_EXCEL_TOKEN_FUNC_V    = $41;
  INT_EXCEL_TOKEN_FUNC_A    = $61;

  //VAR: variable number of arguments:
  INT_EXCEL_TOKEN_FUNCVAR_R = $22;
  INT_EXCEL_TOKEN_FUNCVAR_V = $42;
  INT_EXCEL_TOKEN_FUNCVAR_A = $62;

  { Special tokens }
  INT_EXCEL_TOKEN_TEXP      = $01;  // cell belongs to shared formula

  { Built-in/worksheet functions }
  INT_EXCEL_SHEET_FUNC_COUNT      = 0;
  INT_EXCEL_SHEET_FUNC_IF         = 1;
  INT_EXCEL_SHEET_FUNC_ISNA       = 2;
  INT_EXCEL_SHEET_FUNC_ISERROR    = 3;
  INT_EXCEL_SHEET_FUNC_SUM        = 4;
  INT_EXCEL_SHEET_FUNC_AVERAGE    = 5;
  INT_EXCEL_SHEET_FUNC_MIN        = 6;
  INT_EXCEL_SHEET_FUNC_MAX        = 7;
  INT_EXCEL_SHEET_FUNC_ROW        = 8;
  INT_EXCEL_SHEET_FUNC_COLUMN     = 9;
  INT_EXCEL_SHEET_FUNC_NA         = 10;
  INT_EXCEL_SHEET_FUNC_NPV        = 11;
  INT_EXCEL_SHEET_FUNC_STDEV      = 12;
  INT_EXCEL_SHEET_FUNC_DOLLAR     = 13;
  INT_EXCEL_SHEET_FUNC_FIXED      = 14;  // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_SIN        = 15;
  INT_EXCEL_SHEET_FUNC_COS        = 16;
  INT_EXCEL_SHEET_FUNC_TAN        = 17;
  INT_EXCEL_SHEET_FUNC_ATAN       = 18;
  INT_EXCEL_SHEET_FUNC_PI         = 19;
  INT_EXCEL_SHEET_FUNC_SQRT       = 20;
  INT_EXCEL_SHEET_FUNC_EXP        = 21;
  INT_EXCEL_SHEET_FUNC_LN         = 22;
  INT_EXCEL_SHEET_FUNC_LOG10      = 23;
  INT_EXCEL_SHEET_FUNC_ABS        = 24; // $18
  INT_EXCEL_SHEET_FUNC_INT        = 25;
  INT_EXCEL_SHEET_FUNC_SIGN       = 26;
  INT_EXCEL_SHEET_FUNC_ROUND      = 27; // $1B
  INT_EXCEL_SHEET_FUNC_LOOKUP     = 28;
  INT_EXCEL_SHEET_FUNC_INDEX      = 29;
  INT_EXCEL_SHEET_FUNC_REPT       = 30;
  INT_EXCEL_SHEET_FUNC_MID        = 31;
  INT_EXCEL_SHEET_FUNC_LEN        = 32;
  INT_EXCEL_SHEET_FUNC_VALUE      = 33;
  INT_EXCEL_SHEET_FUNC_TRUE       = 34;
  INT_EXCEL_SHEET_FUNC_FALSE      = 35;
  INT_EXCEL_SHEET_FUNC_AND        = 36;
  INT_EXCEL_SHEET_FUNC_OR         = 37;
  INT_EXCEL_SHEET_FUNC_NOT        = 38;
  INT_EXCEL_SHEET_FUNC_MOD        = 39;
  INT_EXCEL_SHEET_FUNC_DCOUNT     = 40;
  INT_EXCEL_SHEET_FUNC_DSUM       = 41;
  INT_EXCEL_SHEET_FUNC_DAVERAGE   = 42;
  INT_EXCEL_SHEET_FUNC_DMIN       = 43;
  INT_EXCEL_SHEET_FUNC_DMAX       = 44;
  INT_EXCEL_SHEET_FUNC_DSTDEV     = 45;
  INT_EXCEL_SHEET_FUNC_VAR        = 46;
  INT_EXCEL_SHEET_FUNC_DVAR       = 47;
  INT_EXCEL_SHEET_FUNC_TEXT       = 48;
  INT_EXCEL_SHEET_FUNC_LINEST     = 49;   // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_TREND      = 50;   // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_LOGEST     = 51;   // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_GROWTH     = 52;   // BIFF2 has different parameters

  INT_EXCEL_SHEET_FUNC_PV         = 56;
  INT_EXCEL_SHEET_FUNC_FV         = 57;
  INT_EXCEL_SHEET_FUNC_NPER       = 58;
  INT_EXCEL_SHEET_FUNC_PMT        = 59;
  INT_EXCEL_SHEET_FUNC_RATE       = 60;
  INT_EXCEL_SHEET_FUNC_MIRR       = 61;
  INT_EXCEL_SHEET_FUNC_IRR        = 62;
  INT_EXCEL_SHEET_FUNC_RAND       = 63;
  INT_EXCLE_SHEET_FUNC_MATCH      = 64;
  INT_EXCEL_SHEET_FUNC_DATE       = 65; // $41
  INT_EXCEL_SHEET_FUNC_TIME       = 66; // $42
  INT_EXCEL_SHEET_FUNC_DAY        = 67;
  INT_EXCEL_SHEET_FUNC_MONTH      = 68;
  INT_EXCEL_SHEET_FUNC_YEAR       = 69;
  INT_EXCEL_SHEET_FUNC_WEEKDAY    = 70;   // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_HOUR       = 71;
  INT_EXCEL_SHEET_FUNC_MINUTE     = 72;
  INT_EXCEL_SHEET_FUNC_SECOND     = 73;
  INT_EXCEL_SHEET_FUNC_NOW        = 74;
  INT_EXCEL_SHEET_FUNC_AREAS      = 75;
  INT_EXCEL_SHEET_FUNC_ROWS       = 76;
  INT_EXCEL_SHEET_FUNC_COLUMNS    = 77;
  INT_EXCEL_SHEET_FUNC_OFFSET     = 78;

  INT_EXCEL_SHEET_FUNC_SEARCH     = 82;
  INT_EXCEL_SHEET_FUNC_TRANSPOSE  = 83;

  INT_EXCEL_SHEET_FUNC_TYPE       = 86;

  INT_EXCEL_SHEET_FUNC_ATAN2      = 97;
  INT_EXCEL_SHEET_FUNC_ASIN       = 98;
  INT_EXCEL_SHEET_FUNC_ACOS       = 99;
  INT_EXCEL_SHEET_FUNC_CHOOSE     = 100;
  INT_EXCEL_SHEET_FUNC_HLOOKUP    = 101;    // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_VLOOKUP    = 102;    // BIFF2 has different parameters

  INT_EXCEL_SHEET_FUNC_ISREF      = 105;

  INT_EXCEL_SHEET_FUNC_LOG        = 109;

  INT_EXCEL_SHEET_FUNC_CHAR       = 111;
  INT_EXCEL_SHEET_FUNC_LOWER      = 112;
  INT_EXCEL_SHEET_FUNC_UPPER      = 113;
  INT_EXCEL_SHEET_FUNC_PROPER     = 114;
  INT_EXCEL_SHEET_FUNC_LEFT       = 115;
  INT_EXCEL_SHEET_FUNC_RIGHT      = 116;
  INT_EXCEL_SHEET_FUNC_EXACT      = 117;
  INT_EXCEL_SHEET_FUNC_TRIM       = 118;
  INT_EXCEL_SHEET_FUNC_REPLACE    = 119;
  INT_EXCEL_SHEET_FUNC_SUBSTITUTE = 120;
  INT_EXCEL_SHEET_FUNC_CODE       = 121;

  INT_EXCEL_SHEET_FUNC_FIND       = 124;
  INT_EXCEL_SHEET_FUNC_CELL       = 125;
  INT_EXCEL_SHEET_FUNC_ISERR      = 126;
  INT_EXCEL_SHEET_FUNC_ISTEXT     = 127;
  INT_EXCEL_SHEET_FUNC_ISNUMBER   = 128;
  INT_EXCEL_SHEET_FUNC_ISBLANK    = 129;
  INT_EXCEL_SHEET_FUNC_T          = 130;
  INT_EXCEL_SHEET_FUNC_N          = 131;

  INT_EXCEL_SHEET_FUNC_DATEVALUE  = 140;
  INT_EXCEL_SHEET_FUNC_TIMEVALUE  = 141;
  INT_EXCEL_SHEET_FUNC_SLD        = 142;
  INT_EXCEL_SHEET_FUNC_SYD        = 143;
  INT_EXCEL_SHEET_FUNC_DDB        = 144;

  INT_EXCEL_SHEET_FUNC_CLEAN      = 162;
  INT_EXCEL_SHEET_FUNC_MDETERM    = 163;
  INT_EXCEL_SHEET_FUNC_MINVERSE   = 164;
  INT_EXCEL_SHEET_FUNC_MMULT      = 165;

  INT_EXCEL_SHEET_FUNC_IPMT       = 167;
  INT_EXCEL_SHEET_FUNC_PPMT       = 168;
  INT_EXCEL_SHEET_FUNC_COUNTA     = 169;

  INT_EXCEL_SHEET_FUNC_PRODUCT    = 183;
  INT_EXCEL_SHEET_FUNC_FACT       = 184;

  INT_EXCEL_SHEET_FUNC_DPRODUCT   = 189;
  INT_EXCEL_SHEET_FUNC_ISNONTEXT  = 190;

  INT_EXCEL_SHEET_FUNC_STDEVP     = 193;
  INT_EXCEL_SHEET_FUNC_VARP       = 194;
  INT_EXCEL_SHEET_FUNC_DSTDEVP    = 195;
  INT_EXCEL_SHEET_FUNC_DVARP      = 196;
  INT_EXCEL_SHEET_FUNC_TRUNC      = 197;  // BIFF2 has different parameters
  INT_EXCEL_SHEET_FUNC_ISLOGICAL  = 198;
  INT_EXCEL_SHEET_FUNC_DCOUNTA    = 199;

  // No BIFF2 after 199

  INT_EXCEL_SHEET_FUNC_TODAY      = 221;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_MEDIAN     = 227;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_SINH       = 229;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_COSH       = 230;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_TANH       = 231;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ASINH      = 232;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ACOSH      = 233;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ATANH      = 234;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_INFO       = 244;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_AVEDEV     = 269;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_BETADIST   = 270;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_BETAINV    = 272;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_BINOMDIST  = 273;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_CHIDIST    = 274;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_CHIINV     = 275;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_EVEN       = 279;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_FLOOR      = 285;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_CEILING    = 288;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ODD        = 298;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_PERMUT     = 299;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_POISSON    = 300;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_SUMSQ      = 321;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_CONCATENATE= 336;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_POWER      = 337;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_RADIANS    = 342;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_DEGREES    = 343;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_SUMIF      = 345;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_COUNTIF    = 346;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_COUNTBLANK = 347;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_DATEDIF    = 351;  // not available in BIFF2

  INT_EXCEL_SHEET_FUNC_HYPERLINK  = 359;  // BIFF8 only

  { Control Tokens, Special Tokens }
//  01H tExp Matrix formula or shared formula
//  02H tTbl Multiple operation table
//  15H tParen Parentheses
//  18H tNlr Natural language reference (BIFF8)
  INT_EXCEL_TOKEN_TATTR = $19; // tAttr Special attribute
//  1AH tSheet Start of external sheet reference (BIFF2-BIFF4)
//  1BH tEndSheet End of external sheet reference (BIFF2-BIFF4)


implementation

end.

