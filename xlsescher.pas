{ xlsEscher }

{@@ ----------------------------------------------------------------------------
  The unit xlsExcel provides basic support for the hierarchy of shapes and
  drawings in Microsoft Office files ("Escher", "OfficeArt") as it is needed for
  the BIFF record MSODRAWING (Cell comments, charts).

  AUTHORS: Werner Pamler

  DOCUMENTATION:
    Office Drawing 97-2007 Binary Format Specification
      http://www.digitalpreservation.gov/formats/digformatspecs/OfficeDrawing97-2007BinaryFormatSpecification.pdf
    [MS-ODRAW].pdf
      https://msdn.microsoft.com/en-us/library/office/cc441433%28v=office.12%29.aspx
    [MS-PPT].pdf
      https://msdn.microsoft.com/en-us/library/office/cc313106%28v=office.12%29.aspx
    [MS-XLS].pdf
      https://msdn.microsoft.com/en-us/library/office/cc313154%28v=office.12%29.aspx

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}

unit xlsEscher;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  { Record IDs }
  MSO_ID_OFFICEART_DG_CONTAINER      = $F002;
  MSO_ID_OFFICEART_SPGR_CONTAINER    = $F003;
  MSO_ID_OFFICEART_SP_CONTAINER      = $F004;
  MSO_ID_OFFICEART_FDG               = $F008;
  MSO_ID_OFFICEART_FSPGR             = $F009;
  MSO_ID_OFFICEART_FSP               = $F00A;
  MSO_ID_OFFICEART_FOPT              = $F00B;
  MSO_ID_OFFICEART_CLIENTTEXTBOX     = $F00D;
  MSO_ID_OFFICEART_CLIENTANCHORSHEET = $F010;
  MSO_ID_OFFICEART_CLIENTDATA        = $F011;

  { Record version }
  MSO_VER_CONTAINER = $0F;

  { Shape types }
  MSO_SPT_MIN = 0;
  MSO_SPT_NOTPRIMITIVE = MSO_SPT_MIN;
  MSO_SPT_RECTANGLE = 1;
  MSO_SPT_ROUNDRECTANGLE = 2;
  MSO_SPT_ELLIPSE = 3;
  MSO_SPT_DIAMOND = 4;
  MSO_SPT_ISOCELESTRIANGLE = 5;
  MSO_SPT_RIGHT_TRIANGLE = 6;
  MSO_SPT_PARALLELOGRAM = 7;
  MSO_SPT_TRAPEZOID = 8;
  MSO_SPT_HEXAGON = 9;
  MSO_SPT_OCTAGON = 10;
  MSO_SPT_PLUS = 11;
  MSO_SPT_STAR = 12;
  MSO_SPT_ARROW = 13;
  MSO_SPT_THICKARROW = 14;
  MSO_SPT_HOMEPLAT = 15;
  MSO_SPT_CUBE = 16;
  MSO_SPT_BALLOON = 17;
  MSO_SPT_SEAL = 18;
  MSO_SPT_ARC = 19;
  MSO_SPT_LINE = 20;
  MSO_SPT_PLAQUE = 21;
  MSO_SPT_CAN = 22;
  MSO_SPT_DONUT = 23;
  MSO_SPT_TEXTSIMPLE = 24;
  MSO_SPT_TEXTOCTAGON = 25;
  MSO_SPT_TEXTHEXAGON = 26;
  MSO_SPT_TEXTCURVE = 27;
  MSO_SPT_TEXTWAVE = 28;
  MSO_SPT_TEXTRING = 29;
  MSO_SPT_TEXTONCURVE = 30;
  MSO_SPT_TEXTONRING = 31;
  MSO_SPT_STRAIGHTCONNECTOR1 = 32;
  MSO_SPT_BENTCONNECTOR2 = 33;
  MSO_SPT_BENTCONNECTOR3 = 34;
  MSO_SPT_BENTCONNECTOR4 = 35;
  MSO_SPT_BENTCONNECTOR5 = 36;
  MSO_SPT_CURVEDCONNECTOR2 = 37;
  MSO_SPT_CURVEDCONNECTOR3 = 38;
  MSO_SPT_CURVEDCONNECTOR4 = 39;
  MSO_SPT_CURVEDCONNECTOR5 = 40;
  MSO_SPT_CALLOUT1 = 41;
  MSO_SPT_CALLOUT2 = 42;
  MSO_SPT_CALLOUT3 = 43;
  MSO_SPT_ACCENTCALLOUT1 = 44;
  MSO_SPT_ACCENTCALLOUT2 = 45;
  MSO_SPT_ACCENTCALLOUT3 = 46;
  MSO_SPT_BORDERCALLOUT1 = 47;
  MSO_SPT_BORDERCALLOUT2 = 48;
  MSO_SPT_BORDERCALLOUT3 = 49;
  MSO_SPT_ACCENTBORDERCALLOUT1 = 50;
  MSO_SPT_ACCENTBORDERCALLOUT2 = 51;
  MSO_SPT_ACCENTBORDERCALLOUT3 = 52;
  MSO_SPT_RIBBON = 53;
  MSO_SPT_RIBBON2 = 54;
  MSO_SPT_CHEVRON = 55;
  MSO_SPT_PENTAGON = 56;
  MSO_SPT_NOSMOKING = 57;
  MSO_SPT_SEAL8 = 58;
  MSO_SPT_SEAL16 = 59;
  MSO_SPT_SEAL32 = 60;
  MSO_SPT_WEDGERECTCALLOUT = 61;
  MSO_SPT_WEDGERRECTCALLOUT = 62;
  MSO_SPT_WEDGEELLIPSECALLOUT = 63;
  MSO_SPT_WAVE = 64;
  MSO_SPT_FOLDERCORNER = 65;
  MSO_SPT_LEFTARROW = 66;
  MSO_SPT_DOWNARROW = 67;
  MSO_SPT_UPARROW = 68;
  MSO_SPT_LEFTRIGHTARROW = 69;
  MSO_SPT_UPDOWNARROW = 70;
  MSO_SPT_IRREGULARSEAL1 = 71;
  MSO_SPT_IRREGULARSEAL2 = 72;
  MSO_SPT_LIGNTNINGBOLT = 73;
  MSO_SPT_HEART = 74;
  MSO_SPT_PICTUREFRAME = 75;
  MSO_SPT_QUADARROW = 76;
  MSO_SPT_LEFTARROWCALLOUT = 77;
  MSO_SPT_RIGHTARROWCALLOUT = 78;
  MSO_SPT_UPARROWCALLOUT = 79;
  MSO_SPT_DOWNARROWCALLOUT = 80;
  MSO_SPT_LEFTRIGHTARROWCALLOUT = 81;
  MSO_SPT_UPDOWNARROWCALLOUT = 82;
  MSO_SPT_QUADARROWCALLOUT = 83;
  MSO_SPT_BEVEL = 84;
  MSO_SPT_LEFTBRACKET = 85;
  MSO_SPT_RIGHTBRACKET = 86;
  MSO_SPT_LEFTBRACE = 87;
  MSO_SPT_RIGHTBRACE = 88;
  MSO_SPT_LEFTUPARROW = 89;
  MSO_SPT_BENTUPARROW = 90;
  MSO_SPT_BENTARROW = 91;
  MSO_SPT_SEAL25 = 92;
  MSO_SPT_STRIPEDRIGHTARROW = 83;
  MSO_SPT_NOTCHEDRIGHTARROW = 84;
  MSO_SPT_BLOCKARC = 95;
  MSO_SPT_SMILIEYFACE = 96;
  MSO_SPT_VERTICALSCROLL = 97;
  MSO_SPT_HORIZONTALSCROLL = 98;
  MSO_SPT_CICRULARARROW = 99;
  MSO_SPT_NOTCHEDCIRCULARARROW = 100;
  MSO_SPT_UTURNARROW = 101;
  MSO_SPT_CURVEDRIGHTARROW = 102;
  MSO_SPT_CURVEDLEFTARROW = 103;
  MSO_SPT_CURVEDUPARROW = 104;
  MSO_SPT_CURVEDDOWNARROW = 105;
  MSO_SPT_CLOUDCALLOUT = 106;
  MSO_SPT_ELLIPSERIBBON = 107;
  MSO_SPT_ELLIPSERIBBON2 = 108;
  MSO_SPT_FLOWCHARTPROCESS = 109;
  MSO_SPT_FLOWCHARTDECISION = 110;
  MSO_SPT_FLOWCHARTINPUTOUTPUT = 111;
  MSO_SPT_FLOWCHARTPREDEFINEDPROCESS = 112;
  MSO_SPT_FLOWCHARTINTERNALSTORAGE = 113;
  MSO_SPT_FLOWCHARTDOCUMENT = 114;
  MSO_SPT_FLOWCHARTMULTIDOCUMENT = 115;
  MSO_SPT_FLOWCHARTTERMINATOR = 116;
  MSO_SPT_FLOWCHARTPREPARATION = 117;
  MSO_SPT_FLOWCHARTMANUALINPUT = 118;
  MSO_SPT_FLOWCHARTMANUALOPERATION = 119;
  MSO_SPT_FLOWCHARTCONNECTOR = 120;
  MSO_SPT_FLOWCHARTPUNCHEDCARD = 121;
  MSO_SPT_FLOWCHARTPUNCHEDTAPE = 122;
  MSO_SPT_FLOWCHARTSUMMINGJUNCTION = 123;
  MSO_SPT_FLOWCHARTOR = 124;
  MSO_SPT_FLOWCHARTCOLLATE = 125;
  MSO_SPT_FLOWCHARTSORT = 126;
  MSO_SPT_FLOWCHARTEXTRACT = 127;
  MSO_SPT_FLOWCHARTMERGE = 128;
  MSO_SPT_FLOWCHARTOFFLINESTORAGE = 129;
  MSO_SPT_FLOWCHARTONLINESTORAGE = 130;
  MSO_SPT_FLOWCHARTMAGNETICTAPE = 131;
  MSO_SPT_FLOWCHARTMAGNETICDISK = 132;
  MSO_SPT_FLOWCHARTMAGNETICDRUM = 133;
  MSO_SPT_FLOWCHARTDISPLAY = 134;
  MSO_SPT_FLOWCHARTDELAY = 135;
  MSO_SPT_TEXTPLAINTEXT = 136;
  MSO_SPT_TEXTSTOP = 137;
  MSO_SPT_TEXTTRIANGLE = 138;
  MSO_SPT_TEXTTRIANGLEINVERTED = 139;
  MSO_SPT_TEXTCHEVRON = 140;
  MSO_SPT_TEXTCHEVRONINVERTED = 141;
  MSO_SPT_TEXTRINGINSIDE = 142;
  MSO_SPT_TEXTRINGOUTSIDE = 143;
  MSO_SPT_TEXTARCHUPCURVE = 144;
  MSO_SPT_TEXTARCHDOWNCURVE = 145;
  MSO_SPT_TEXTCIRCLECURVE = 146;
  MSO_SPT_TEXTBUTTONCURVE = 147;
  MSO_SPT_TEXTARCHUPPOUR = 148;
  MSO_SPT_TEXTARCHDOWNPOUR = 149;
  MSO_SPT_TEXTCIRCLEPOUR = 150;
  MSO_SPT_TEXTBUTTONPOUR = 151;
  MSO_SPT_TEXTCURVEUP = 152;
  MSO_SPT_TEXTCURVEDOWN = 153;
  MSO_SPT_TEXTCASCADEUP = 154;
  MSO_SPT_TEXTCASCADEDOWN = 155;
  MSO_SPT_TEXTWAVE1 = 156;
  MSO_SPT_TEXTWAVE2 = 157;
  MSO_SPT_TEXTWAVE3 = 158;
  MSO_SPT_TEXTWAVE4 = 159;
  MSO_SPT_TEXTINFLATE = 160;
  MSO_SPT_TEXTDEFLATE = 161;
  MSO_SPT_TEXTINFLATEBOTTOM = 162;
  MSO_SPT_TEXTDEFLATEBOTTOM = 163;
  MSO_SPT_TEXTINFLATETOP = 164;
  MSO_SPT_TEXTDEFLATETOP = 165;
  MSO_SPT_TEXTDEFLATEINFLATE = 166;
  MSO_SPT_TEXTDEFLATEINFLATEDEFLATE = 167;
  MSO_SPT_TEXTFADERIGHT = 168;
  MSO_SPT_TEXTFADELEFT = 169;
  MSO_SPT_TEXTFADEUP = 170;
  MSO_SPT_TEXTFADEDOWN = 171;
  MSO_SPT_TEXTSLANTUP = 172;
  MSO_SPT_TEXTSLANTDOWN = 173;
  MSO_SPT_TEXTCANUP = 174;
  MSO_SPT_TEXTCANDOWN = 175;
  MSO_SPT_FLOWCHARTALTERNATEPROCESS = 176;
  MSO_SPT_FLOWCHARTOFFPAGECONNECTOR = 177;
  MSO_SPT_CALLOUT90 = 178;
  MSO_SPT_ACCENTCALLOUT90 = 179;
  MSO_SPT_BORDERCALLOUT90 = 180;
  MSO_SPT_ACCENTBORDERCALLOUT90 = 181;
  MSO_SPT_LEFTRIGHTUPARROW = 182;
  MSO_SPT_SUN = 183;
  MSO_SPT_MOON = 184;
  MSO_SPT_BRACKETPAIR = 185;
  MSO_SPT_BRACEPAIR = 186;
  MSO_SPT_SEAL4 = 187;
  MSO_SPT_DOUBLEWAVE = 188;
  MSO_SPT_ACTIONBUTTONBLANK = 189;
  MSO_SPT_ACTIONBUTTONHOME = 190;
  MSO_SPT_ACTIONBUTTONHELP = 191;
  MSO_SPT_ACTIONBUTTONINFORMATION = 192;
  MSO_SPT_ACTIONBUTTONFORWARDNEXT = 193;
  MSO_SPT_ACTIONBUTTONBACKPREVIOUS = 194;
  MSO_SPT_ACTIONBUTTONEND = 195;
  MSO_SPT_ACTIONBUTTONBEGINNING = 196;
  MSO_SPT_ACTIONBUTTONRETURN = 197;
  MSO_SPT_ACTIONBUTTONDOCUMENT = 198;
  MSO_SPT_ACTIONBUTTONSOUND = 199;
  MSO_SPT_ACTIONBUTTONMOVIE = 200;
  MSO_SPT_HOSTCONTROL = 201;
  MSO_SPT_TEXTBOX = 202;
  MSO_SPT_NIL = $0FFF;
  MSO_SPT_MAX = MSO_SPT_NIL;

  { Bits in OfficeArtFSp record }
  MSO_FSP_BITS_GROUP              = $00000001;
  MSO_FSP_BITS_CHILD              = $00000002;
  MSO_FSP_BITS_PATRIARCH          = $00000004;
  MSO_FSP_BITS_DELETED            = $00000008;
  MSO_FSP_BITS_OLESHAPE           = $00000010;
  MSO_FSP_BITS_HASMASTER          = $00000020;
  MSO_FSP_BITS_FLIPHOR            = $00000040;
  MSO_FSP_BITS_FLIPVERT           = $00000080;
  MSO_FSP_BITS_CONNECTOR          = $00000100;
  MSO_FSP_BITS_HASANCHOR          = $00000200;
  MSO_FSP_BITS_BACKGROUND         = $00000400;
  MSO_FSP_BITS_HASSHAPETYPE       = $00000800;

  { Identifier of property array items if OfficeArtFOpt record }
  MSO_FOPT_ID_TEXTID              = $0080;
  MSO_FOPT_ID_TEXTDIRECTION       = $008B;
  MSO_FOPT_ID_TEXTBOOL            = $00BF;
  MSO_FOPT_ID_CONNECTIONPOINTTYPE = $0158;
  MSO_FOPT_ID_FILLCOLOR           = $0181;
  MSO_FOPT_ID_FILLBACKGROUNDCOLOR = $0183;
  MSO_FOPT_ID_FILLFOREGROUNDCOLOR = $0185;
  MSO_FOPT_ID_FILLBOOL            = $01BF;
  MSO_FOPT_ID_SHADOWCOLOR         = $0201;
  MSO_FOPT_ID_SHADOWBOOL          = $023F;
  MSO_FOPT_ID_GROUPBOOL           = $03BF;

procedure WriteMSOClientAnchorSheetRecord(AStream: TStream;
  ATopRow, ALeftCol, ABottomRow, ARightCol,
  ALeftMargin, ARightMargin, ATopMargin, ABottomMargin: Word;
  AMoveIntact, AResizeIntact: Boolean);
procedure WriteMSOClientDataRecord(AStream: TStream);
procedure WriteMSOClientTextboxRecord(AStream: TStream);
procedure WriteMSODgContainer(AStream: TStream; ASize: DWord);
procedure WriteMSOFDgRecord(AStream: TStream; ANumShapes, ADrawingID, ALastObjID: Word);
procedure WriteMSOFOptRecord_Comment(AStream: TStream);
procedure WriteMSOProperty(AStream: TStream; APropertyID: Word; AValue: DWord);
procedure WriteMSOFSpRecord(AStream: TStream; AShapeID: DWord; AShapeType: Word; ABits: DWord);
procedure WriteMSOFSpGrRecord(AStream: TStream; ALeft, ATop, ARight, ABottom: DWord);
procedure WriteMSOHeader(AStream: TStream; AType, AVersion, AInstance: Word; ARecSize: DWord);
procedure WriteMSOSpContainer(AStream: TStream; ASize: DWord);
procedure WriteMSOSpGrContainer(AStream: TStream; ASize: DWord);

implementation

uses
  fpsutils;

type
  TsMSOHeader = packed record
    Version_Instance: Word;
    RecordType: Word;
    RecordSize: DWord;
  end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtClientAnchorSheet record to a stream.

  The OfficeArtClientAnchorSheet structure specifies the anchor position of
  a drawing object embedded in a sheet.

  Ref: [MS-XLS].pdf
-------------------------------------------------------------------------------}
procedure WriteMSOClientAnchorSheetRecord(AStream: TStream;
  ATopRow, ALeftCol, ABottomRow, ARightCol,
  ALeftMargin, ARightMargin, ATopMargin, ABottomMargin: Word;
  AMoveIntact, AResizeIntact: Boolean);
const
  fMOVE = $0001; // specifies whether the shape will be kept intact when the cells are moved.
  fSIZE = $0002; // specifies whether the shape will be kept intact when the cells are resized.
var
  flags: Word;
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_CLIENTANCHORSHEET, 0, 0, 18);

  flags := 0;
  if AMoveIntact then begin
    AResizeIntact := true;
    flags := flags or fMOVE;
  end;
  if AResizeIntact then
    flags := flags or fSIZE;
  AStream.WriteWord(WordToLE(flags));

  // Column of the cell under the top left corner of the bounding rectangle of the shape.
  AStream.WriteWord(WordLEToN(ALeftCol));

  // x coordinate of the top left corner of the bounding rectangle relative to
  // the corner of the underlying cell.
  // The value is expressed as 1024th’s of that cell’s width.
  AStream.WriteWord(WordLEToN(ALeftMargin));

  // Row of the cell under the top left corner of the bounding rectangle of the shape.
  AStream.WriteWord(WordLEToN(ATopRow));

  // y coordinate of the top left corner of the bounding rectangle relative to
  // the corner of the underlying cell.
  // The value is expressed as 256th’s of that cell’s height.
  AStream.WriteWord(WordLEToN(ATopMargin));

  // Column of the cell under the bottom right corner of the bounding rectangle
  // of the shape.
  AStream.WriteWord(WordToLE(ARightCol));

  // x coordinate of the bottom right corner of the bounding rectangle relative
  // to the corner of the underlying cell.
  // The value is expressed as 1024th’s of that cell’s width.
  AStream.WriteWord(WordToLE(ARightMargin));

  // Row of the cell under the bottom right corner of the bounding rectangle
  // of the shape.
  AStream.WriteWord(WordToLE(ABottomRow));

  // y coordinate of the bottom right corner of the bounding rectangle relative
  // to the corner of the underlying cell.
  // The value is expressed as 256th’s of that cell’s height.
  AStream.WriteWord(WordToLE(ABottomMargin));
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtClientData record to a stream

  The OfficeArtClientData structure specifies the client data of a drawing
  object.

  MUST be the last structure of the rgChildRec field of the current
  MSODRAWING BIFF record.

  The next record MUST be OBJ which contains the detailed data information
  about this drawing object.
-------------------------------------------------------------------------------}
procedure WriteMSOClientDataRecord(AStream: TStream);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_CLIENTDATA, 0, 0, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtClientTextbox record to a stream
-------------------------------------------------------------------------------}
procedure WriteMSOClientTextboxRecord(AStream: TStream);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_CLIENTTEXTBOX, 0, 0, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtDgContainer record to a stream.

  The OfficeArtDgContainer record specifies the container for all file records
  for the objects in an MSO drawing.
-------------------------------------------------------------------------------}
procedure WriteMSODgContainer(AStream: TStream; ASize: DWord);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_DG_CONTAINER, MSO_VER_CONTAINER, 0, ASize);
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArt FDG record to a stream.

  The OfficeArtFDG record specifies the number of shapes, the drawing identifier,
  and the shape identifier of the last shape in a drawing.
-------------------------------------------------------------------------------}
procedure WriteMSOFdgRecord(AStream: TStream; ANumShapes, ADrawingID, ALastObjID: Word);
begin
  if ADrawingID > $0FFE then
    raise Exception.CreateFmt('[WriteMSOFdgRecord] Invalid drawing identifier $%.4x', [ADrawingID]);
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_FDG, 0, ADrawingID, 8);
  AStream.WriteDWord(DWordToLE(ANumShapes));
  AStream.WriteDWord(DWordToLE(ALastObjID));
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtFOpt as it is used for a cell comment record to the stream

  The OfficeArtFOPT record specifies a table of OfficeArtRGFOPTE records,

  The OfficeArtRGFOPTE record specifies a property table, which consists of an
  array of fixed-size property table entries, followed by a variable-length
  field of complex data.
-------------------------------------------------------------------------------}
procedure WriteMSOFOptRecord_Comment(AStream: TStream);
const
  NUM_PROPERTIES = 13;
begin
  // Escher header
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_FOPT, 3, NUM_PROPERTIES, NUM_PROPERTIES*6);

  // TextID
  WriteMSOProperty(AStream, MSO_FOPT_ID_TEXTID, 0);

  // Text direction
  WriteMSOProperty(AStream, MSO_FOPT_ID_TEXTDIRECTION, 2);  // 2 = "determined from text string"

  // Boolean properties of text in a shape
  WriteMSOProperty(AStream, MSO_FOPT_ID_TEXTBOOL, $00080008);

  // Type of connection point
  WriteMSOProperty(AStream, MSO_FOPT_ID_CONNECTIONPOINTTYPE, 0);

  // Fill color
  WriteMSOProperty(AStream, MSO_FOPT_ID_FILLCOLOR, $00E1FFFF);

  // Background color of fill
  WriteMSOProperty(AStream, MSO_FOPT_ID_FILLBACKGROUNDCOLOR, $00E1FFFF);

  // Foreground color of fill
  WriteMSOProperty(AStream, MSO_FOPT_ID_FILLFOREGROUNDCOLOR, $100000F4);

  // Fill style boolean properties
  WriteMSOProperty(AStream, MSO_FOPT_ID_FILLBOOL, $00100010);

  // Line foreground color for black-and-white mode
  WriteMSOProperty(AStream, $01C3, $100000F4);

  // Shadow color
  WriteMSOProperty(AStream, MSO_FOPT_ID_SHADOWCOLOR, 0);

  // Shadow color primary color modifier if in black-and-white mode
  WriteMSOProperty(AStream, $0203, $100000F4);

  // Shadow style boolean properties
  WriteMSOProperty(AStream, MSO_FOPT_ID_SHADOWBOOL, $00030003);

  // Group shape boolean properties
  WriteMSOProperty(AStream, MSO_FOPT_ID_GROUPBOOL, $00020002);
end;

{@@ ----------------------------------------------------------------------------
  Writes a property of the FOPT array
-------------------------------------------------------------------------------}
procedure WriteMSOProperty(AStream: TStream; APropertyID: Word;
  AValue: DWord);
begin
  AStream.WriteWord(WordToLE(APropertyID));
  AStream.WriteDWord(DWordToLE(AValue));
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtFSp record to the stream

  The OfficeArtFSP record specifies an instance of a shape.
  The record header contains the shape type, and the record itself contains
  the shape identifier and a set of bits that further define the shape.
-------------------------------------------------------------------------------}
procedure WriteMSOFSpRecord(AStream: TStream; AShapeID: DWord;
  AShapeType: Word; ABits: DWord);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_FSP, 2, AShapeType, 8);
  AStream.WriteDWord(DWordToLE(AShapeID));
  AStream.WriteDWord(DWordToLE(ABits));
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtFSpGr record to the stream.

  The OfficeArtFSPGR record specifies the coordinate system of the group shape
  that the anchors of the child shape are expressed in.
  This record is present only for group shapes.
-------------------------------------------------------------------------------}
procedure WriteMSOFSpGrRecord(AStream: TStream; ALeft, ATop, ARight, ABottom: DWord);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_FSPGR, 1, 0, 16);
  AStream.WriteDWord(DWordToLE(ALeft));
  AStream.WriteDWord(DWordToLE(ATop));
  AStream.WriteDWord(DWordToLE(ARight));
  AStream.WriteDWord(DWordToLE(ABottom));
end;

{ Writes the header of an MSO subrecord used internally by MSODRAWING records }
procedure WriteMSOHeader(AStream: TStream; AType, AVersion, AInstance: Word;
  ARecSize: DWord);
var
  rec: TsMSOHeader;
begin
  rec.Version_Instance := WordToLE((AVersion and $000F) + AInstance shl 4); //and $FFF0) shr 4);
  // To do: How to handle Version_Instance on big-endian machines?
  // Version_Instance combines bits 0..3 for "version" and 4..15 for "instance"
  rec.RecordType := WordToLE(AType);
  rec.RecordSize := DWordToLE(ARecSize);
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an OffcieARtSpContainer record to the stream.
  The OfficeArtSpContainer record specifies a shape container.
-------------------------------------------------------------------------------}
procedure WriteMSOSpContainer(AStream: TStream; ASize: DWord);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_SP_CONTAINER, MSO_VER_CONTAINER, 0, ASize);
end;

{@@ ----------------------------------------------------------------------------
  Writes an OfficeArtSpGrContainer record to a stream.
  The OfficeArtSpgrContainer record specifies a container for groups of shapes.
  The group container contains a variable number of shape containers and other
  group containers. Each group is a shape. The first container MUST be an
  OfficeArtSpContainer record, which MUST contain shape information for the
  group.
-------------------------------------------------------------------------------}
procedure WriteMSOSpGrContainer(AStream: TStream; ASize: DWord);
begin
  WriteMSOHeader(AStream, MSO_ID_OFFICEART_SPGR_CONTAINER, MSO_VER_CONTAINER, 0, ASize);
end;


end.

