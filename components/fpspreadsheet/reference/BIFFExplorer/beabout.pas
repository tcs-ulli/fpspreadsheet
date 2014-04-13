unit beAbout;

{$mode objfpc}{$H+}

interface

uses
  Classes, IpHtml, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls;

type

  { TAboutForm }

  TAboutForm = class(TForm)
    Bevel1: TBevel;
    BtnClose: TButton;
    IconImage: TImage;
    HTMLViewer: TIpHtmlPanel;
    LblTitle: TLabel;
    Panel1: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure HTMLViewerHotClick(Sender: TObject);
  private
    { private declarations }
    function CreateCredits: String;
  public
    { public declarations }
  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.lfm}

uses
  LCLIntf, types, beHTML;

{ TAboutForm }

function TAboutForm.CreateCredits: String;
var
  html: THTMLDocument;
  clrs: THeaderColors = (clBlack, clBlack, clBlack, clBlack, clBlack);
begin
  html := THTMLDocument.Create;
  try
    clrs[h3] := LblTitle.Font.Color;
    clrs[h4] := LblTitle.Font.Color;
    with html do begin
      BeginDocument('Credits', clrs, false);
        AddHeader(h3, 'Credits');
        AddHeader(h4, 'Libraries');
        BeginBulletList;
          AddListItem(Hyperlink(
            'Free Pascal',
            'www.freepascal.org')
          );
          AddListItem(Hyperlink(
            'Lazarus',
            'www.lazarus.freepascal.org')
          );
          AddListItem(HyperLink(
            'fpspreadsheet',
            'http://sourceforge.net/p/lazarus-ccr/svn/HEAD/tree/components/fpspreadsheet/')
          );
        EndBulletList;

        AddEmptyLine;

        AddHeader(h4, 'Icons');
        BeginBulletList;
          AddListItem(HyperLink(
            'Fugue icons',
            'http://p.yusukekamiyamane.com/')
            + ' (for toolbar icons)');
          AddListItem(HyperLink(
            'Nuvola icons',
            'www.icon-king.com/projects/nuvola/') +
            ' (for application icon');
        EndBulletList;

        AddEmptyLine;

        AddHeader(h4, 'Used documentation');
        BeginBulletList;
          AddListItem(Hyperlink(
            'OpenOffice.org''s Documentation of the Microsoft Excel File Format',
            'http://www.openoffice.org/sc/excelfileformat.pdf') +
            ' (see folder "fpspreadsheet/reference")'
          );
          AddListItem(Hyperlink(
            '[MS-XLS]: Excel Binary File Format (.xls) Structure',
            'http://msdn.microsoft.com/en-us/library/cc313154%28v=office.12%29.aspx'
          ));
          AddListItem(HyperLink(
            'Excel97-2007BinaryFileFormat(xls)Specification',
            'http://download.microsoft.com/download/0/B/E/0BE8BDD7-E5E8-422A-ABFD-4342ED7AD886/Excel97-2007BinaryFileFormat(xls)Specification.pdf'
          ));
        EndBulletList;

      EndDocument;

      Result := Lines.Text;
    end;
  finally
    html.Free;
  end;
end;


procedure TAboutForm.FormCreate(Sender: TObject);
var
  ico: TIcon;
  sz: TSize;
begin
  ico := TIcon.Create;
  try
    ico.Assign(Application.Icon);
    sz.cx := 48;
    sz.cy := 48;
    ico.Current := ico.GetBestIndexForSize(sz);
    IconImage.Picture.Assign(ico);
  finally
    ico.Free;
  end;

  HTMLViewer.SetHTMLFromStr(CreateCredits);
end;


procedure TAboutForm.HTMLViewerHotClick(Sender: TObject);
begin
  OpenURL(HTMLViewer.HotURL);
end;


end.

