unit sfCurrencyForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ButtonPanel, StdCtrls, ExtCtrls, Buttons;

type

  { TCurrencyForm }

  TCurrencyForm = class(TForm)
    BtnAdd: TBitBtn;
    BtnDelete: TBitBtn;
    ButtonPanel: TButtonPanel;
    LblInfo: TLabel;
    CurrencyListbox: TListBox;
    Panel1: TPanel;
    procedure BtnAddClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  CurrencyForm: TCurrencyForm;

implementation

uses
  fpscurrency;

{ TCurrencyForm }

procedure TCurrencyForm.FormCreate(Sender: TObject);
begin
  GetRegisteredCurrencies(CurrencyListbox.Items);
  CurrencyListbox.ItemIndex := CurrencyListbox.Items.Count-1;
end;

procedure TCurrencyForm.BtnAddClick(Sender: TObject);
var
  s: String;
  i: Integer;
begin
  s := InputBox('Input', 'Currency symbol:', '');
  if s <> '' then begin
    i := CurrencyListbox.Items.IndexOf(s);
    if i = -1 then
      i := CurrencyListbox.Items.Add(s);
    CurrencyListbox.ItemIndex := i;
  end;
end;

procedure TCurrencyForm.BtnDeleteClick(Sender: TObject);
begin
  if CurrencyListbox.ItemIndex > -1 then
    CurrencyListbox.Items.Delete(CurrencyListbox.ItemIndex);
end;

procedure TCurrencyForm.OKButtonClick(Sender: TObject);
begin
  RegisterCurrencies(CurrencyListbox.Items, true);
end;


initialization
  {$I sfCurrencyForm.lrs}

end.

