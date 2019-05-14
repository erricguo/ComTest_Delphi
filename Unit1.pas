unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, CommTypes, CommDev, StdCtrls, POSButtons, POSChkBox, ComCtrls;

type
  TForm1 = class(TForm)
    ckInvPrinterStatus: TPOSCheckBox;
    cdInvPrinterDefaultReset: TPOSCheckBox;
    Button1: TPOSButton;
    Button2: TPOSButton;
    btFF: TPOSButton;
    btLF: TPOSButton;
    Label5: TLabel;
    ComboBox2: TComboBox;
    ComboBox1: TComboBox;
    Label6: TLabel;
    ComboBox3: TComboBox;
    Label7: TLabel;
    edUniName: TEdit;
    btPrint: TPOSButton;
    btCutPaper: TPOSButton;
    Memo2: TMemo;
    InvPrinter: TInvPrinter;
    StatusBar: TStatusBar;
    ckDSR: TPOSCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CheckStatus(xConnect, xTransmit: integer);
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  try
    InvPrinter.DefaultReset := cdInvPrinterDefaultReset.Checked;
    InvPrinter.ConnectType := CommTypes.TConnectMode(ComboBox2.ItemIndex);
    InvPrinter.Port := TDevicePort(ComboBox1.ItemIndex);
    InvPrinter.PrinterType := TPrinterType(ComboBox3.ItemIndex);
    InvPrinter.UniName := edUniName.Text;
    InvPrinter.CheckDSR := ckDSR.Checked;

    if InvPrinter.Connect then
    begin
      ShowMessage('已連線');
      CheckStatus(1, 2)
    end
    else
    begin
      ShowMessage('連線失敗');
        CheckStatus(0, 2);
    end;
  except
    CheckStatus(0, 0);
  end;
end;

procedure TForm1.CheckStatus(xConnect, xTransmit: integer);
begin
    with StatusBar.Panels do
    begin
        case xConnect of
            0:  Items[0].Text := '斷線!';
            1:  Items[0].Text := '已連接!';
            2:  Items[0].Text := '';
        end;
        case xTransmit of
            0:  Items[1].Text := '傳送失敗!';
            1:  Items[1].Text := '已傳送成功!';
            2:  Items[1].Text := '';
        end;
    end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
 i:integer;
begin
  //初始發票機資料
  ComboBox3.Items.Clear;
  for i := Low(StrPrinterType) to High(StrPrinterType) do
    ComboBox3.Items.Add(StrPrinterType[i]);
end;

end.
