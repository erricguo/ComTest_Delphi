unit CommTypes;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls;

type
  TWindowPlat = (wp_WinNT, wp_Win32s, wp_Win98);

  //Define Communication Device Port
  TDevicePort     = (cp_COM1, cp_COM2, cp_COM3,
                     cp_COM4, cp_COM5, cp_COM6,
                     lp_LPT1, lp_LPT2, lp_LPT3,
                     lp_LPT4, lp_LPT5, lp_LPT6,
                     cp_COM7, cp_COM8, cp_COM9);

  //Define Communication Device Baud Rate
  TBaudRateType   = (brt_110,      brt_300,    brt_600,
                     brt_1200,     brt_2400,   brt_4800,
                     brt_9600,     brt_14400,  brt_19200,
                     brt_38400,    brt_56000,  brt_57600,
                     brt_115200,   brt_128000, brt_256000);

  //Define Communication Device Data Bits
  TDataBitsType   = (dbt_5, dbt_6, dbt_7, dbt_8);

  //Define Communication Device Stop Bits;
  TStopBitsType   = (sbt_OneStopBit, sbt_One5StopBits, sbt_TwoStopBits);

  //Define Communication Device Parity Check
  TParityCheckType= (pct_NoParity, pct_OddParity, pct_EvenParity, pct_MarkParity, pct_SpaceParity);

  //Define Printer Type
  TPrinterType    = (TP_3688,
                     EPSON_TM200,
                     EPSON_TMT882,
                     EPSON_TMU295,
                     EPSON_TMU300,
                     EPSON_RPU420,
                     EPSON_TMU675,
                     EPSON_MD332S,
                     NIXDORF_20,
                     NIXDORF_77,
                     NIXDORF_210,
                     IBM_4610,
                     NormalType,
                     STAR_SP320,
                     EZ2_PrintServer,
                     WindowsPrinter,
                     WP_T810, RP_600,
                     PP8000, EPSON_TM88IV,
                     RP_600_V9, EPSON_TMT70);

  //Define Device Connection Mode
  TConnectMode    = (cm_ParallelPort, cm_SerialPort, cm_PrintServer);

  //Define Printer Receipt or Journal Mode
  TPtrDir         = (pd_Both, pd_Receipt, pd_Journal);

  //Define Printer Printing Out Mode
  TOutputMode     = (om_Printer, om_Screen);

  //Define Cutomer Display Type
  TDisplayerType  = (CD_5220,
                     CD_8240,
                     VC_108,
                     CC_110,
                     WD_304,
                     RS_304,
                     PD_2100,
                     IBM_SURE_ONE,
                     NIXDORF_BAL3G,
                     NIXDORF_BAL3,
                     DM_D110,
                     DM_D500,
                     PD_7200,
                     FUJITSU_KD_290X,
                     CD_9256);

  //Define Bar Code Scanner Type
  TBarCodeScannerType = (HS_232, FJ_M411);

  //Define EDC Type
  TEdcType            = (H5000,
                         S9000,
                         SAGEM,
                         VX570_400,
                         ICT220_CATHAY,
                         VX510_ESun,
                         NETPOS3000,
                         AS320,
                         VEGA_9000,                                             //20130115 ADD BY 4084
                         AS320_200,
                         AllPay,
                         AS320_UnionPay,
                         SinoPac,
                         TaiHsin_400,             //20170525 add by 07113
                         FirstBank_400,
                         UniversalEDC,
                         ESUN_S80RF_600,  //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002
                         CTBC_AS320_250   //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002
                         );

  PCreditCard = ^TCreditCard;
  CREDITCARD = record
                Host_id,
                Card_Name,
                Auth_no,
                Card_no,
                Ref_no,
                Response,
                Transfinished,
                Timerblockreset  :   AnsiString;
                //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
                TransDate,
                TransAmt,
                InvoiceEncryptionCardNo,
                SHACardNo:AnsiString;
                //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
  end;
  TCreditCard = CREDITCARD;

  //�w�q��L���
  PKeyboard = ^KEYBOARD;
  KEYBOARD = record
      ScanCode:   integer;        //��L�y�X
      FuncCode,                   //�\��N�X
      PayID,                      //�I�ڥN�X
      Remark:     String;         //�Ƶ�����
      Level:      integer;        //�v������
end;
  TKeyboard = KEYBOARD;

    //�w�q�Ȥ���ܾ�
  PCustDisplay = ^CUSTDISPLAY;
  CUSTDISPLAY = record
      Kind,                       //����\��O
      DisplayType,                //�������
      Port,                       //�q�T��N�X
      Words:          integer;    //��ܦr��
      WelcomeStr:      string;    //����ݭԻy
      ShowDateTime:   boolean;    //�۰���ܤ���ɶ�
      ShowDisc    :   boolean;    //��ܧ馩���B
      UniName:         string;    //�����|�W��
  end;
  TCustDiaplay = CUSTDISPLAY;

  //�w�q�o����
  PInvoicePrinter = ^INVOICEPRINTER;
  INVOICEPRINTER = record
      Enabled:        boolean;    //�ϥΦL���
      Kind,                       //�L����\��O
      PrinterType:    Integer;    //�L�������
      Port:           Integer;    //�q�T��N�X
      UniName:         string;    //�����|�W��
      PrintDelay:     Integer;    //�C�L����ɶ�
      CheckLPTStatus,             //�ˬd�L���
      DefaultReset:   boolean;    //�ǰeDefault Reset���O
      PaperCnt:       Integer;    //�C�L�i��
  end;
  TInvoicePrinter = INVOICEPRINTER;

  //�w�q���y��
  PScanner = ^SCANNER;
  SCANNER = record
      Kind,                       //���y���\��O
      ScannerType:    Integer;    //���y������
      Port,                       //�q�T��N�X
      BaudRate,                   //�t�v
      DataBits,                   //��Ʀ줸
      StopBits,                   //����줸
      ParityCheck:    Integer;    //�P���ˬd
      UniName:         string;    //�����|�W��
  end;
  TScanner = SCANNER;

  //�w�q����Ū�d��
  PICReader = ^ICREADER;
  ICREADER = record
      OpenReady,                      //�}��com���A
      ReaderLink,                     //gS516200, //'�PŪ�d���s�u';
      ReaderAutoLink:     boolean;    //gS516210, //'Ū�d���۰ʳs�u';;
      SytemType,                      //gS516220, //'�t������'
      RederModel,                     //gS516230, //'Ū�d������';
      COMPort:            integer;    //gS516240, //'��X��';
      PrintFormat:        string;     //gS516250, //'�����d���C�L�榡';
  end;
  TICREADER = ICREADER;

  //Define Base Communication Device Component
  TBaseCOM = class(TComponent)
  private
    FHandle: THandle;
    FDevicePort: TDevicePort;
    FBaudRate: TBaudRateType;
    FDataBits: TDataBitsType;
    FStopBits: TStopBitsType;
    FUniName: AnsiString;
    FParityCheck: TParityCheckType;
    FInbondBuffer, FOutbondBuffer, FWriteTimeOut, FDelay: integer;
    FCheckDSR: boolean;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure SetUniName(Value: AnsiString); virtual;
    function space(s: integer): AnsiString;
    function StrSpace(Str: AnsiString; i: integer): AnsiString;
  published
    property CheckDSR: boolean read FCheckDSR write FCheckDSR;
    property Port: TDevicePort read FDevicePort write FDevicePort;
    property WTimeOut: integer read FWriteTimeOut write FWriteTimeOut;
    property Delay: integer read FDelay write FDelay;
    property BaudRate: TBaudRateType read FBaudRate write FBaudRate;
    property DataBits: TDataBitsType read FDataBits write FDataBits;
    property StopBits: TStopBitsType read FStopBits write FStopBits;
    property ParityCheck: TParityCheckType read FParityCheck write FParityCheck;
    property InbondBuffer: integer read FInbondBuffer write FInbondBuffer;
    property OutbondBuffer: integer read FOutbondBuffer write FOutbondBuffer;
    property UniName: AnsiString read FUniName write SetUniName;
  end;

var
    DevicePortStrValue: array[TDevicePort] of PChar = ('COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6',
                                                       'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6',
                                                       'COM7', 'COM8', 'COM9');
    BaudRateValue: array[TBaudRateType] of DWORD = (CBR_110,   CBR_300,    CBR_600,  CBR_1200,  CBR_2400,   CBR_4800,
                                                    CBR_9600,  CBR_14400,  CBR_19200,CBR_38400, CBR_56000,  CBR_57600,
                                                    CBR_115200,CBR_128000, CBR_256000);
    DataBitsValue: array[TDataBitsType] of integer = (5, 6, 7, 8);
    StopBitsValue: array[TStopBitsType] of integer = (ONESTOPBIT, ONE5STOPBITS, TWOSTOPBITS);
    ParityCheckValue: array[TParityCheckType] of integer = (NOPARITY, ODDPARITY, EVENPARITY, MARKPARITY, SPACEPARITY);

    //�w�q�귽
    StrComPortValue   : array[0..8]  of AnsiString = ('COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9');
    StrDevicePortValue: array[0..14] of AnsiString = ('COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6',
                                                      'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6',
                                                      'COM7', 'COM8', 'COM9');
    StrBaudRateValue: array[0..14] of AnsiString = ('CBR_110',   'CBR_300',    'CBR_600',  'CBR_1200',  'CBR_2400',   'CBR_4800',
                                                    'CBR_9600',  'CBR_14400',  'CBR_19200','CBR_38400', 'CBR_56000',  'CBR_57600',
                                                    'CBR_115200','CBR_128000', 'CBR_256000');
    StrDataBitsValue:    array[0..3] of AnsiString = ('5', '6', '7','8');
    StrStopBitsValue:    array[0..2] of AnsiString = ('ONESTOPBIT', 'ONE5STOPBITS', 'TWOSTOPBITS');
    StrParityCheckValue: array[0..4] of AnsiString = ('NOPARITY', 'ODDPARITY', 'EVENPARITY', 'MARKPARITY', 'SPACEPARITY');

    StrPrinterType: array[0..21] of AnsiString = ('TP_3688',      'EPSON_TM200',
                                                  'EPSON_TMT882', 'EPSON_TMU295',
                                                  'EPSON_TMU300', 'EPSON_RPU420',
                                                  'EPSON_TMU675', 'EPSON_MD332S',
                                                  'NIXDORF_20',   'NIXDORF_77',
                                                  'NIXDORF_210',  'IBM_4610',
                                                  'NormalType',   'STAR_SP320',
                                                  'EZ2_PrintServer', 'WindowsPrinter',
                                                  'WP_T810',      'RP_600',
                                                  'PP8000',       'EPSON_TM88IV',
                                                  'RP_600_V9',    'EPSON_TMT70');

    StrConnectMode:   array[0..2]  of AnsiString = ('ParallelPort', 'SerialPort', 'cm_PrintServer');
    StrPtrDir:        array[0..2]  of AnsiString = ('Both', 'Receipt', 'Journal');
    StrOutputMode:    array[0..1]  of AnsiString = ('Printer', 'Screen');
    StrDisplayerType: array[0..14] of AnsiString = ('CD_5220', 'CD_8240',
                                                    'VC_108',  'CC_110',
                                                    'WD_304',  'RS_304',
                                                    'PD_2100', 'IBM_SURE_ONE',
                                                    'NIXDORF_BAL3G', 'NIXDORF_BAL3',
                                                    'DM_D110', 'DM_D500', 'PD_7200', 'FUJITSU_KD_290X', 'CD_9256');

    StrBarCodeScannerType: array[0..1] of AnsiString = ('HS_232', 'FJ_M411');
    StrEdcType:            array[0..17] of AnsiString = ('H5000', 'S9000', 'SAGEM',
                                                        'VX570_400', 'ICT220_CATHAY', 'VX510_ESun',
                                                        'NETPOS3000', 'AS320', 'VEGA_9000', 'AS320_200',
                                                        'AllPay', 'AS320_UnionPay', 'SinoPac','TaiHsin_400',
                                                        'FirstBank_400', 'UniversalEDC', //20171122 edc.exe
                                                        'ESUN_S80RF_600','CTBC_AS320_250');   //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002

implementation

{$R *.res}

//TCommDev

constructor TBaseCOM.Create(AOwner: TComponent);
begin
    inherited Create(AOwner);
    FDelay := 0;
    FInbondBuffer := 2048;
    FOutbondBuffer := 2048;
    FWriteTimeOut := 300;
    FDevicePort := cp_COM1;
end;

destructor TBaseCOM.Destroy;
begin
    inherited Destroy;
end;

procedure TBaseCOM.SetUniName(Value: AnsiString);
begin
    FUniName := Value;
end;

function TBaseCOM.space(s: integer): AnsiString;
var
    i: integer;
begin
    Result := '';
    for i :=1 to s do
        Result := Result + ' ';
end;

function TBaseCOM.StrSpace(Str: AnsiString; i: integer): AnsiString;
var
    mLen, mLength: integer;
    mRight: boolean;
begin
    mLen := 1;
    mRight := true;

    if (i < 0) then mRight := false;
    if (i < 0) then
        mLen   := -i
    else
        mLen   := i;

    if (mLen - Length(Str) < 0) then
        Result := copy(Str, 1, mLen)
    else
        mLength := mLen - Length(Str);

    if mRight then
        Result := space(mLength) + Str
    else
        Result := Str + space(mLength);
end;


end.
