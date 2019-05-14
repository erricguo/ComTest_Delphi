unit CommDev;
//20120320 modi by 3188修正CD8240可與TP3688共用

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Dialogs, forms,
  CommTypes, zlportio, Printers, extctrls, Stdctrls, poscom;

type
  //TInvPrinter
  TOnPtrErrorEvent = procedure(Sender: TObject; Msg: string;
      ErrorCode: integer; var Retry: boolean) of object;

  TInvPrinter = class(TBaseCom)
  private
    FHandle: THandle;
    FDCB: DCB;
    FDevicePort: TDevicePort;
    FPrinter: TPrinterType;
    FOutputMode: TOutputMode;
    FPtrDir: TPtrDir;
    FStart, FEnd, FLine, FReset, FLineFeed, FFormFeed, FStamp,
        FCutPaper, FCashDrawer, FDir, FOtherEnd, FValidate: AnsiString;
    FLineCount: integer;
    FConnectMode: TConnectMode;
    FPrintStringList: TStringList;
    FActive, FCheckLPTStatus: boolean;
    FErrorMsg: string;
    FOnPtrError: TOnPtrErrorEvent;
    FRetry, FDefaultReset: boolean;
    FTimeOuts: COMMTIMEOUTS;
    procedure SetDefaultCommand;
    procedure SetPrinter(Value: TPrinterType);
    procedure SetOutPutMode(Value: TOutputMode);
    procedure SetConnectMode(Value: TConnectMode);
    procedure SetDirection(Value: TPtrDir);
    procedure SetActive(Value: boolean);
    function WriteString(SendString, OrgString: AnsiString): boolean;
    function WriteUTF8String(SendString, OrgString: AnsiString): boolean;
    function Connected: boolean;
    procedure ApplySetup;
    procedure SetPrinterIndex(Value: integer);
    function GetPrinters: TStrings;
    function GetPrintStringList: TStringList;
    function GetPrinterIndex: integer;
    function GetPlatForm: TWindowPlat;
    procedure StrConvert(str1: AnsiString; var str2: AnsiString);
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    function Connect:boolean;
    function DisConnect: boolean;
    function FF: boolean;
    function LF(Lines: integer): boolean;
    function Reset: boolean;
    function Stamp: boolean;
    function CutPaper: boolean;
    procedure BeginDoc();
    procedure EndDoc();
    function Valid(SendString: AnsiString): boolean;
    function OpenCashDrawer: boolean;
    function SendCOMString(SendString: AnsiString): boolean;
    function SendCOMUTF8String(SendString: AnsiString): boolean;
    function SendString(SendString: AnsiString): boolean;
    function PrintBarcode(xBarcode: AnsiString): boolean;
    function IsOnLine: boolean;
    procedure SetUniName(Value: AnsiString); override;
    function PrintQRcode(xBarcode: AnsiString; xBackHeight:integer): boolean;       //WP-T810用
    function PrintQRcode2014(xBarcode: AnsiString; xBackHeight:integer;xOffset:integer=2;xLv:Integer=6): boolean;    //20180903 add by 02953 for Cust-20180903001
    function PrintBarcode2(xBarcode: AnsiString): boolean;                          //WP-T810用
    function PrintBarcode3(xBarcode: AnsiString): boolean;
    function Logo:boolean;
    procedure EndPageMode;
published
    property PortHandle: THandle read FHandle write FHandle;
    property Port;
    property Delay;
    property DefaultReset: boolean read FDefaultReset write FDefaultReset;
    property CheckDSR;
    property CheckLPTStatus: boolean read FCheckLPTStatus write FCheckLPTStatus;
    property WTimeOut;
    property BaudRate;
    property DataBits;
    property StopBits;
    property ParityCheck;
    property InbondBuffer;
    property OutbondBuffer;
    property UniName;
    property CheckConnected:Boolean read Connected;
    property ErrorMsg: string read FErrorMsg write FErrorMsg;
//    property Printer: TPrinterType read FPrinter write SetPrinter;
	  property WinPrinters: TStrings read GetPrinters;
    property PrinttStringList: TStringList read GetPrintStringList;
    property Active: boolean read FActive write SetActive;
    property Direction: TPtrDir read FPtrDir write SetDirection;
    property PrinterType: TPrinterType read FPrinter write SetPrinter;
    property OutPutMode: TOutputMode read FOutPutMode write SetOutPutMode;
    property ConnectType: TConnectMode read FConnectMode write SetConnectMode;
  	property PrinterIndex: Integer read GetPrinterIndex write SetPrinterIndex nodefault;
    property OnPtrError: TOnPtrErrorEvent read FOnPtrError write FOnPtrError;
  end;

  //TDisplay

  TDisplay = class(TBaseCom)
  private
    FHandle: THandle;
    FDisplayer: TDisplayerType;
    FDCB: DCB;
    FUpperScroll: boolean;
    FStart, FEnd, FWelcomeUpperRow, FWelcomeLowRow, FTopCommand, FUpperCommand, FLowCommand,
        FTopLine, FUpperLine, FLowerLine, FScrollCommand, FHorzontalScrollMode: AnsiString;
    FRowLength: integer;
    FTimeOuts: COMMTIMEOUTS;
    procedure SetDisplayerDefault(Value: TDisplayerType);
    function  WriteCom(SendString: AnsiString): boolean;
    procedure ApplyComSetup;
    procedure SetUpperScroll(Value: boolean);
    procedure SetRowLength(Value: integer);
    procedure SetUniName(Value: AnsiString);
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Initial;             //20120320 modi by 3188修正CD8240可與TP3688共用
    procedure SetDisplayerCommand; //20120320 modi by 3188修正CD8240可與TP3688共用
    function Connect: boolean;
    function Connected: boolean;
    function DisConnect: boolean;
    function DisplayString2(UpperRow, LowRow: AnsiString): boolean;
    function DisplayString3(TopRow, UpperRow, LowRow: AnsiString): boolean;
    function ShowLine(line: integer; UpperRow: AnsiString): boolean;
    function ShowLine0(TopRow:   AnsiString): boolean;
    function ShowLine1(UpperRow: AnsiString): boolean;
    function ShowLine2(LowRow:   AnsiString): boolean;
    function ScrollString(UpperRow: AnsiString): boolean;
    function NoCommand(Value: AnsiString): boolean;
    function DoUpperScroll(Value: boolean): boolean;
    function ClearAll: boolean;
  published
    property Port;
    property PortHandle: THandle read FHandle write FHandle;  //20120320 modi by 3188修正CD8240可與TP3688共用
    property Delay;
    property CheckDSR;
    property WTimeOut;
    property BaudRate;
    property DataBits;
    property StopBits;
    property ParityCheck;
    property InbondBuffer;
    property OutbondBuffer;
    property UniName;
    property Display: TDisplayerType read FDisplayer write SetDisplayerDefault;
    property WelcomeUpperRow: AnsiString read FWelcomeUpperRow write FWelcomeUpperRow;
    property WelcomeLowRow: AnsiString read FWelcomeLowRow write FWelcomeLowRow;
    property UpperScroll: boolean read FUpperScroll write SetUpperScroll;
    property RowLength: integer read FRowLength write SetRowLength;
  end;

  //THandyScan

  TOnGetDataEvent = procedure(Sender: TObject; Data: AnsiString; DataLength: integer) of object;
  THandyScan = class(TBaseCOM)
  private
    FComRcv: TPOSComm;
    FHandyScanner: TBarCodeScannerType;
    FOnGetData: TOnGetDataEvent;
    FReadTimeOut, FInterval, FEndLength: integer;
    procedure SetHandyScannerDefault(Value: TBarCodeScannerType);
    procedure SetReadTimeOut(Value: integer);
    procedure SetInterval(Value: integer);
    procedure ApplyComSetup;
    procedure ReceiveData(Sender: TObject; Buffer: AnsiString{Pointer}; BufferLength: Word);
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function Connect: boolean;
    function Connected: boolean;
    function DisConnect: boolean;
    function CheckDSRStatus: boolean;
  published
    property Port;
    property Delay;
    property CheckDSR;
    property WTimeOut;
    property BaudRate;
    property DataBits;
    property StopBits;
    property ParityCheck;
    property InbondBuffer;
    property OutbondBuffer;
    property HandyScanner: TBarCodeScannerType read FHandyScanner write SetHandyScannerDefault;
    property RTimeOut: integer read FReadTimeOut write SetReadTimeOut;
    property Interval: integer read FInterval write SetInterval;
    property EndLength: integer read FEndLength write FEndLength;
    property OnGetData: TOnGetDataEvent read FOnGetData write FOnGetData;
  end;

implementation

function Inp32(wAddr: word): byte; stdcall; external 'inpout32.dll';

function iif(xCheck: Boolean; xValue1, xValue2: Variant): Variant;
begin
  if xCheck then
    Result:= xValue1
  else
    Result:= xValue2;
end;


//TInvPrinter

constructor TInvPrinter.Create(AOwner: TComponent);
begin
    inherited Create(AOwner);
    FDefaultReset := true;
    FRetry  := true;
    FHandle := 0;
    FErrorMsg :=  '發票機有問題!請檢查!!';
    FOutputMode := om_Printer;
    FPrintStringList := TStringList.Create;
    FConnectMode := cm_SerialPort;
    FPrinter := TP_3688;
    SetPrinter(FPrinter);
end;

destructor TInvPrinter.Destroy;
begin
    FPrintStringList.Free();
    inherited Destroy;
end;

procedure TInvPrinter.Assign(Source: TPersistent);
begin
  try
    if Source is TInvPrinter then
    begin
      //置頂
      ConnectType     := TInvPrinter(Source).ConnectType;
      //其他隨意
      FHandle         := TInvPrinter(Source).PortHandle;
      PrinterType     := TInvPrinter(Source).PrinterType;
      Port            := TInvPrinter(Source).Port;
      Delay           := TInvPrinter(Source).Delay;
      CheckDSR        := TInvPrinter(Source).CheckDSR;
      CheckLPTStatus  := TInvPrinter(Source).CheckLPTStatus;
      WTimeOut        := TInvPrinter(Source).WTimeOut;
      BaudRate        := TInvPrinter(Source).BaudRate;
      DataBits        := TInvPrinter(Source).DataBits;
      StopBits        := TInvPrinter(Source).StopBits;
      ParityCheck     := TInvPrinter(Source).ParityCheck;
      InbondBuffer    := TInvPrinter(Source).InbondBuffer;
      OutbondBuffer   := TInvPrinter(Source).OutbondBuffer;
      ErrorMsg        := TInvPrinter(Source).ErrorMsg;
      Direction       := TInvPrinter(Source).Direction;
      OutPutMode      := TInvPrinter(Source).OutPutMode;
      //ConnectType     := TInvPrinter(Source).ConnectType;
      UniName         := TInvPrinter(Source).UniName;
      DefaultReset    := TInvPrinter(Source).DefaultReset;
      Direction       := TInvPrinter(Source).Direction;
//      PrinterType     := TInvPrinter(Source).PrinterType;
//      OutPutMode      := TInvPrinter(Source).OutPutMode;
//      ConnectType     := TInvPrinter(Source).ConnectType;
      PrinterIndex    := TInvPrinter(Source).PrinterIndex;
      OnPtrError      := TInvPrinter(Source).OnPtrError;
      exit;
    end;
  except
    exit;
  end;
  inherited Assign(Source);
end;


procedure TInvPrinter.StrConvert(str1: AnsiString; var str2: AnsiString);
var
    len, i, j, ischinese: integer;
begin
    len := 0;
    i := 0;
    j := 0;
    ischinese := 0;

    len := Length(str1);
    for i := 0 to len - 1 do
    begin
        if (ord(str1[i]) < 0) then
        begin
            if (ischinese = 0) then
            begin
                if (ord(str1[i + 1]) < 0) then
                begin
                    ischinese := ischinese + 1;
                    str2[j] := char(28); j := j + 1;
                    str2[j] := '&';      j := j + 1;
                end
                else
                    str1[i] := ' ';

                str2[j] := str1[i]; j := j + 1;
            end
            else
            begin
                if (ord(str1[i + 1]) >= 0) then
                begin
                    if (ischinese mod 2 <> 0) then
                    begin
                        ischinese := 0;
                        str2[j] := str1[i];     j := j + 1;
                        str2[j] := char(28);    j := j + 1;
                        str2[j] := '.';         j := j + 1;
                    end
                    else                   // 前段已是偶數中文
                    begin
                        ischinese := 0;
                        str2[j] := char(28); j := j + 1; //先結束字元放棄
                        str2[j] := '.';      j := j + 1;
                        str2[j] := ' ';      j := j + 1;
                    end;
                end
                else
                begin
                    str2[j] := str1[i]; j := j + 1;
                    ischinese := ischinese + 1;
                end;
            end;
        end
        else
        begin
            if (ischinese > 0) then
            begin
                ischinese := 0;
                str2[j] := char(28); j := j + 1;
                str2[j] := '.';      j := j + 1;
            end;
            str2[j] := str1[i]; j := j + 1;
        end;
    end;
    str2[j] := char(0);
end;

function TInvPrinter.GetPlatForm: TWindowPlat;
begin
    case Win32Platform of
        VER_PLATFORM_WIN32s         : Result := wp_Win32s;
        VER_PLATFORM_WIN32_WINDOWS  : Result := wp_Win98;
        VER_PLATFORM_WIN32_NT       : Result := wp_WinNT;
    end;
end;

procedure TInvPrinter.SetOutPutMode(Value: TOutPutMode);
begin
    FPrintStringList.Clear();
    FOutPutMode := Value;
end;

procedure TInvPrinter.SetConnectMode(Value: TConnectMode);
begin
    FConnectMode := Value;
    SetPrinter(FPrinter);
end;

procedure TInvPrinter.SetDirection(Value: TPtrDir);
begin
    //Both(收執聯與存根聯同時列印)
    //Journal(存根聯列印)
    //Receitp(收執聯列印)
    FPtrDir := Value;
    SetDefaultCommand;
end;

procedure TInvPrinter.SetUniName(Value: AnsiString);
begin
    inherited;
    if (Connected) then
        ApplySetup;
end;

procedure TInvPrinter.SetPrinterIndex(Value: integer);
begin
    try
        Printers.Printer.PrinterIndex := Value;
    except
    end;
end;

function TInvPrinter.GetPrinters: TStrings;
begin
    Result := Printers.Printer.Printers;
end;

function TInvPrinter.GetPrintStringList: TStringList;
begin
    Result := FPrintStringList;
end;

function TInvPrinter.GetPrinterIndex: integer;
begin
    Result := -1;
    try
        Result := Printers.Printer.PrinterIndex;
    except
    end;
end;

function TInvPrinter.Connected: boolean;
begin
    //傳回目前連線狀態
    if ((FConnectMode <> cm_PrintServer) and (FOutPutMode = om_Printer) and
        (FPrinter <> WindowsPrinter)) then
        result := (FHandle > 0)
    else
        result := true;
end;

function TInvPrinter.Connect: boolean;
begin
    //連接
    SetActive(true);
    result := FActive;
end;

function TInvPrinter.DisConnect: boolean;
begin
    //中斷連線
    CloseHandle(FHandle);
    FActive := false;
    FHandle := 0;
    result := true;
end;

function TInvPrinter.FF: boolean;
begin
    //發票換頁並切紙
    if (Connected) then
    begin
        Result := WriteString(FStart + FFormFeed  + FDir + FEnd, ' ');
    end
    else
        Result := false;
end;

function TInvPrinter.LF(Lines: integer): boolean;
var
    i: integer;
begin
    //Lines:欲空白列數
    Result := false;
    if (Connected) then
        for i := 0 to Lines - 1 do
            if (FPrinter = TP_3688) then
                Result := WriteString(FStart + FLineFeed + FDir + '1' + FEnd, ' ')
            else
                Result := WriteString(FStart + FLineFeed + FDir + FEnd, ' ')
    else
        Result := false;
end;


function TInvPrinter.Logo: boolean;
begin
  //蓋店名章
  if (Connected) then
  begin
    SendComString(#28#112#1#0);
    Result := True;
  end else
    Result := false;
end;

function TInvPrinter.Stamp: boolean;
begin
    //蓋店名章
    if (Connected) then
        Result := WriteString(FStart + FStamp + FEnd, '')
    else
        Result := false;
end;

function TInvPrinter.CutPaper: boolean;
begin
    //立即切紙
    if (Connected) then
        Result := WriteString(FStart + FCutPaper + FEnd, '')
    else
        Result := false;
end;

function TInvPrinter.OpenCashDrawer: boolean;
begin
    //開錢櫃
    if (Connected) then
        Result := WriteString(FCashDrawer, '')
    else
        Result := false;
end;

function TInvPrinter.Valid(SendString: AnsiString): boolean;
begin
    //發票機認證
    if (Connected) then
        Result := WriteString(FStart + FValidate + SendString + FEnd + FOtherEnd, SendString)
    else
        Result := false;
end;

procedure TInvPrinter.BeginDoc;
begin
    try
        Printers.Printer.BeginDoc;
    except
    end;
    FLineCount := 0;
end;

procedure TInvPrinter.EndDoc;
begin
    try
        Printers.Printer.EndDoc;
    except;
    end;
end;

procedure TInvPrinter.EndPageMode;
begin
  //中斷PageMode
  SendComString(#12);
end;

function TInvPrinter.SendCOMString(SendString: AnsiString): boolean;
begin
    //送出字串(無其它指令)
    if (Connected) then
        Result := WriteString(SendString, SendString)
    else
        Result := false;
end;

function TInvPrinter.SendCOMUTF8String(SendString: AnsiString): boolean;
begin
    //送出字串
    //送出字串(無其它指令)
    if (Connected) then
        Result := WriteUTF8String(SendString, SendString)
    else
        Result := false;
end;

function TInvPrinter.SendString(SendString: AnsiString): boolean;
var
    RetString: AnsiString;
begin
    //送出字串
    if (Connected) then
    begin
        if (FPrinter = EPSON_TMU300) then
        begin
            setlength(RetString,32);   //20150327 add by 04707 for C30-20150326003
            StrConvert(SendString, RetString);
            Result := WriteString(FStart + FLine + FDir + RetString + FEnd + FOtherEnd, SendString);
        end
        else
        if (FPrinter = EPSON_TMU295) then
            Result := WriteString(StrSpace(SendString,-35), SendString)
        else
        if (FPrinter = TP_3688) then
            Result := WriteString(FStart + FLine + FDir + '1' + SendString + FEnd + FOtherEnd, SendString)
        else
            Result := WriteString(FStart + FLine + FDir + SendString + FEnd + FOtherEnd, SendString);
    end
    else
        Result := false;
end;

//列印條碼
function TInvPrinter.PrintBarcode(xBarcode: AnsiString): boolean;
begin
//設定基本列印命令
  case (FPrinter) of
    EPSON_TMT882, WP_T810, RP_600, RP_600_V9, PP8000, EPSON_TM88IV, EPSON_TMT70:
      begin
        //GS h n #29 #104, n //條碼高度
        //GS w n #29 #119, n //條碼寬度 def: 3   Max words 17
        //GS k m d1...dk #29 #107 #69(39code)    //列印條碼
        xBarcode := copy(xBarcode, 1, 17);
        Result := SendCOMString(#29#104#64#13#29#119#3#13#29#107#69 + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+Ansichar(13));
        Result := SendString(xBarcode);
      end;
  end;
end;


function TInvPrinter.PrintBarcode2(xBarcode: AnsiString): boolean;
begin
  //設定基本列印命令
  xBarcode := copy(xBarcode, 1, 19);
  case (FPrinter) of
    WP_T810:
    begin
      //GS h n #29 #104, n //條碼高度
      //GS w n #29 #119, n //條碼寬度 def: 3   Max words 17
      //GS k m d1...dk #29 #107 #69(39code)    //列印條碼
      SendCOMString(#29#33#0#13#27#97#1#13);               //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#64#13);           //寬度、高度 //64
      Result := SendCOMString(#29#107#69 + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+Ansichar(13));
    end ;

    RP_600, RP_600_V9:
    begin
      //GS h n #29 #104, n //條碼高度
      //GS w n #29 #119, n //條碼寬度 def: 3   Max words 17
      //GS k m d1...dk #29 #107 #69(39code)    //列印條碼
      SendCOMString(#29#33#0#13);              //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#60#13);           //寬度、高度
      Result := SendCOMString(Ansichar(29) + Ansichar(107) + Ansichar(69) + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+ Ansichar(12));
    end ;

    PP8000, EPSON_TM88IV, EPSON_TMT70:
    begin
      //GS h n #29 #104, n //條碼高度
      //GS w n #29 #119, n //條碼寬度 def: 3   Max words 17
      //GS k m d1...dk #29 #107 #69(39code)    //列印條碼
      SendCOMString(#29#33#0#13);              //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#49#13);           //寬度、高度
      Result := SendCOMString(Ansichar(29) + Ansichar(107) + Ansichar(69) + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+ Ansichar(12));
    end ;

    EPSON_TMT882:
    begin
      SendCOMString(#29#33#1#13#27#97#1#13);               //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#70#13);           //寬度、高度
      Result := SendCOMString(#29#107#69 + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+Ansichar(13));
      SendCOMString(#29#33#0);
    end;
  end;
end;

function TInvPrinter.PrintBarcode3(xBarcode: AnsiString): boolean;
begin
  //設定基本列印命令
  xBarcode := copy(xBarcode, 1, 19);
  case (FPrinter) of
    WP_T810:
    begin
      //GS h n #29 #104, n //條碼高度
      //GS w n #29 #119, n //條碼寬度 def: 3   Max words 17
      //GS k m d1...dk #29 #107 #69(39code)    //列印條碼
      SendCOMString(#29#33#0#13#27#97#1#13);               //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#40#13);           //寬度、高度 //64
      Result := SendCOMString(#29#107#69 + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+Ansichar(13));
    end ;

    RP_600, RP_600_V9, PP8000, EPSON_TM88IV, EPSON_TMT70:
    begin
      //GS h n #29 #104, n //條碼高度
      //GS w n #29 #119, n //條碼寬度 def: 3   Max words 17
      //GS k m d1...dk #29 #107 #69(39code)    //列印條碼
//      SendCOMString(#29#33#0#13);              //放大字體  //對其方式
//      SendCOMString(#29#119#1#13#29#104#60#13);           //寬度、高度
//      Result := SendCOMString(Ansichar(29) + Ansichar(107) + Ansichar(69) + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+ Ansichar(12));
      SendCOMString(#29#33#0#13);                         //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#40#13);           //寬度、高度
      Result := SendCOMString(Ansichar(29) + Ansichar(107) + Ansichar(69) + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode));
    end ;

    EPSON_TMT882:
    begin
      SendCOMString(#29#33#1#13#27#97#1#13);               //放大字體  //對其方式
      SendCOMString(#29#119#1#13#29#104#70#13);           //寬度、高度
      Result := SendCOMString(#29#107#69 + Ansichar(Length(xBarcode)) + pAnsichar(xBarcode)+Ansichar(13));
      SendCOMString(#29#33#0);
    end;
  end;
end;

function TInvPrinter.PrintQRcode(xBarcode: AnsiString;
  xBackHeight: integer): boolean;
var
  s_lngQRLength,m_lngQRLength : integer;
begin
  //QrCode ..由於廠商沒給使用說明書。外部也找不著。目前找不到規則..
  case (FPrinter) of
    WP_T810:
    begin
      s_lngQRLength := Length(xBarcode) + 3  mod 256 ;
      m_lngQRLength := strtoint(formatfloat('0',(Length(xBarcode) + 3)  / 256)) ;

      SendCOMString(#29#33#0#27#97#2);                                                 //放大字體  //對其方式
      SendCOMString(#27#75 + Ansichar (xBackHeight) + #13);                            //倒退長度
      SendCOMString(#29#40#107#4#0#49#65#50#0);                                        //Code Model
      SendCOMString(#29#40#107#3#0#49#67#3);                                           //Code Module Size
      SendCOMString(#29#40#107#3#0#49#68#48);                                          //Err. Correction Lv.
      SendCOMString(#29#40#107+ Ansichar(s_lngQRLength)+ Ansichar(m_lngQRLength)+ #49#80#48);  //QR Content Length
      SendCOMString(pAnsichar(xBarcode));                                              //const..
      SendCOMString(#29#40#107#3#0#49#81#48);                                          //Print QR Image
      SendCOMString(#27#97#0#13);                                                      //靠左對齊
    end;

    RP_600, PP8000, EPSON_TM88IV, EPSON_TMT70:
    begin
      s_lngQRLength := Length(xBarcode) + 3  mod 256 ;
      m_lngQRLength := strtoint(formatfloat('0',(Length(xBarcode) + 3)  / 256)) ;

      SendCOMString(#29#33#0#27#97#2);
      SendCOMString(#29#40#107#4#0#49#65#50#0);                                        //Code Model
      SendCOMString(#29#40#107#3#0#49#67#3);                                           //Code Module Size
      SendCOMString(#29#40#107#3#0#49#68#48);                                          //Err. Correction Lv.
      SendCOMString(#29#40#107+ Ansichar(s_lngQRLength)+ Ansichar(m_lngQRLength)+ #49#80#48);  //QR Content Length
      SendCOMString(pAnsichar(xBarcode));                                              //const..
      SendCOMString(#29#40#107#3#0#49#81#48);                                          //Print QR Image
    end;
  end;
end;

function TInvPrinter.PrintQRcode2014(xBarcode: AnsiString;
  xBackHeight: integer;xOffset:integer=2;xLv:Integer=6): boolean;    //20180903 add by 02953 for Cust-20180903001
var
  s_lngQRLength,m_lngQRLength : integer;
  mxBarcode :UTF8String;
begin
  //QrCode ..由於廠商沒給使用說明書。外部也找不著。目前找不到規則..
  mxBarcode :=  UTF8Encode(xBarcode);
  s_lngQRLength := (Length(mxBarcode)+3 )  mod 256 ;
  //m_lngQRLength := strtoint(formatfloat('0',(Length(mxBarcode) + 3)  / 256)) ;  //20161108 mark by 02953 for Cust-20161108001
  m_lngQRLength := (Length(mxBarcode)+3 )  div 256 ;                              //20161108 modi by 02953 for Cust-20161108001

  case (FPrinter) of
      WP_T810:
      begin
        //SendCOMString(#27#97+Ansichar(xOffset));    //20180907 mark by 07113
        SendCOMString(#27#97+Ansichar(xOffset)+ #13); //20180907 modi by 07113
        if xBackHeight <= 90 then
        begin
          SendCOMString(#27#75 + Ansichar (xBackHeight) + #13);                            //倒退長度  //列印針頭回推30為一行
        end else
        begin
          SendCOMString(#27#75 + Ansichar(90) + #13);
          SendCOMString(#27#75 + Ansichar(xBackHeight-90) + #13);
        end;
        //SendCOMString(#29#40#107#3#0#49#118#6);                                          //lv6  //20180903 mark by 02953 for Cust-20180903001
        SendCOMString(#29#40#107#3#0#49#118+Ansichar(xLv));                              //lv6    //20180903 modi by 02953 for Cust-20180903001
        SendCOMString(#29#40#107#4#0#49#65#50#0);                                        //Code Model
        SendCOMString(#29#40#107#3#0#49#67#3);                                           //Code Module Size
        SendCOMString(#29#40#107#3#0#49#69#48);                                          //Err. Correction Lv.
        //SendCOMString(#29#40#107+ Ansichar(s_lngQRLength)+ Ansichar(m_lngQRLength)+ #49#80#48);  //QR Content Length                        //20161108 mark by 02953 for Cust-20161108001
        SendCOMString(#29#40#107+ AnsiString(Ansichar(s_lngQRLength))+ AnsiString(Ansichar(m_lngQRLength))+ #49#80#48);  //QR Content Length  //20161108 modi by 02953 for Cust-20161108001
        SendCOMString(pAnsichar(mxBarcode));                                                                              //20161108 modi by 02953 for Cust-20161108001
        //SendCOMUTF8String(pAnsichar(xBarcode));                                              //const..                  //20161108 mark by 02953 for Cust-20161108001
        //SendCOMString(pAnsichar(xBarcode));                                              //const..
        SendCOMString(#29#40#107#3#0#49#81#48);                                          //Print QR Image
        SendCOMString(#27#97#0#13);                                     //靠左對齊 0是靠左
      end;

    RP_600_V9: //20151112 add by 04707
      begin
        //SendCOMString(#29#40#107#3#0#49#118#6);                                          //lv6  //20180903 mark by 02953 for Cust-20180903001
        SendCOMString(#29#40#107#3#0#49#118+Ansichar(xLv));                              //lv6    //20180903 modi by 02953 for Cust-20180903001
        SendCOMString(#29#40#107#4#0#49#65#50#0);                                        //Code Model
        SendCOMString(#29#40#107#3#0#49#67#2);                                           //Code Module Size dot
        SendCOMString(#29#40#107#3#0#49#69#48);                                          //Err. Correction Lv.
        SendCOMString(#29#40#107#147#1#49#80#48);                                        //固定長度400+3=256+144+3=403
        SendCOMUTF8String(pAnsichar(xBarcode));                                          //const..
        SendCOMString(#29#40#107#3#0#49#81#48);                                          //Print QR Image
        if xOffset = 2 then
          SendCOMString(#12);    //end page mode
      end;

    RP_600, PP8000, EPSON_TM88IV, EPSON_TMT70:
      begin
        //SendCOMString(#29#40#107#3#0#49#118#6);                                          //lv6  //20180903 mark by 02953 for Cust-20180903001
        SendCOMString(#29#40#107#3#0#49#118+Ansichar(xLv));                              //lv6    //20180903 modi by 02953 for Cust-20180903001
        SendCOMString(#29#40#107#4#0#49#65#50#0);                                        //Code Model
        if FPrinter in [EPSON_TM88IV, PP8000, EPSON_TMT70] then
          SendCOMString(#29#40#107#3#0#49#67#4)                                          //Code Module Size dot
        else
          SendCOMString(#29#40#107#3#0#49#67#3);                                         //Code Module Size dot
        SendCOMString(#29#40#107#3#0#49#69#48);                                          //Err. Correction Lv.
        //SendCOMString(#29#40#107+ Ansichar(s_lngQRLength)+ Ansichar(m_lngQRLength)+ #49#80#48);  //QR Content Length    //20161108 mark by 02953 for Cust-20161108001
        SendCOMString(#29#40#107+ AnsiString(Ansichar(s_lngQRLength))+ AnsiString(Ansichar(m_lngQRLength))+ #49#80#48);  //QR Content Length  //20161108 modi by 02953 for Cust-20161108001
        //SendCOMUTF8String(pAnsichar(xBarcode));                                         //const..                       //20161108 mark by 02953 for Cust-20161108001
        SendCOMString(pAnsichar(mxBarcode));                                                                              //20161108 modi by 02953 for Cust-20161108001
        SendCOMString(#29#40#107#3#0#49#81#48);                                          //Print QR Image

        if xOffset = 2 then
          SendCOMString(#12);    //end page mode
      end;
  end;
end;

function TInvPrinter.IsOnLine: boolean;
var
    ErrCount: integer;
    Port: Word;
    data: Byte;
    State: DWORD;

begin
    //印表機是否連線
    //mark by 01753
//    if (not FCheckLPTStatus) or (not(csDesigning in ComponentState)) then
//    begin
//        Result := true;
//        exit;
//    end;
    //mark by 01753
    try
        ErrCount := 1;
        Port := $379;
//        FWinPlatForm := GetPlatForm;
        if ((FOutPutMode = om_Screen) or (FPrinter = WindowsPrinter)) then
        begin
            result := true;
            exit;
        end;

        while (true) do
        begin
            Result := true;
            // On Line  : $df or ((data and $08) = $08)
            // Power off: $f7 or 7F
            data := 0;
            if (FConnectMode = cm_ParallelPort) and FCheckLPTStatus then
            begin
                ErrCount := 0;
                try
//                    if ZlIOStarted then
//                    begin
                      //while (ErrCount < 3) do
                      //begin
                          //有一行被我刪掉了= =~
                          data := Inp32($379);

                          if ( ( //((data and $40) = $40) and     //沒必要的判斷
                                 //((data and $20) =  $0) and     //出紙   低電位表示沒問題
                                 ((data and $10) = $10) ) or      //Select 判斷有無ONLINE
                                 (data = 6) ) then                //不懂 保留
                          begin
                              Result := true;
                              break;
                          end else
                          begin
  //                        if ((data and $20) <>  $0) then
  //                        begin
  //                          FErrorMsg := '請檢查紙張是否安裝妥當。';
  //                        end else
  //                        if ((data and $08) <> $08) then
  //                        begin
  //                          FErrorMsg := '印表機出現異常訊息。';
  //                        end else
                            if  ((data and $10) <> $10) then
                              FErrorMsg := '設備沒有啟動完整。';
  //                        end;

                            Result := false;
                          end;
                          ErrCount := ErrCount + 1;
                      //end;
//                    end  else
//                    begin
//                      Result := false;
//                      FErrorMsg := '裝置已被佔用，請檢查印表機清單或請重開電腦。';
//                    end;

//                    if ZlIOStarted then
//                        zliostop;



                except
                  Result := false;
                end;
            end
            else
            if ((FConnectMode = cm_SerialPort) and (UniName = '') and (CheckDSR)) then
            begin
                ErrCount := 0;
                while (ErrCount < 3) do
                begin
                    if (GetCommModemStatus(FHandle, State)) then
                        Result := ((State and MS_DSR_ON) <> 0)
                    else
                        Result := false;

                    if Result then exit;

                    ErrCount := ErrCount + 1;
                end;
            end;

            if not Result then
            begin
              if FRetry then  //20131029 modi by 01753
                if Assigned(FOnPtrError) then
                  FOnPtrError(self, FErrorMsg, data, FRetry);

                if not FRetry then
                begin
                    Result := false;
                    break;
                end;
            end
            else
            begin
                Result := true;
                break;
            end;
        end;
    except
        Result := false;
    end;
end;

procedure TInvPrinter.SetActive(Value: boolean);
var
  mPort: PChar;
begin
    //開啟設備
    if not(csDesigning in ComponentState) then
    begin
        if (Value) then
        begin
            if (( FConnectMode <> cm_PrintServer) and (FPrinter <> WindowsPrinter)) then
            begin
                if UniName = '' then
                  mPort := DevicePortStrValue[Port]
                else
                  mPort := PChar(UniName);

                case (FConnectMode) of
                    (cm_SerialPort):
                        begin
                            if (Connected) then
                                DisConnect;

            	            FHandle := CreateFile(mPort, GENERIC_READ or GENERIC_WRITE, 0, 0,
                                                OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
                        end;

                    (cm_ParallelPort):
                        begin
                          if Connected then
                            DisConnect;

                          if IsOnLine then
              	            FHandle := CreateFile(mPort, GENERIC_WRITE, 0, 0,
                                        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
                        end;
                   else
                     if (Connected) then
                       DisConnect;
                end;

              ApplySetup;
              FActive := FHandle <> INVALID_HANDLE_VALUE;
              if FDefaultReset then
              begin
                if FActive then 
                begin
                  FActive := WriteString(char(13), '');
                  if FActive then
                    Reset;
                end;
              end;
            end
            else //使用PrintServer
                FActive := true;
        end
        else
        begin
            CloseHandle(FHandle);
            FActive := false;
            FHandle := 0;
        end
    end;
end;

procedure TInvPrinter.ApplySetup;
begin
    //套用設定
    if (FConnectMode = cm_SerialPort) then
    begin
        SetupComm(FHandle, InbondBuffer, OutbondBuffer);
       	GetCommState(FHandle, FDCB);
        FDCB.BaudRate := BaudRateValue[BaudRate];
        FDCB.ByteSize := DataBitsValue[DataBits];
        FDCB.Parity   := ParityCheckValue[ParityCheck];
        FDCB.StopBits := StopBitsValue[StopBits];
        FDCB.XonChar  := char(17);
        FDCB.XoffChar := char(19);
      	SetCommState(FHandle, FDCB);
    end;

	SetDefaultCommand();
  FTimeOuts.ReadIntervalTimeout         := MAXDWORD;
  FTimeOuts.ReadTotalTimeoutMultiplier  := 3;
  FTimeOuts.ReadTotalTimeoutConstant    := 100;
  FTimeOuts.WriteTotalTimeoutMultiplier := 3;
  FTimeOuts.WriteTotalTimeoutConstant   := WTimeOut;
	SetCommTimeouts(FHandle, FTimeOuts);
end;

procedure TInvPrinter.SetDefaultCommand;
begin
    //設定基本列印命令
    case (FPrinter) of
        TP_3688:
            begin
                FStart := #27 + #27;
                FEnd := #13;
                FReset := 'R';
                FLine := 'P';
                FLineFeed := 'L';
                FFormFeed := 'V';
                FStamp := 'S';
                FCutPaper := 'C';
                FValidate := FLine + 'V1';
                FCashDrawer := #27 + #27 + 'G' + #13;
                FOtherEnd := '';
                case (FPtrDir) of
                    pd_Receipt: FDir := 'R';
                    pd_Journal: FDir := 'J';
                  else
                     FDir := 'B';
                end;
            end;

        EPSON_RPU420:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + #64;
                FLine := '';
                FLineFeed := #10;
                FFormFeed := #12;
                FStamp := #27 + #111;
                FCutPaper := #29 + #86 + #48;
                //FValidate := FLine + 'V1';
                FValidate := #27 + 'c0' + #8;
                FCashDrawer := #27 + 'p' + #0 + #100+ #250;
                FOtherEnd := #10;
                case (FPtrDir) of
                    pd_Receipt: FDir := #27 + 'c0' + #2;
                    pd_Journal: FDir := #27 + 'c0' + #1;
                  else
                    FDir := #27 + 'z' + #1;
                    //FDir := #27 + 'c0' + #3;
                end;
            end;

        NIXDORF_77,
        NIXDORF_210:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := '';
                FLine := '';
                FStamp := '';
                FCutPaper := #27 + 'm';
                FValidate := '';
                FCashDrawer := #27 + 'p' + #0 + #100 + #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        NIXDORF_20:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := '';
                FLine := '';
                FStamp := '';
                FCutPaper := #27 + 'm';
                FValidate := '';
                FCashDrawer := '';
                FOtherEnd := '';
                FDir := '';
            end;

        STAR_SP320:
            begin
                FStart := '';
                FEnd := #13;
                FReset := '';
                FLineFeed := #10;
                FFormFeed := #12;
                FLine := '';
                FStamp := '';
                FCutPaper := #29 + 'V' + #1;
                FValidate := '';
                FCashDrawer := #7;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        NormalType:
            begin
                FStart := '';
                FEnd := #13;
                FReset := '';
                FLineFeed := #10;
                FFormFeed := #12;
                FLine := '';
                FStamp := '';
                FCutPaper := #29 + 'V' + #1;
                FValidate := '';
                FCashDrawer := #27 + 'p' + #48 + #16+ #48;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        EPSON_TM200:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FCutPaper := #29+ 'V' + #1;
                FFormFeed := #$1c + #$28 + #$4c +
                             #$02 + #0 + #$42 + #49 + FCutPaper;
                FLine := '';
                FStamp := '';
                FValidate := '';
                FCashDrawer := #27 + 'p' + #0 + #100+ #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        EPSON_TMT882:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := #29 + 'V' + #1;
                FLine := '';
                FStamp := '';
                FCutPaper := #29 + 'V1';
                FValidate := '';
                FCashDrawer := #27 + 'p' + #0 + #100 + #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        EPSON_TMU295:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@' + #27 + 'c30' + #27 + 'c40' + #27 + 'q';
                FLineFeed := #10;
                FFormFeed := '';
                FLine := '';
                FStamp := '';
                FCutPaper := '';
                FValidate := '';
                FCashDrawer := '';
                FOtherEnd := '';
                FDir := '';
            end;

        IBM_4610:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := #12;
                FLine := '';
                FStamp := '';
                FCutPaper := #29 + 'V1';
                FValidate := '';
                FCashDrawer := '';
                FOtherEnd := '';
                FDir := '';
            end;
            
        EPSON_TMU300:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := #12;
                FLine := '';
                FStamp := '';
                FCutPaper := #27 + 'i';
                FValidate := '';
                FCashDrawer := #27 + 'p' + #0 + #100+ #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        EPSON_TMU675:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := #29 + 'V' + #1;
                FLine := '';
                FStamp := '';
                FCutPaper := #29 + 'V' + #1;
                FValidate := #27 + 'c0' + #4;
                FCashDrawer := #27 + 'p' + #0 + #100 + #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        EPSON_MD332S:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #10;
                FFormFeed := #12;
                FLine := '';
                FStamp := '';
                FCutPaper := '';
                FValidate := '';
                FCashDrawer := #27 + 'p';
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        WindowsPrinter:
            begin
                FStart := '';
                FEnd := '';
                FReset := '';
                FLineFeed := '';
                FFormFeed := '';
                FLine := '';
                FStamp := '';
                FCutPaper := '';
                FValidate := '';
                FCashDrawer := '';
                FOtherEnd := '';
                FDir := '';
            end;

        WP_T810:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #27#100#1;
                FFormFeed := #29 + 'V' + #1;
                FLine := '';
                FStamp := '';
                FCutPaper := #13#29#86#66#16;
                FValidate := '';
                FCashDrawer := #27 + 'p' + #0 + #100 + #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;

        RP_600, RP_600_V9, PP8000, EPSON_TM88IV, EPSON_TMT70:
            begin
                FStart := '';
                FEnd := #13;
                FReset := #27 + '@';
                FLineFeed := #27#100#1;
                FFormFeed := #29 + 'V' + #1;
                FLine := '';
                FStamp := '';
                FCutPaper := #13#29#86#66#16;
                FValidate := '';
                FCashDrawer := #27 + 'p' + #0 + #100 + #250;
                FOtherEnd := FLineFeed + #13;
                FDir := '';
            end;
    end;
end;

procedure TInvPrinter.SetPrinter(Value: TPrinterType);
begin
    //設定原始通訊設定值
    FPrinter := Value;
    case (FPrinter) of
        TP_3688:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
            end;

        NIXDORF_77, NIXDORF_20:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_OddParity;
            end;

        STAR_SP320, NormalType:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
            end;

        EPSON_TM200,
        EPSON_TMT882,
        EPSON_TMU295,
        EPSON_TMU300,
        EPSON_RPU420,
        EPSON_MD332S,
        WP_T810,
        RP_600,
        RP_600_V9,
        PP8000,
        EPSON_TM88IV,
        EPSON_TMT70:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
            end;

        NIXDORF_210, EPSON_TMU675:
            begin
                BaudRate := brt_19200;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
            end;

        IBM_4610:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
            end;

        WindowsPrinter:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
            end;
    end;
end;

function TInvPrinter.WriteString(SendString, OrgString: AnsiString): boolean;
var
    numWrite, State: DWORD;
    sndStr: PChar; //20131025
begin
    //將字串送到指定通訊埠
    Result := false;
    sndStr := pchar(SendString);    //20131025
    if (FOutPutMode = om_Printer) then
    begin
        Sleep(Delay);
        //if ((FConnectMode = cm_SerialPort){ and (UniName = '')} and (CheckDSR)) then
        //begin
        //    GetCommModemStatus(FHandle, State);
        //    Result := IsOnLine;
        //end
        //else
        //    Result := IsOnLine;
        Result := IsOnLine;

        if ((Result) and (FPrinter <> WindowsPrinter)) then
            Result := WriteFile(FHandle, sndStr^, Length(SendString), numWrite, nil);
    end
    else
    if (FOutPutMode = om_Screen) then
    begin
        if (OrgString <> '') then
            FPrintStringList.Add(OrgString);
        Result := true;
    end;
end;

function TInvPrinter.WriteUTF8String(SendString, OrgString: AnsiString): boolean;
var
    numWrite, State: DWORD;
    sndStr: pchar; //20131025
    mUTF8SendString :UTF8String;
begin
    //將字串送到指定通訊埠
    Result := false;
    //sndStr := pAnsichar(SendString);
    mUTF8SendString :=  UTF8Encode(SendString);
    sndStr := pchar(mUTF8SendString);      //20131025

    if (FOutPutMode = om_Printer) then
    begin
        Sleep(Delay);

        Result := IsOnLine;

        if ((Result) and (FPrinter <> WindowsPrinter)) then
          Result := WriteFile(FHandle, sndStr^, Length(mUTF8SendString), numWrite, nil);
    end
    else
    if (FOutPutMode = om_Screen) then
    begin
        if (OrgString <> '') then
            FPrintStringList.Add(OrgString);
        Result := true;
    end;

end;

function TInvPrinter.Reset: boolean;
begin
    Result := false;
    if (Connected) then
    begin
        if (FPrinter = NIXDORF_77) then
        begin
            WriteString(FStart + FReset + FEnd, '');
            WriteString(FStart + #27 + 't' + #128 + FEnd, '');
            WriteString(FStart + #27 + 'z' + #1 + FEnd, '');
        end
        else
        if (FPrinter = EPSON_RPU420) then
        begin
            WriteString(FStart + FReset + FEnd, '');
            WriteString(FStart + #27 + 'c4' + #0 + FEnd, '');
        end
        else
        if (FPrinter = NIXDORF_20) then
        begin
            WriteString(FStart + FReset + FEnd, '');
            result := WriteString(FStart + #27 + 't' + #144 + FEnd, '');
        end
        else
        if (FPrinter in [EPSON_TM200, EPSON_TMT882, WP_T810, RP_600, RP_600_V9, PP8000, EPSON_TM88IV, EPSON_TMT70]) then
        begin
            WriteString(FStart + FReset + FEnd, '');
            WriteString(FStart + #27 + 'c4' + #0 + FEnd, '');
        end
        else
            result := WriteString(FStart + FReset + FEnd, '');
    end
    else
        result := false;
end;

//TDisplay

constructor TDisplay.Create(AOwner: TComponent);
begin
    inherited Create(AOwner);
    FHandle := 0;
    FRowLength := 20;
    FWelcomeUpperRow :=  space(FRowLength);
    FWelcomeLowRow   :=  space(FRowLength);
    SetDisplayerDefault(FDisplayer);
end;

destructor TDisplay.Destroy;
begin
    inherited Destroy;
end;

procedure TDisplay.SetDisplayerDefault(Value: TDisplayerType);
begin
    //內定資料
    FDisplayer := Value;

    BaudRate := brt_9600;
    DataBits := dbt_8;
    StopBits := sbt_OneStopBit;
    ParityCheck := pct_NoParity;
    FUpperScroll := false;

    case FDisplayer of
        NIXDORF_BAL3G, NIXDORF_BAL3:
                ParityCheck := pct_OddParity;
        IBM_SURE_ONE:
                StopBits := sbt_TwoStopBits;
        CD_8240, CD_9256:
                FRowLength := 15;
        PD_7200:
                Delay := 100;
    end;
end;

procedure TDisplay.SetDisplayerCommand;
begin
    //設定客顯基本顯示命令
    case FDisplayer of
        VC_108:
            begin
                FStart := #27; // 0x1b
                FEnd := #13;
                FUpperCommand := #81 + #65;   // 0x51+0x41
                FLowCommand   := #115;   //0x73
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := '';
            end;

        NIXDORF_BAL3G,
        NIXDORF_BAL3:
            begin
                FStart := #27;
                FEnd := #13;
                FUpperCommand := '[1;1H';
                FLowCommand   := '[2;1H';
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := #19;
            end;

        CD_5220:
            begin
                FStart := #27 + #27;
                FEnd := #13;
                FUpperCommand := 'QA';
                FLowCommand   := 'QB';
                FTopCommand   := FUpperCommand;
                FScrollCommand := 'QD';
                FHorzontalScrollMode := #19;
            end;

        CC_110:
            begin
                FStart := #27;
                FEnd := #13;
                FUpperCommand := 'QA';
                FLowCommand   := 'QB';
                FTopCommand   := FUpperCommand;
                FScrollCommand := 'QD';
                FHorzontalScrollMode := #19;
            end;

        WD_304:
            begin
                FStart := #27;
                FEnd := #13;
                FUpperCommand := 'qE';
                FLowCommand   := 'qF';
                FTopCommand   := FUpperCommand;
                FScrollCommand := 'uD';
                FHorzontalScrollMode := #19;
            end;

        RS_304:
            begin
                FStart := #27;
                FEnd := #13;
                FUpperCommand := 'uB';
                FLowCommand   := 'uC';
                FTopCommand   := FUpperCommand;
                FScrollCommand := 'uD';
                FHorzontalScrollMode := #19;
            end;

        PD_2100:
            begin
                FStart := #27 + #72;
                FEnd := #13;
                FUpperCommand := #0;
                FLowCommand   := #20;
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := #18 + #10;
            end;

        IBM_SURE_ONE:
            begin
                FStart := #16;
                FEnd := #13;
                FUpperCommand := #0;
                FLowCommand   := #20;
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := #18+ #10;
            end;

        DM_D110:
            begin
                FStart := '';
                FEnd := #13;
                FUpperCommand := #31 + #36 + #1 + #1;
                FLowCommand   := #31 + #36 + #1 + #2;
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := '';
            end;

        DM_D500:
            begin
                FStart := '';
                FEnd := #13;
                FUpperCommand := #31 + #36 + #1 + #2;
                FLowCommand   := #31 + #36 + #1 + #3;
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := #31 + #3;
            end;

        CD_8240:
            begin
                FStart := #27;
                FEnd := #13;
                FTopCommand   := char($6c) + #1 + #1 + FEnd + FStart + 'Q';
                FUpperCommand := char($6c) + #1 + #2 + FEnd + FStart + 'Q';
                FLowCommand   := char($6c) + #1 + #3 + FEnd + FStart + 'Q';
                FScrollCommand := 'Q';
                FHorzontalScrollMode := #19;
            end;

        CD_9256:
            begin
                FStart := #27;
                FEnd := #13;
                FTopCommand   := char($6c) + #1 + #1 + FEnd;
                FUpperCommand := char($6c) + #1 + #2 + FEnd;
                FLowCommand   := char($6c) + #1 + #3 + FEnd;
                FScrollCommand := 'Q';
                FHorzontalScrollMode := #19;
            end;

        PD_7200:
            begin
                FStart := char($1f) + char($43) + #1;
                FEnd := #13;
                FUpperCommand := char($1f) + char($24)+ char($1) + char($1) + FEnd;
                FLowCommand   := char($1f) + char($24)+ char($1) + char($2) + FEnd;
                FTopCommand   := FUpperCommand;
                FScrollCommand := '';
                FHorzontalScrollMode := char($1f) + char($03) + FEnd;
            end;
        //20090210 新增富士通客顯VFD
        FUJITSU_KD_290X:
            begin
                FStart := '';
                FEnd := #$0d;
                FTopCommand   := #$1b + #$5b + #$31 + #$3b + #$31 + #$48;
                FUpperCommand := #$1b + #$5b + #$32 + #$3b + #$31 + #$48;
                FLowCommand   := #$1b + #$5b + #$33 + #$3b + #$31 + #$48;
                FScrollCommand := '';
                FHorzontalScrollMode := '';
            end;
    end;
end;

function TDisplay.WriteCom(SendString: AnsiString): boolean;
var
    numWrite: DWORD;
    sndStr: PAnsiChar;
begin
    //送入通訊設備
    sndStr := PAnsiChar(SendString);
    Sleep(Delay);
    Result := WriteFile(FHandle, sndStr^, Length(SendString), numWrite, Nil);
end;

procedure TDisplay.ApplyComSetup;
begin
    //套用設定
    SetupComm(FHandle, InbondBuffer, OutbondBuffer);
    GetCommState(FHandle, FDCB);
    FDCB.BaudRate := BaudRateValue[BaudRate];
    FDCB.ByteSize := DataBitsValue[DataBits];
    FDCB.Parity   := ParityCheckValue[ParityCheck];
    FDCB.StopBits := StopBitsValue[StopBits];
    FDCB.XonChar  := char(17);
    FDCB.XoffChar := char(19);
  	SetCommState(FHandle, FDCB);
    SetDisplayerCommand();

  FTimeOuts.ReadIntervalTimeout         := MAXDWORD;
  FTimeOuts.ReadTotalTimeoutMultiplier  := 3;
  FTimeOuts.ReadTotalTimeoutConstant    := 100;
  FTimeOuts.WriteTotalTimeoutMultiplier := 3;
  FTimeOuts.WriteTotalTimeoutConstant   := WTimeOut;
	SetCommTimeouts(FHandle, FTimeOuts);
end;

procedure TDisplay.SetUpperScroll(Value: boolean);
begin
    FUpperScroll := Value;
    DoUpperScroll(FUpperScroll);
end;

function TDisplay.DoUpperScroll(Value: boolean): boolean;
begin
    Result := false;
    if (Connected) then
        if (FUpperScroll) then
        begin
            WriteCom(FStart + FHorzontalScrollMode);
            Result := WriteCom(FStart + FScrollCommand + FUpperLine + FEnd);
        end
        else
            Result := WriteCom(FStart + FUpperCommand + FUpperLine + FEnd);
end;

procedure TDisplay.SetRowLength(Value: integer);
begin
    //設定輸出顯示字數
    FRowLength := Value;
    if (Connected) then
        ApplyComSetup;
end;

procedure TDisplay.Initial;
var
    Cmd: AnsiString;
begin

    if (Connected() = false)  then exit;

    case FDisplayer of
        VC_108:
            begin
                Cmd := #27 + #81 + #63 + #12 + #13;
                WriteCom(Cmd);
                Cmd := #31 + #115 + #1;
                WriteCom(Cmd);
            end;

        NIXDORF_BAL3G,
        NIXDORF_BAL3:
            begin
                Cmd := #27 + #82 + #146;
                WriteCom(Cmd);
                Cmd := #27 + #91 + #49 + #66;
                WriteCom(Cmd);
                Cmd := #27 + #91 + #49 + #73;
                WriteCom(Cmd);
            end;

        CD_5220,
        CD_9256,
        CD_8240:
            begin
                Cmd := #27 + #64;
                WriteCom(Cmd);
                Cmd := #27 + #17;
                WriteCom(Cmd);
                Cmd := #27 + #95 + #0;
                WriteCom(Cmd);
            end;

        PD_2100:
            begin
                Cmd := #27 + #73;
                WriteCom(Cmd);
                Cmd := #22;
                WriteCom(Cmd);
            end;

        IBM_SURE_ONE:
            begin
                Cmd := #31 + #17;
                WriteCom(Cmd);
                WriteCom(Cmd);
            end;

        DM_D110,
        DM_D500:
            begin
                Cmd := #27 + #64;
                WriteCom(Cmd);
                Cmd := #31 + #1;
                WriteCom(Cmd);
                Cmd := #31 + #67 + #0;
                WriteCom(Cmd);
            end;
        PD_7200:
            begin
                Cmd := char($1b) + char($40) + #13;
                WriteCom(Cmd);
                Cmd := char($1f) + char($1) + #13;
                WriteCom(Cmd);
            end;
        //20090210 新增富士通客顯VFD
        FUJITSU_KD_290X:
            begin
                //設定 Power ON
                Cmd := #$1b + #$5c + #$3f + #$53 + #$45;
                WriteCom(Cmd);

                //選擇 Character Code Setting
                //【Code】 1BH,5CH,43H,43H,Pn
                //Pn 文字代碼 備註
                //30H Shift JIS 啟動時預設
                //31H JIS
                //32H KSC5601
                //33H GB2312
                //34H Big5
                Cmd := #$1b + #$5C + #$43 + #$43 + #$34;
                WriteCom(Cmd);

                //清除客顯
                Cmd := #$1b + #$5b + #$32 + #$4a;
                WriteCom(Cmd);
            end;

    end;
end;

function TDisplay.Connect: boolean;
var
  mPort: PChar;
begin
    //連接指定通訊埠
    if not(csDesigning in ComponentState) then
    begin
        if (Connected) then
            DisConnect;

        if UniName = '' then
          mPort := DevicePortStrValue[Port]
        else
          mPort := PChar(UniName);

        FHandle := CreateFile(mPort, GENERIC_READ or GENERIC_WRITE, 0, 0,
                             OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

        if (FHandle = INVALID_HANDLE_VALUE) then
        begin
            Result := false;
            exit;
        end;

        ApplyComSetup;
        Initial;
        ClearAll;
        Result := DisplayString2(FWelcomeUpperRow, FWelcomeLowRow);
    end;
end;

function TDisplay.Connected: boolean;
begin
    //查詢是否連線
    Result := FHandle > 0;
end;

function TDisplay.DisConnect: boolean;
begin
    //中斷連接
    if (Connected) then
    begin
        CloseHandle(FHandle);
        FHandle := 0;
        Result := true;
    end
    else
        Result := false;
end;

function TDisplay.DisplayString2(UpperRow, LowRow: AnsiString): boolean;
begin

    //顯示字串
    Result := false;
    FUpperLine := UpperRow;
    FLowerLine := LowRow;
    if (Connected) then
        if (FUpperScroll) then
        begin
            if ((FDisplayer = DM_D110) or (FDisplayer = DM_D500)) then
            begin
                Result := true;
                exit;
            end;
            WriteCom(FStart + FLowCommand + Copy(LowRow, 1, FRowLength) + FEnd);
            Result := DoUpperScroll(FUpperScroll);
        end
        else
            if ((UpperRow = '') and (LowRow = '')) then
                ClearAll
            else
            begin
                if ((FDisplayer = DM_D110) or (FDisplayer = DM_D500)) then
                begin
                    WriteCom(FUpperCommand);
                    WriteCom(FStart + Copy(UpperRow, 1, FRowLength) + FEnd);
                    WriteCom(FLowCommand);
                    Result := WriteCom(FStart + Copy(LowRow, 1, FRowLength) + FEnd);
                end
                else
                begin
//                    if FDisplayer = PD_7200 then
//                      WriteCom(char($14) + char($1c) + #13#10);
                      
                    Result := WriteCom(FStart + FUpperCommand + Copy(UpperRow, 1, FRowLength) + FEnd);
                    if (Result) then
                        if (FDisplayer = VC_108) then
                            Result := WriteCom(FStart + FLowCommand + Copy(LowRow, 1, FRowLength))
                        else
                            Result := WriteCom(FStart + FLowCommand + Copy(LowRow, 1, FRowLength) + FEnd);
                end;
                exit;
            end
    else
        Result := false;
end;

function TDisplay.DisplayString3(TopRow, UpperRow, LowRow: AnsiString): boolean;
begin
    //顯示字串
    ShowLine0(TopRow);
    Result := DisplayString2(UpperRow, LowRow);
end;

function TDisplay.ShowLine(Line: integer; UpperRow: AnsiString): boolean;
var
    Cmd: AnsiString;
begin
    //顯示上行資料
    if (FDisplayer <> DM_D500) then
    begin
        Result := true;
        exit;
    end;

    if (Connected) then
    begin
        Cmd := #31 + #36 + #1 + char(Line);
        WriteCom(Cmd);
        Result := WriteCom(FStart + Copy(UpperRow, 1, FRowLength) + FEnd);
    end
    else
        Result := false;
end;

function TDisplay.ShowLine0(TopRow: AnsiString): boolean;
begin
    //顯示顯示最上行(第0行)資料
    FTopLine := TopRow;
    if (Connected) then
        if (FUpperScroll) then
            Result := DoUpperScroll(FUpperScroll)
        else
            Result := WriteCom(FStart + FTopCommand + Copy(TopRow, 1, FRowLength) + FEnd)
    else
        Result := false;
end;

function TDisplay.ShowLine1(UpperRow: AnsiString): boolean;
begin
    //顯示上行資料(Middle Row)
    FUpperLine := UpperRow;
    if (Connected) then
        if (FUpperScroll) then
            Result := DoUpperScroll(FUpperScroll)
        else
            Result := WriteCom(FStart + FUpperCommand + Copy(UpperRow, 1, FRowLength) + FEnd)
    else
        Result := false;
end;

function TDisplay.ShowLine2(LowRow: AnsiString): boolean;
begin
    //顯示下行資料(Bottom Row)
    FLowerLine := LowRow;
    if (Connected) then
        Result := WriteCom(FStart + FLowCommand + Copy(LowRow, 1, FRowLength) + FEnd)
    else
        Result := false;
end;

function TDisplay.ScrollString(UpperRow: AnsiString): boolean;
begin
    if (Connected) then
    begin
        WriteCom(FStart + FHorzontalScrollMode);
        Result := WriteCom(FStart + FScrollCommand + Copy(UpperRow, 1, FRowLength) + FEnd);
    end
    else
        Result := false;
end;

function TDisplay.NoCommand(Value: AnsiString): boolean;
begin
    //無控制指令
    if (Connected) then
        Result := WriteCom(Value)
    else
        Result := false;
end;

function TDisplay.ClearAll: boolean;
begin
    //清除顯示資料
    if (FDisplayer in [CD_8240, CD_9256]) then
        ShowLine0(space(FRowLength));

    Result := DisplayString2(space(FRowLength), space(FRowLength));
end;

procedure TDisplay.SetUniName(Value: AnsiString);
begin
    inherited;
    if (Connected) then
        ApplyComSetup;
end;

//THandyScan

constructor THandyScan.Create(AOwner: TComponent);
begin
    inherited Create(AOwner);
    FComRcv := TPOSComm.Create(nil);
    FComRcv.OnReceiveData := ReceiveData;
    SetHandyScannerDefault(FHandyScanner);
end;

destructor THandyScan.Destroy;
begin
    Disconnect;
    FreeAndNil(FComRcv);
    inherited Destroy;
end;

procedure THandyScan.SetHandyScannerDefault(Value: TBarCodeScannerType);
begin
    //內定資料
    FHandyScanner := Value;
    case (FHandyScanner) of
        HS_232:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
                FEndLength := 2;
            end;
        FJ_M411:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_OddParity;
                FEndLength := 2;
            end;
    end;
end;

procedure THandyScan.SetReadTimeOut(Value: integer);
begin
    //設定逾時
    FReadTimeOut := Value;
    if (Connected) then
        ApplyComSetup;
end;

procedure THandyScan.SetInterval(Value: integer);
begin
    FInterval := Value;
end;

procedure THandyScan.ApplyComSetup;
var
  mPort: AnsiString;
begin
    //套用設定
    if UniName = '' then
      mPort := DevicePortStrValue[Port]
    else
      mPort := UniName;

    with FComRcv do
    begin
      CommName  := mPort;
      BaudRate  := BaudRateValue[Self.BaudRate];
      Parity    := TParity(iif(Self.ParityCheck = pct_NoParity, 0,
                           iif(Self.ParityCheck = pct_OddParity, 1, 2)));
      ByteSize  := TByteSize(iif(Self.DataBits = dbt_5, 0,
                             iif(Self.DataBits = dbt_6, 1,
                             iif(Self.DataBits = dbt_7, 2, 3))));//DataBitsValue[DataBits];
      StopBits  := TStopBits(iif(Self.StopBits = sbt_OneStopBit, 0,
                             iif(Self.StopBits = sbt_One5StopBits, 1, 2)));  //StopBitsValue[StopBits];
      //DsrSensitivity := CheckDSR;
    end;
end;

function THandyScan.CheckDSRStatus: boolean;
var
  mErrCount: integer;
  State: DWORD;
begin
  mErrCount := 0;
  while (mErrCount < 3) do
  begin
      State := FComRcv.GetModemState;
      Result := ((State and MS_DSR_ON) <> 0);

      if Result then exit;

      mErrCount := mErrCount + 1;
  end;
end;

procedure THandyScan.ReceiveData(Sender: TObject; Buffer: AnsiString{Pointer}; BufferLength: Word);
var
  end_da, end_ad, end_a, end_d, mTmpData, Value: AnsiString;
begin
    end_da  := #13#10;
    end_ad  := #10#13;
    end_a   := #10;
    end_d   := #13;
    mTmpData := '';
    try
        if (Connected) then
        begin
            Value := AnsiString(Buffer);
            if (Value <> '') then
            begin
                while (Length(Value) <> 0) do
                begin
                    if (Pos(end_da, Value) > 0) then
                    begin
                        mTmpData := Copy(Value, 1, Pos(end_da, Value) - 1);
                        if Assigned(FOnGetData) then
                            FOnGetData(self, mTmpData, Length(mTmpData));
                        Value := Copy(Value, Pos(end_da, Value) + Length(end_da), Length(Value));
                    end
                    else
                    if (Pos(end_ad, Value) > 0) then
                    begin
                        mTmpData := Copy(Value, 1, Pos(end_ad, Value) - 1);
                        if Assigned(FOnGetData) then
                            FOnGetData(self, mTmpData, Length(mTmpData));
                        Value := Copy(Value, Pos(end_ad, Value) + Length(end_ad), Length(Value));
                    end
                    else
                    if (Pos(end_a, Value) > 0) then
                    begin
                        mTmpData := Copy(Value, 1, Pos(end_a, Value) - 1);
                        if Assigned(FOnGetData) then
                            FOnGetData(self, mTmpData, Length(mTmpData));
                        Value := Copy(Value, Pos(end_a, Value) + Length(end_a), Length(Value));
                    end
                    else
                    if (Pos(end_d, Value) > 0) then
                    begin
                        mTmpData := Copy(Value, 1, Pos(end_d, Value) - 1);
                        if Assigned(FOnGetData) then
                            FOnGetData(self, mTmpData, Length(mTmpData));
                        Value := Copy(Value, Pos(end_d, Value) + Length(end_d), Length(Value));
                    end
                    else
                    begin
                        if Assigned(FOnGetData) then
                            FOnGetData(self, Value, Length(Value));
                        exit;
                    end;
                end;
            end;
        end;
    except
    end;
end;

function THandyScan.Connect: boolean;
begin
  try
    ApplyComSetup;
    FComRcv.StartComm;
    Result := Connected;
  except
    Result := false;
  end;
end;

function THandyScan.Connected:boolean;
begin
  //查詢是否連線
  Result := FComRcv.Handle <> 0;
end;

function THandyScan.DisConnect:boolean;
begin
    if (Connected) then
    begin
      FComRcv.StopComm;
      Result := true;
    end
    else
      Result := false;
end;

end.

