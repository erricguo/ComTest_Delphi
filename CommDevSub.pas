//20100715 add by 3188 �W�[����@��(Ingenico 144�榡ICT220_CATHAY)
//20101028 add by 3188 �W�[�ɤs(Vx510 144)
//20130107 add by 3188 �W�[����H��NETPOS3000 200�榡
//20130117 add by 3188 �W�[����AS320 144�榡
//20170525 add by 07113 �s�W�x�s400 TaiHsin_400
//20171117 add by 07113 �s�W�@��400 FirstBank_400
//20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002 �s�W�ɤs600 ESUN_S80RF_600
//20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002 �s�W���H250 CTBC_AS320_250
unit CommDevSub;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Dialogs, extctrls,
  Stdctrls, CommTypes, ShellApi, Forms, ATFileNotification;

type
  //TEDCCom

  TOnEDCErrorEvent = procedure(Sender: TObject; ErrorCode: integer;
            var Suspend: boolean) of object;
  TOnEDCMsgEvent = procedure(Sender: TObject; MsgCode, Second: integer;
            var Suspend: boolean) of object;
  TOnGetEDCEvent = procedure(Sender: TObject; CardValue: TCreditCard) of object;

  TOnEDCResetTimeEvent = procedure(Sender: TObject) of object;  //20171122 edc.exe

  TEDCCom = class(TBaseCom)
  private
    FHandle: THandle;
    FEdcType: TEdcType;
    FDCB: DCB;
    FTimeOuts: COMMTIMEOUTS;
    FEDCDataLength, FEDCTimeOut: integer;
    FEndLength, FReadTimeOut: integer;
    FSuspend, FDebugMode: boolean;
    FOnEDCError: TOnEDCErrorEvent;
    FOnEDCMsg: TOnEDCMsgEvent;
    FOnGetEDC: TOnGetEDCEvent;
    Notification: TATFileNotification;                        //20171122 edc.exe
    FState:Integer;                                           //20171122 edc.exe
    FTotSec: integer;                                         //20171122 edc.exe
    FOnEDCResetTime: TOnEDCResetTimeEvent;                    //20171122 edc.exe
    procedure RUN_DOS_COMMAND(xExeFileName, xExeDirectory, xParameters: String; xShow: Integer= SW_SHOW);
    procedure SetEdcTypeDefault(Value: TEdcType);
    procedure SetReadTimeOut(Value: integer);
    function ReadCom: string;
    procedure ApplyComSetup;
    function Connect: boolean;
    function Connected: boolean;
    function DisConnect: boolean;
    function getc: AnsiChar;
    function puts(str: AnsiString; len: integer): integer;
    function Cal_Lrc(xValue: PAnsiChar; xLen: integer): Ansichar;
    function TxTransmit(const xBuffer: PAnsiChar; var xRetValue: PAnsiChar): boolean;
    procedure DebugLog(Msg: AnsiString);
    procedure NotifyFile(AFileName : String);                 //20171122 edc.exe
    procedure NotificationChanged(Sender: TObject);
    procedure SetOnEDCResetTime(const Value: TOnEDCResetTimeEvent);
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function SendToEDC(TransType, Payment, StoreNo, PosNo, TrnNo: AnsiString): boolean;
    function SendToEDC2(TransType, Payment, StoreNo, PosNo, TrnNo: string; CardType:Integer): boolean;  //20130115 ADD BY 4084
    function SendToEDCV3(TransType, Payment, HostID: AnsiString;
        FSLength,   //����
        PaymentS,   //�I�ڰ_�l��m
        PaymentE,   //�I�ںI���m
        ResponseS,  //�q�T���Ұ_�l��m
        ResponseE,  //�q�T���ҺI���m
        AppS,       //���v�X�_�l��m
        AppE,       //���v�X�I���m
        CardS,      //�d���_�l��m
        CardE       //�d���I���m
        :integer ): boolean;
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
    property DebugMode: boolean read FDebugMode write FDebugMode;
    property EDCType: TEdcType read FEdcType write SetEdcTypeDefault;
    property EDCTimeOut: integer read FEDCTimeOut write FEDCTimeOut;
    property RTimeOut: integer read FReadTimeOut write SetReadTimeOut;
    property OnEDCError: TOnEDCErrorEvent read FOnEDCError write FOnEDCError;
    property OnEDCMsg: TOnEDCMsgEvent read FOnEDCMsg write FOnEDCMsg;
    property OnGetEDC: TOnGetEDCEvent read FOnGetEDC write FOnGetEDC;
    property OnEDCResetTime: TOnEDCResetTimeEvent  read FOnEDCResetTime write SetOnEDCResetTime;
  end;

  //TKeyHook
  PKBDLLHOOKSTRUCT = ^KBDLLHOOKSTRUCT;
  KBDLLHOOKSTRUCT = record
            VKCode: DWORD;
            ScanCode: DWORD;
            Flags: DWORD;
            Time: DWORD;
            dwExtraInfo: DWORD;
  end;
  TOnHookKeyUpEvent = procedure(Sender: TObject; ScanCode: dword;
      Key: Ansichar; KeyState: TShiftState) of object;
  TKeyHook = class(TComponent)
  private
    FEnabled: boolean;
    FHookWin, FHook: HHOOK;
    FHookKeyUp: TOnHookKeyUpEvent;
    FAltKey, FShiftKey, FCtrlKey: boolean;
    FPreSysKeyUp, FSendkey: boolean;
    FStute:Boolean;
    FShiftState:Boolean;
  protected
    procedure SetEnabled(Value: boolean);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    procedure ResetKeyoardState;
    property HookWin: HHOOK read FHookWin write FHookWin;
    property Enabled: boolean read FEnabled write SetEnabled default false;
    property PreSysKeyUp: boolean read FPreSysKeyUp write FPreSysKeyUp;
    property Sendkey: boolean read FSendkey write FSendkey;
    property OnHookKeyUpEvent: TOnHookKeyUpEvent read FHookKeyUp write FHookKeyUp;
  end;

implementation

const
    //MAX_DATA    = 500; //20180621 mark by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002
    MAX_DATA    = 800; //20180621 modi by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002
    NULL        = #$0;
    STX         = #$02;
    ETX         = #$03;
    ACK         = #$06;
    NAK         = #$15;
    EDCPath  = 'ECR_Exe';
    ANotiEDCPath = 'ECR_Exe\out.dat';
//    FileName = 'ecr_in.dat';
//    ReadFileName = 'ecr_out.dat';
    FileName = 'in.dat';
    ReadFileName = 'out.dat';

var
    KeyHook: TKeyHook;

//TEDCCom

constructor TEDCCom.Create(AOwner: TComponent);
begin
    inherited Create(AOwner);
    FEDCTimeOut := 60;
    FHandle := 0;
    FReadTimeOut := 300;
    SetEdcTypeDefault(FEdcType);
end;

destructor TEDCCom.Destroy;
begin
    inherited Destroy;
end;

function TEDCCom.Connected: boolean;
begin
    //�d�߬O�_�s�u
    Result := FHandle > 0;
end;

procedure TEDCCom.ApplyComSetup;
begin
  //�M�γ]�w
  SetupComm(FHandle, InbondBuffer, OutbondBuffer);
  GetCommState(FHandle, FDCB);
  FDCB.BaudRate := BaudRateValue[BaudRate];
  FDCB.ByteSize := DataBitsValue[DataBits];
  FDCB.Parity   := ParityCheckValue[ParityCheck];
  FDCB.StopBits := StopBitsValue[StopBits];
  FDCB.XonChar  := char(17);
  FDCB.XoffChar := char(19);
  SetCommState(FHandle, FDCB);

  FTimeOuts.ReadIntervalTimeout         := MAXDWORD;
  FTimeOuts.ReadTotalTimeoutMultiplier  := 3;
  FTimeOuts.ReadTotalTimeoutConstant    := FReadTimeOut;
  FTimeOuts.WriteTotalTimeoutMultiplier := 3;
  FTimeOuts.WriteTotalTimeoutConstant   := WTimeOut;
  SetCommTimeouts(FHandle, FTimeOuts);
end;

procedure TEDCCom.SetEdcTypeDefault(Value: TEdcType);
begin
    //���w���
    FEdcType := Value;
    //�]�w�ǿ��w
    case (FEdcType) of
        H5000, S9000, ICT220_CATHAY, VX510_ESun, NETPOS3000, AS320_200, AllPay, AS320_UnionPay, SinoPac:
            begin
                BaudRate := brt_9600;
                DataBits := dbt_7;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_EvenParity;
                FEndLength := 1;
            end;
        VX570_400, AS320, FirstBank_400:  //20171117 add by 07113
            begin
                BaudRate := brt_9600;
                DataBits := dbt_8;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_NoParity;
                FEndLength := 1;
            end;
        SAGEM, ESUN_S80RF_600, CTBC_AS320_250: //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002 add ESUN_S80RF_600,CTBC_AS320_250
            begin
                BaudRate := brt_19200;
                DataBits := dbt_7;
                StopBits := sbt_OneStopBit;
                ParityCheck := pct_EvenParity;
                FEndLength := 1;
            end;
        //20171117 add by 07113
        TaiHsin_400:
            begin
              BaudRate := brt_115200;
              DataBits := dbt_8;
              StopBits := sbt_OneStopBit;
              ParityCheck := pct_NoParity;
              FEndLength := 1;
            end;
         //20171117 add by 07113
    end;

    //�]�w�ǿ��ƪ���
    case (FEdcType) of
      H5000, S9000, SAGEM, ICT220_CATHAY, VX510_ESun, AS320, AllPay, AS320_UnionPay, SinoPac:
        FEDCDataLength  := 144;
      VX570_400, TaiHsin_400, FirstBank_400:  //20171117 add by 07113 add TaiHsin_400,FirstBank_400
        FEDCDataLength  := 400;
      NETPOS3000, AS320_200:
        FEDCDataLength  := 200;
      //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
      ESUN_S80RF_600:
        FEDCDataLength  := 600;
      CTBC_AS320_250:
        FEDCDataLength  := 250;
      //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
    end;
end;

procedure TEDCCom.SetOnEDCResetTime(const Value: TOnEDCResetTimeEvent);
begin
  FOnEDCResetTime := Value;
  FTotSec :=0;
end;

//procedure TEDCCom.SetOnEDCResetTime(const Value: TOnEDCResetTimeEvent);
//begin
//  FOnEDCResetTime := Value;
//  FTotSec :=0;
//end;

function TEDCCom.Connect: boolean;
begin
    //�s�����w�q�T��
    if not(csDesigning in ComponentState) then
    begin
        if (Connected) then
            DisConnect;

    	FHandle := CreateFile(DevicePortStrValue[Port], GENERIC_READ or GENERIC_WRITE, 0, 0,
                               OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

    	if (FHandle = INVALID_HANDLE_VALUE) then
        begin
            Result := false;
            exit;
        end;

        ApplyComSetup;
        Result := true;
    end;
end;

function TEDCCom.DisConnect:boolean;
begin
    if (Connected) then
    begin
        FlushFileBuffers(FHandle);
        CloseHandle(FHandle);
        FHandle := 0;
        Result := true;
    end
    else
        Result := false;
end;

function TEDCCom.ReadCom: string;
var
    ReadChar: char;
    success: boolean;
    numRead: DWORD;
begin
    Result := '';
    success := ReadFile(FHandle, ReadChar, 1, numRead, Nil);

    if (numRead <> 0) then
    begin
        Result := Result + ReadChar;
        while (numRead <> 0) do
        begin
            ReadFile(FHandle, ReadChar, 1, numRead, Nil);
            if (numRead <> 0) then
                Result := Result + char(ReadChar);
        end;
    end;
end;

procedure TEDCCom.RUN_DOS_COMMAND(xExeFileName, xExeDirectory,
  xParameters: String; xShow: Integer);
// xShow �ѼƽаѦ�  ShowWindow �� nCmdShow
var
  Result: Boolean;
  ShellExInfo: TShellExecuteInfo;
begin
  FillChar(ShellExInfo, SizeOf(ShellExInfo), 0);
  with ShellExInfo do
  begin
    cbSize:= SizeOf(ShellExInfo);
    fMask:= see_Mask_NoCloseProcess;
    Wnd:= Application.Handle;
    lpFile:= PChar(xExeFileName);
    lpDirectory:= PChar(xExeDirectory);
    lpParameters:= PChar(xParameters);
    nShow:= xShow;
  end;

  Result:= ShellExecuteEx(@ShellExInfo);
  if Result then
  begin
//    while WaitForSingleObject(ShellExInfo.HProcess, 10)= WAIT_TIMEOUT do
//    begin
//      Application.ProcessMessages;
//      if Application.Terminated then
//        Break;
//    end;
  end;
end;

procedure TEDCCom.SetReadTimeOut(Value: integer);
begin
    FReadTimeOut := Value;
    if (Connected) then
        ApplyComSetup;
end;

function TEDCCom.getc: AnsiChar;
var
    ReadChar: Ansichar;
    success: boolean;
    numRead: DWORD;
begin
    Result := Ansichar(0);
    success := ReadFile(FHandle, ReadChar, 1, numRead, Nil);
    if (numRead <> 0) then
        Result := ReadChar;
end;

procedure TEDCCom.NotificationChanged(Sender: TObject);
var
  SL : TStringList;
  mCrd: TCreditCard;
  mRetStr:String;
  mBol:Boolean;
  ApRunPath:String;
begin
  ApRunPath:=ExtractFilePath(Application.ExeName);
  if FileExists(ApRunPath + ANotiEDCPath) then
  begin
    SL := TStringList.Create;
    try
      SL.LoadFromFile(ApRunPath + ANotiEDCPath);
      mRetStr := SL.Text;
      if FDebugMode then
        DebugLog('Rcv>' + SL.Text);                                     //�O������LOG
    finally
      SL.Free;
    end;

    mBol:= False;
    with mCrd do
    begin
      case (FEdcType) of
        VEGA_9000:
          begin
            Host_id     := '';
            Card_Name   := ''; //�O�d���� //copy(mRetStr,  8,  8);      //..
            Ref_no      := trim(copy(mRetStr, 93, 12));                 //Reference
            Card_no     := trim(copy(mRetStr, 17, 20));
            Auth_no     := trim(copy(mRetStr, 61,  9));
            Response    := trim(copy(mRetStr, 82,  4));
          end;

        TaiHsin_400:
          begin
            Host_id     := trim(copy(mRetStr,  3,  2));
            Ref_no      := trim(copy(mRetStr,  5,  6));
            Card_no     := trim(copy(mRetStr, 11, 19));
            Auth_no     := trim(copy(mRetStr, 58,  9));
            Response    := trim(copy(mRetStr, 79,  4));
            Transfinished := trim(copy(mRetStr, 260,  1));
            Timerblockreset := trim(copy(mRetStr, 261,  4));
          end;
      end;
      case (FEdcType) of
        VEGA_9000:
          begin
            if ((Response <> '0000') and (Response <> '00TD')) then
            begin
              Host_id     := '';
              Card_Name   := '';
              Card_no     := '';
              Auth_no     := '';
              Ref_no      := '';
              mBol := True;
            end else
            begin
              Response := '0000';
              mBol := True;
            end;
          end;

        TaiHsin_400:
          begin
            if ((Response <> '0000') and (Response <> '00TD'))   then
            begin
              Host_id     := '';
              Card_Name   := '';
              Card_no     := '';
              Auth_no     := '';
              Ref_no      := '';
              mBol := True;
            end else
            if ((Response = '0000') and (Transfinished = 'N') and (Timerblockreset = '0250')) then
            begin
              mBol := False;
              DeleteFile(EDCPath +'\'+ ReadFileName);
              if Assigned(FOnEDCResetTime) then
                FOnEDCResetTime(self);

              FState := 1;
            end else
            begin
              Response := '0000';
              mBol := True;
            end;
          end;
      end;
    end;


    if mBol = True then
    begin
      FState := 2;
      if Assigned(FOnGetEDC) then
        FOnGetEDC(self, mCrd);
    end;
  end;
end;

procedure TEDCCom.NotifyFile(AFileName: String);
begin
  with Notification do
  begin
    Stop;
    Options:= [];
    FileName := AFileName ;
    { TODO : �ڥu�����s��� }
    Options:= Options + [foNotifyCreation,foNotifyLastWrite];
    Start;
  end;
end;

function  TEDCCom.puts(str: AnsiString; len: integer): integer;
var
    numWrite: DWORD;
    mSndData: PChar;
begin
    mSndData := PChar(str);
    if WriteFile(FHandle, mSndData^, len, numWrite, Nil) then
        Result := numWrite
    else
        Result := 0;
end;

function TEDCCom.SendToEDC(TransType, Payment, StoreNo, PosNo, TrnNo: AnsiString): boolean;
var
    mRetBuf, mSndBuf, mTmpBuff, mZeroBuff: array[0..MAX_DATA] of AnsiChar; //20180621 modi by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  add mZeroBuff
    mPRetBuf: PAnsiChar;
    mCrd: TCreditCard;
    mPayNum, mTenantNo, mRetStr: AnsiString;
begin
    Result := false;
    if (not Connect) or (not Connected) then
    begin
        DisConnect;
        if assigned(FOnEDCError) then
            FOnEDCError(self, 1, FSuspend);//'��d���s�u����!';
        exit;
    end;
    FillChar(mTmpBuff, FEDCDataLength, AnsiChar(' '));
    FillChar(mZeroBuff, FEDCDataLength, AnsiChar('0')); //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002
    case (FEdcType) of
      H5000, S9000, SAGEM, ICT220_CATHAY, VX510_ESun, SinoPac: //�ϥ�144bytes�榡
        begin
          //Trans_Type ����O(1, 2)
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //Host_ID �Ȧ�O(3, 2)
          case FEdcType of
            ICT220_CATHAY:
                StrCopy(mSndBuf +   2, '03');           //����@��ICT220�T�w�ǰe03 20100716

            SinoPac:
                StrCopy(mSndBuf +   2, '02');           //���׻Ȧ泣�e02
          else
            StrCopy(mSndBuf +   2, '00');
          end;
          //Receipt_No EDCñ��Ǹ�(5, 6)
          StrCopy(mSndBuf +   4, '000000');
          //Card_No �H�Υd�d��(���a�k�ɪť�)(11, 19)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //Card_Expire_Date �H�Υd���Ĵ�(30, 4)
          StrCopy(mSndBuf +  29, '0000');
          //Trans_Amount ������B(34, 12)
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //Trans_Date ������(46, 6)
          StrCopy(mSndBuf +  45, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //Trans_ Time ����ɶ�(52, 6)
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //Approval_No ���v�X(���a�k�ɪť�)(58, 9)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //Auth_Amount �w�����v���B(67, 12)
          StrCopy(mSndBuf +  66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //ECR_Response_Code�q�T�^���X(79, 4)
          StrCopy(mSndBuf +  78, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //EDC_Terminal_ID EDC�ݥ����N��(83, 8)
          StrCopy(mSndBuf +  82, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //Reference_No �Ȧ����Ǹ�(91, 12)
          StrCopy(mSndBuf +  90, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Exp_Amount ��L���B(103, 1  2)
          StrCopy(mSndBuf + 102, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Store_Id �M�d��(115, 6)
          StrCopy(mSndBuf + 114, PAnsiChar(AnsiString(copy(trim(StoreNo), 1, 6))));
          //POS_Tx_No ���Ⱦ�����Ǹ�(121, 6)
          StrCopy(mSndBuf + 120, PAnsiChar(AnsiString(copy(trim(PosNO), 1, 6))));
          //POS_Define ���Ⱦ��ۭq���(127, 6)
          StrCopy(mSndBuf + 126, PAnsiChar(AnsiString(copy(trim(TrnNo), 1, 6))));
          //Filler �O�d(133, 12)
          StrCopy(mSndBuf + 132, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //TOTAL	144
          //010300000000000000000000000000000000000123400000000    6000000000000000000000000000000000000000000000000000000000032232200001      0000000000000
          //(Format('010300000000000000000000000000000%10.10s00000000%6.6s%-57.57s%3.3s%3.3s%-5.5s%-6.6s0000000000000',
          //        [paystr,'60',SPACE,'A01', '001', '00001', TenantNo])));
        end;

      VX570_400:  //�ϥ�400bytes�榡
        begin
          //1	ECR Indicator	(1, 1) **���n
          StrCopy(mSndBuf +   0, 'I');
          //2	ECR Version Date	(2, 6)
          StrCopy(mSndBuf +   1, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 6))));
          //3	Trans Type Indicator	(8, 1)
          StrCopy(mSndBuf +   7, 'S');
          //4	Trans Type	(9, 2) **���n
          StrCopy(mSndBuf +   8, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //5	CUP Indicator	(11, 1) **���n
          StrCopy(mSndBuf +   10, '0');
          //6	Host ID	(12, 2)
          StrCopy(mSndBuf + 11, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //7	Receipt No	(14, 6)
          StrCopy(mSndBuf + 13, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 6))));
          //8	Card No	(20, 19)
          StrCopy(mSndBuf + 19, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 19))));
          //9	Card Expire Date	(39, 4)
          StrCopy(mSndBuf + 38, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //10	Trans Amount	(43, 12) **���n
          StrCopy(mSndBuf +  42, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //11	Trans Date	(55, 6) **���n
          StrCopy(mSndBuf +  54, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //12	Trans  Time	(61, 6) **���n
          StrCopy(mSndBuf +  60, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //13	Approval No	(67, 9)
          StrCopy(mSndBuf + 66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //14	Wave Card(Contactless) Indicator	(76, 1)
          StrCopy(mSndBuf + 75, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //15	ECR Response Code	(77, 4)
          StrCopy(mSndBuf + 76, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //16	Merchant ID	(81, 15)
          StrCopy(mSndBuf + 80, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 15))));
          //17	Terminal ID	(96, 8)
          StrCopy(mSndBuf + 95, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //18	Exp Amount	(104, 12)
          StrCopy(mSndBuf + 103, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //19	Store Id	(116, 18)
          StrCopy(mSndBuf + 115, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12) + copy(trim(StoreNo), 1, 6))));
          //20	Installment/Redeem Indicator	(134, 1)
          StrCopy(mSndBuf + 133, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //21	RDM Paid Amt	(135, 12)
          StrCopy(mSndBuf + 134, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //22	RDM Point	(147, 8)
          StrCopy(mSndBuf + 146, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //23	Points of Balance	(155, 8)
          StrCopy(mSndBuf + 154, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //24	Redeem Amt	(163, 12)
          StrCopy(mSndBuf + 162, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //25	Installment Period	(175, 2)
          StrCopy(mSndBuf + 174, PAnsiChar(AnsiString(copy(mTmpBuff, 1,  2))));
          //26	Down Payment Amount	(177, 12)
          StrCopy(mSndBuf + 176, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //27	Installment Payment Amount	(189, 12)
          StrCopy(mSndBuf + 188, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //28	Formality Fee	(201, 12)
          StrCopy(mSndBuf + 200, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //29	Card Type	(213, 2)
          StrCopy(mSndBuf + 212, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //30	Batch No	(215, 6)
          StrCopy(mSndBuf + 214, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 6))));
          //31	Start Trans Type	(221, 2)
          StrCopy(mSndBuf + 220, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //32	MP Flag	(223, 1)
          StrCopy(mSndBuf + 222, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //33	MP Response Code	(224, 6)
          StrCopy(mSndBuf + 223, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 6))));
          //34	Legal Notice1	(230, 24)
          StrCopy(mSndBuf + 229, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 24))));
          //35	Legal Notice2	(254, 24)
          StrCopy(mSndBuf + 253, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 24))));
          //36	Legal Notice3	(278, 24)
          StrCopy(mSndBuf + 277, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 24))));
          //37	Reserve	(302, 21)
          StrCopy(mSndBuf + 301, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 21))));
          //38	HG Data	(323, 78)
          StrCopy(mSndBuf + 322, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 78))));
          //Total 400
        end;

      NETPOS3000: //�ϥ�200bytes�榡
        begin
          //Trans_Type ����O(1, 2)
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //Host_ID �Ȧ�O(3, 2)
          StrCopy(mSndBuf +   2, '00');
          //Receipt_No EDCñ��Ǹ�(5, 6)
          StrCopy(mSndBuf +   4, '000000');
          //Card_No �H�Υd�d��(���a�k�ɪť�)(11, 19)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //Card_Expire_Date �H�Υd���Ĵ�(30, 4)
          StrCopy(mSndBuf +  29, '0000');
          //Trans_Amount ������B(34, 12)
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //Trans_Date ������(46, 6)
          StrCopy(mSndBuf +  45, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //Trans_ Time ����ɶ�(52, 6)
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //Approval_No ���v�X(���a�k�ɪť�)(58, 9)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //Auth_Amount �w�����v���B(67, 12)
          StrCopy(mSndBuf +  66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //ECR_Response_Code�q�T�^���X(79, 4)
          StrCopy(mSndBuf +  78, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //EDC_Terminal_ID EDC�ݥ����N��(83, 8)
          StrCopy(mSndBuf +  82, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //Reference_No �Ȧ����Ǹ�(91, 12)
          StrCopy(mSndBuf +  90, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Exp_Amount ��L���B(103, 12)
          StrCopy(mSndBuf + 102, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Store_Id �M�d��(115, 16)
          StrCopy(mSndBuf + 114, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 16))));
          //Amount3 ���B3(131, 12)
          StrCopy(mSndBuf + 130, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Amount4 ���B4(143, 12)
          StrCopy(mSndBuf + 142, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Inqure Type(155, 2)
          StrCopy(mSndBuf + 154, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //Product Code���~�N�X(157, 2)
          StrCopy(mSndBuf + 156, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //R+I Flag�������O(159, 1)
          StrCopy(mSndBuf + 158, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //charge fee flag���O�覡(160, 1)
          StrCopy(mSndBuf + 159, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //Delay fee flag����I�ڤ覡(161, 1)
          StrCopy(mSndBuf + 160, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //Reserved (162, 39)
          StrCopy(mSndBuf + 161, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 39))));
          //TOTAL	200
        end;

      AS320, AS320_UnionPay: //�ϥ�144bytes�榡
        begin
          //Trans_Type ����O(1, 2)
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //Host_ID �Ȧ�O(3, 2)
          if (FEdcType = AS320_UnionPay) then
            StrCopy(mSndBuf +   2, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))))
          else
            StrCopy(mSndBuf +   2, '01');
          //Receipt_No EDCñ��Ǹ�(5, 6)
          StrCopy(mSndBuf +   4, '000000');
          //Card_No �H�Υd�d��(���a�k�ɪť�)(11, 19)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //Card_Expire_Date �H�Υd���Ĵ�(30, 4)
          StrCopy(mSndBuf +  29, '0000');
          //Trans_Amount ������B(34, 12)
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //Trans_Date ������(46, 6)
          StrCopy(mSndBuf +  45, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //Trans_ Time ����ɶ�(52, 6)
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //Approval_No ���v�X(���a�k�ɪť�)(58, 9)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));


          //ECR_Response_Code�q�T�^���X(67, 4)
          StrCopy(mSndBuf +  66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //EDC_Terminal_ID EDC�ݥ����N��(71, 8)
          StrCopy(mSndBuf +  70, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //Reference_No �Ȧ����Ǹ�(79, 12)
          StrCopy(mSndBuf +  78, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));

          //Store_Id �M�d��(91, 16)
          StrCopy(mSndBuf + 90, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 16))));


          //Auth_Amount1 �w�����v���B(107, 12)
          StrCopy(mSndBuf +  106, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Auth_Amount2 �w�����v���B(119, 12)
          StrCopy(mSndBuf +  118, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Auth_Amount3 �w�����v���B(131, 12)
          StrCopy(mSndBuf +  130, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Installment Period ����(143, 2)
          StrCopy(mSndBuf + 142, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //TOTAL	144
        end;

      AS320_200: //�ϥ�200bytes�榡
        begin
          //Trans_Type ����O(1, 2)
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //Host_ID �Ȧ�O(3, 2)
          StrCopy(mSndBuf +   2, '01');
          //Receipt_No EDCñ��Ǹ�(5, 6)
          StrCopy(mSndBuf +   4, '000000');
          //Card_No �H�Υd�d��(���a�k�ɪť�)(11, 19)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //Card_Expire_Date �H�Υd���Ĵ�(30, 4)
          StrCopy(mSndBuf +  29, '****');
          //Trans_Amount ������B(34, 12)
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //Trans_Date ������(46, 6)
          StrCopy(mSndBuf +  45, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //Trans_ Time ����ɶ�(52, 6)
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //Approval_No ���v�X(���a�k�ɪť�)(58, 9)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //Auth_Amount �w�����v���B(67, 12)
          StrCopy(mSndBuf +  66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //ECR_Response_Code�q�T�^���X(79, 4)
          StrCopy(mSndBuf +  78, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //EDC_Terminal_ID EDC�ݥ����N��(83, 8)
          StrCopy(mSndBuf +  82, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //Reference_No �Ȧ����Ǹ�(91, 12)
          StrCopy(mSndBuf +  90, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Exp_Amount ��L���B(103, 12)
          StrCopy(mSndBuf + 102, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Store_Id �M�d��(115, 16)
          StrCopy(mSndBuf + 114, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 16))));
          //Amount3 ���B3(131, 12)
          StrCopy(mSndBuf + 130, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Amount4 ���B4(143, 12)
          StrCopy(mSndBuf + 142, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //Inqure Type(155, 2)
          StrCopy(mSndBuf + 154, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //Product Code���~�N�X(157, 2)
          StrCopy(mSndBuf + 156, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //R+I Flag�������O(159, 1)
          StrCopy(mSndBuf + 158, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //charge fee flag���O�覡(160, 1)
          StrCopy(mSndBuf + 159, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //Delay fee flag����I�ڤ覡(161, 1)
          StrCopy(mSndBuf + 160, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //Reserved (162, 39)
          StrCopy(mSndBuf + 161, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 39))));
        end;

        AllPay: //�ϥ�144bytes�榡
        begin
          //Trans_Type ����O(1, 2)
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //Host_ID �Ȧ�O(3, 2)
          StrCopy(mSndBuf +   2, '02');  //02�H�Υd //04���� �ثe�S�����p�d
          //Receipt_No EDCñ��Ǹ�(5, 6)
          StrCopy(mSndBuf +   4, '000000');
          //Card_No �H�Υd�d��(���a�k�ɪť�)(11, 19)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //ECR_Response_Code�q�T�^���X (30, 4)
          StrCopy(mSndBuf +  29, '0000');
          //Trans_Amount ������B(34, 10)
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 10))));
          //Trans_Date ������(44, 8)
          StrCopy(mSndBuf +  43, PAnsiChar(AnsiString(FormatDateTime('YYYYMMDD', now))));
          //Trans_ Time ����ɶ�(52, 6)
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //TradeNo ����Ǹ�(���a�k�ɪť�)(58, 20)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 20))));
          //Auth_code ���v�X (78, 6)
          StrCopy(mSndBuf +  77, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 6))));
          //EDC_Terminal_ID EDC�ݥ����N��(84, 8)
          StrCopy(mSndBuf +  83, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //Store_No �M�d��(92, 18)
          StrCopy(mSndBuf +  91, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 18))));
          //Start_Get_PAN ����O
          StrCopy(mSndBuf + 109, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //card Type�d���N�X(112, 1)
          StrCopy(mSndBuf + 111, PAnsiChar(AnsiString(copy(trim(StoreNo), 1, 1))));
          //Filler �O�d(113, 32)
          StrCopy(mSndBuf + 112, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 32))));
          //TOTAL	144
        end;
      //20170525 add by 07113
      TaiHsin_400, FirstBank_400:  //20171117 add by 07113
        begin
         //(��m,����)���W��-�Ƶ�
          //(  1,  2)Trans_Type ����O
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //(  3,  2)Host_ID �Ȧ�O
          case FEdcType of
            FirstBank_400:
              StrCopy(mSndBuf +   2, '02');           //�@��A8���e02
          else
              StrCopy(mSndBuf +   2, '00');
          end;
          //(  5,  6)Receipt_No �ݥ���ñ��Ǹ�
          StrCopy(mSndBuf +   4, '000000');
          //( 11, 19)Card_No �H�Υd�d��(���a�k�ɪť�)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //( 30,  4)Reserve �O�d
          StrCopy(mSndBuf +  29, '0000');
          //( 34, 12)Trans_Amount ������B
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //( 46,  6)Trans_Date ������
          StrCopy(mSndBuf +  45, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //( 52,  6)Trans_ Time ����ɶ�
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //( 58,  9)Approval_No ���v�X(���a�k�ɪť�)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //( 67, 12)Reserve �O�d
          StrCopy(mSndBuf +  66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //( 79,  4)ECR_Response_Code �q�T�^���X
          StrCopy(mSndBuf +  78, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //( 83,  8)EDC_Terminal_ID EDC�ݥ����N��
          StrCopy(mSndBuf +  82, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //( 91, 12)Reserve �O�d
          StrCopy(mSndBuf +  90, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(103, 12)Reserve �O�d
          StrCopy(mSndBuf + 102, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(115, 18)Store_Id/RRN or Order_No �d���B�����εo�����X
          StrCopy(mSndBuf + 114, PAnsiChar(AnsiString(copy(trim(StoreNo), 1, 6))+
                                           AnsiString(copy(trim(PosNO), 1, 6))+
                                           AnsiString(copy(trim(TrnNo), 1, 6))));
          //(133,  2)Reserve �O�d
          StrCopy(mSndBuf + 132, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(135,  2)Card_type �d�O
          StrCopy(mSndBuf + 134, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(137,  1)CUP Indicator C:CUP, N:Default
          StrCopy(mSndBuf + 136, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(138,  7)Reserve �O�d
          StrCopy(mSndBuf + 137, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 7))));
          //(145, 63)Reserve �O�d
          StrCopy(mSndBuf + 144, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 63))));
          //(208,  8)Issuer ID �o�d���N�� (Smart Pay �^�e)
          StrCopy(mSndBuf + 207, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //(216, 44)Card No. Vehicle �q�l�o���d������
          StrCopy(mSndBuf + 215, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 44))));
          //(260,  1)Transaction finished Y = ������� (���~�򱵦����) N = ����|������ (�~�򱵦����)
          StrCopy(mSndBuf + 259, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(261,  4)Timer block reset �O�ɮɶ����m Transaction finished ����=N�APOS�ݧP�_�����
          //                           1. ���� = �� ���Ρ�0000���A�ɶ����ݭ��m
          //                           2. ���� = ��0060���APOS �ݭ��s�]�m�O�ɮɶ��ñq60 ��_�l
          StrCopy(mSndBuf + 260, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //(265,  1)Signature type �v��ñ�������O ��1�� = �qñ��� ��2�� = �ȥ���� ��3�� = �Kñ���
          StrCopy(mSndBuf + 264, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(266,135)Reserve �O�d
          StrCopy(mSndBuf + 265, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 135))));
          //TOTAL	400
        end;
      //20170525 add by 07113
      //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
      ESUN_S80RF_600:
        begin
          //(��m,����)���W��-�Ƶ�
          //(  1,  2)Trans Type                 ����O
          StrCopy(mSndBuf +   0,  PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //(  3,  2)Host ID                    �Ȧ�O
          StrCopy(mSndBuf +   2,  '01');
          //(  5,  6)Receipt No                 �ݥ���ñ��Ǹ�
          StrCopy(mSndBuf +   4,  '000000');
          //( 11, 19)Card Number                �H�Υd�d��(���a�k�ɪť�)
          StrCopy(mSndBuf +  10,  '0000000000000000000');
          //( 30,  4)Expiration Date            �H�Υd���Ĵ�
          StrCopy(mSndBuf +  29,  '0000');
          //( 34, 12)Trans Amount               ������B
          StrCopy(mSndBuf +  33,  PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //( 46,  6)Trans Date                 ������
          StrCopy(mSndBuf +  45,  PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //( 52,  6)Trans Time                 ����ɶ�
          StrCopy(mSndBuf +  51,  PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //( 58,  6)Approval Code              ���v�X(���a�k�ɪť�)
          StrCopy(mSndBuf +  57,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 6))));
          //( 64,  4)ECR Response Code          �q�T�^���X
          StrCopy(mSndBuf +  63,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //( 68,  8)Terminal ID                �ݥ����N��
          StrCopy(mSndBuf +  67,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //( 76, 12)Reference No               �Ȧ����Ǹ�
          StrCopy(mSndBuf +  75,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //( 88,  7)Reserve                    �O�d
          StrCopy(mSndBuf +  87,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 7))));
          //( 95,  2)Installment Period         ��������
          StrCopy(mSndBuf +  94,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //( 97,  8)Down Payment               �������B
          StrCopy(mSndBuf +  96,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //(103,  8)Installment Payment        �C�����B
          StrCopy(mSndBuf +  104, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //(113, 10)Redeem Amt                 �����B
          StrCopy(mSndBuf +  112, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(123, 20)Store ID                   �d��
          StrCopy(mSndBuf +  122, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 20))));
          {StrCopy(mSndBuf +  122, PAnsiChar(AnsiString(copy(trim(StoreNo), 1, 8))+
                                            AnsiString(copy(trim(PosNO),   1, 6))+
                                            AnsiString(copy(trim(TrnNo),   1, 6))));  }
          //(143,  2)START Trans Type           Ū���d��������O (SALE 01, REFUND 02)(Ū���d����ƶ��e�����)
          StrCopy(mSndBuf +  142, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(145, 60)Invoice Encryption Card No �q�l�o���[�K�d��(���a�k�ɪť�)
          StrCopy(mSndBuf +  144, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 60))));
          //(205, 20)Order Number               �q��s��
          StrCopy(mSndBuf +  204, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 20))));
          //(225,150)Order Information          �q���T
          StrCopy(mSndBuf +  224, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 150))));
          //(375,  2)Code Index                 ���X�I���A
          StrCopy(mSndBuf +  374, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(377, 25)E-Coupons Product Code     �I���鲣�~�N�X
          StrCopy(mSndBuf +  376, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 25))));
          //(402, 14)E-Coupons ExpireDateTime   �I�����Ĥ���ɶ�
          StrCopy(mSndBuf +  401, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 14))));
          //(416, 14)E-coupons RedeemDateTime   �I������ɶ�
          StrCopy(mSndBuf +  415, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 14))));
          //(430,  3)E-coupons Redeem Balance   �i�I���l�B
          StrCopy(mSndBuf +  429, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 3))));
          //(433,  1)Ticket Type                �q�l���ҦW���� 1-ECC,2-iCash,3-iPASS,4-HappyCash
          StrCopy(mSndBuf +  432, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(434, 19)Ticket Card Number         �q�l���ҥd��
          StrCopy(mSndBuf +  433, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 19))));
          //(453, 10)Ticket Reference Number    �q�l���ҥ���Ǹ�
          StrCopy(mSndBuf +  452, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(463, 10)Ticket Batch Number        �妸���X
          StrCopy(mSndBuf +  462, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(473, 10)Ticket Pre-Balance         ����e�l�B
          StrCopy(mSndBuf +  472, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(483, 10)Ticket Auto load Amount    �۰ʥ[�Ȫ��B
          StrCopy(mSndBuf +  482, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(492, 10)Ticket Balance             ���O���x�Ⱦl�B
          StrCopy(mSndBuf +  492, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(503, 16)Ticket SAM ID              �w���Ҳսs��
          StrCopy(mSndBuf +  502, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 16))));
          //(519, 10)Ticket Store ID            �����s��
          StrCopy(mSndBuf +  518, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 10))));
          //(529, 72)Reserve                    �O�d
          StrCopy(mSndBuf +  528, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 72))));
        end;
      CTBC_AS320_250:
        begin
          //(��m,����)���W��-�Ƶ�
          //(  1,  2)Trans Type                 ����O
          StrCopy(mSndBuf +   0,  PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //(  3,  2)Host ID                    �Ȧ�O
          StrCopy(mSndBuf +   2,  '00');        //20181016 modi by 02953 �_����EDC�Ҧ��ҥH�e00 �ѥd����ܥd�O
          //(  5,  6)Invoice No.                �վ\�s��
          StrCopy(mSndBuf +   4,  '000000');
          //( 11, 19)Card No.                   �H�Υd�d��(���a�k�ɪť�)
          StrCopy(mSndBuf +  10,  '0000000000000000000');
          //( 30,  4)Card Exp Date              �H�Υd���Ĵ�
          StrCopy(mSndBuf +  29,  '0000');
          //( 34, 12)Trans Amt                  ������B
          StrCopy(mSndBuf +  33,  PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //( 46,  6)Trans Date                 ������
          StrCopy(mSndBuf +  45,  PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //( 52,  6)Trans Time                 ����ɶ�
          StrCopy(mSndBuf +  51,  PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //( 58,  9)Approval Code              ���v�X(���a�k�ɪť�)
          StrCopy(mSndBuf +  57,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //( 67, 12)Amount 1                   ���B1�A���p�Ʀ����]�A�p���I
          StrCopy(mSndBuf +  66,  PAnsiChar(AnsiString(copy(mTmpBuff, 1,12))));
          //( 79,  4)Resp Code                  �^�ХN�X
          StrCopy(mSndBuf +  78,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //( 83,  8)Terminal ID                �ݥ����s��
          StrCopy(mSndBuf +  82,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //( 91, 12)Ref No.                    �Ǹ�
          StrCopy(mSndBuf +  90,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(103, 12)Amount 2                   ���B2�A���p�Ʀ����]�A�p���I
          StrCopy(mSndBuf +  102,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(115, 16)Store ID                   �d��
          StrCopy(mSndBuf +  114,  PAnsiChar(AnsiString(copy(mTmpBuff, 1, 16))));     //20181019 modi by 02953 �d���a�ť�
          {StrCopy(mSndBuf +  114, PAnsiChar(AnsiString(copy(trim(StoreNo), 1, 6))+   //20181019 mark by 02953 �d���a�ť�
                                           AnsiString(copy(trim(PosNO),   1, 5))+
                                           AnsiString(copy(trim(TrnNo),   1, 5))));}
          //(131, 12)Amount 3                   ���B3�A���p�Ʀ����]�A�p���I
          StrCopy(mSndBuf +  130, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(143, 12)Amount 4                   ���B4�A���p�Ʀ����]�A�p���I
          StrCopy(mSndBuf +  142, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(155,  2)Inquiry Type               1.Inquiry �G ��01���� Sale�B��02���� Refund 2.�����I�ڴ��ơG ��00��EDC��J�B�D��00��POS��J
          StrCopy(mSndBuf +  154, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(157,  2)Product Code               ���~�N�X(00~99�A�e�ťե�EDC�W���)
          StrCopy(mSndBuf +  156, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(159,  1)R+I Flag                   �������O�G(�e0�Ϊťե�EDC�W���) 1.�L���Q�B2.���Q
          StrCopy(mSndBuf +  158, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(160,  1)Charge Fee Flag            ���O�覡(�e0�Ϊťե�EDC�W���)�i1�j�G�@����� �i2�j�G���O����
          StrCopy(mSndBuf +  159, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(161,  1)Delay Payment Flag         ����I�ڤ覡(�e0�Ϊťե�EDC�W���)�i1�j�G�L����I�ڡi2�j�G����I�ڡi3�j�G�u�ʥI��
          StrCopy(mSndBuf +  160, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(162, 50)SHA Card No                �H�Υd��q�l�o������]��Host ID : 01,02,03,04, 11,12�A�Ρ^01-06�G���d���e���X07-50�G�H�Υd����d���H SHA-256 ���ͤ� SHA �Ȫ��� > 16�G�ШϥΧ���d���h���� ���� = 16�G�ШϥΧ���d���h���� ���� < 16�G�ШϥΧ���d���h��
          StrCopy(mSndBuf +  161, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 50))));
          //(212,  1)Active E-Invoice Issuer    �ӥd���O�_�O�_���w�[�J�H�Υd��q�l�o�����㤧�o�d��0 : ���[�J1 : �w�[�J
          StrCopy(mSndBuf +  211, PAnsiChar(AnsiString(copy(mTmpBuff, 1,  1))));
          //(213, 38)Reserved                   �O�d
          StrCopy(mSndBuf +  212, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 38))));
        end;
      //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
    end;
    mPRetBuf := @mRetBuf;
    if not TxTransmit(mSndBuf, mPRetBuf) then
    begin
        DisConnect;
        if assigned(FOnEDCError) then
            FOnEDCError(self, 2, FSuspend);//'��d����ƶǰe���� !'
    end
    else
    begin
        DisConnect;
        mRetStr := String(mRetBuf);
        with mCrd do
        begin
          case (FEdcType) of
            H5000, S9000, SAGEM, ICT220_CATHAY, VX510_ESun, NETPOS3000, AS320, AS320_200, AllPay, AS320_UnionPay, SinoPac: //�ϥ�144bytes�榡
              begin
                Host_id     := trim(copy(mRetStr,  3,  2));
                Ref_no      := trim(copy(mRetStr,  5,  6));
                Card_Name   := ''; //�O�d���� //copy(mRetStr,  8,  8);
                Card_no     := trim(copy(mRetStr, 11, 19));
                if FEdcType = VX510_ESun then
                begin
                  Auth_no     := trim(copy(mRetStr, 58,  6));
                  Response    := trim(copy(mRetStr, 64,  4));
                end  else
                if (FEdcType = AS320) or (FEdcType = AS320_UnionPay) then
                begin
                  Auth_no     := trim(copy(mRetStr, 58,  9));
                  Response    := trim(copy(mRetStr, 67,  4));
                end else
                if FEdcType = AllPay then
                begin
                  Auth_no     := trim(copy(mRetStr, 78,  6));
                  Response    := trim(copy(mRetStr, 30,  4));
                end else
                begin
                  Auth_no     := trim(copy(mRetStr, 58,  9));
                  Response    := trim(copy(mRetStr, 79,  4));
                end;
              end;
            VX570_400:
              begin
                Host_id     := trim(copy(mRetStr, 12,  2));
                Card_Name   := ''; //�O�d���� //copy(mRetStr,  8,  8);
                Ref_no      := trim(copy(mRetStr, 14,  6));
                Card_no     := trim(copy(mRetStr, 20, 19));
                Auth_no     := trim(copy(mRetStr, 67,  9));
                Response    := trim(copy(mRetStr, 77,  4));
              end;
            //20170525 add by 07113
            TaiHsin_400, FirstBank_400:  //20171117 add by 07113
              begin
                Host_id     := trim(copy(mRetStr,  3,  2));
                Ref_no      := trim(copy(mRetStr,  5,  6));
                Card_no     := trim(copy(mRetStr, 11, 19));
                Auth_no     := trim(copy(mRetStr, 58,  9));
                Response    := trim(copy(mRetStr, 79,  4));
              end;
            //20170525 add by 07113
            //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
            ESUN_S80RF_600:
              begin
                Host_id     := trim(copy(mRetStr,  3,  2));
                Ref_no      := trim(copy(mRetStr, 76, 12));
                Card_no     := trim(copy(mRetStr, 11, 19));
                Auth_no     := trim(copy(mRetStr, 58,  6));
                Response    := trim(copy(mRetStr, 64,  4));
                TransDate               := trim(copy(mRetStr, 46,   6));
                TransAmt                := trim(copy(mRetStr, 34,  12));
                InvoiceEncryptionCardNo := trim(copy(mRetStr,145,  60));
              end;
            CTBC_AS320_250:
              begin
                Host_id     := trim(copy(mRetStr,  3,  2));
                Ref_no      := trim(copy(mRetStr, 91, 12));
                Card_no     := trim(copy(mRetStr, 11, 19));
                Auth_no     := trim(copy(mRetStr, 58,  9));
                Response    := trim(copy(mRetStr, 79,  4));
                TransDate               := trim(copy(mRetStr, 46,  6));
                TransAmt                := trim(copy(mRetStr, 34, 12));
                SHACardNo               := trim(copy(mRetStr,162, 50));
              end;
            //20180621 add  by 02953 for BQ34�]�Ϊk�H�_0000679978_POS_0002  ��
          end;
          if Response <> '0000' then
          begin
            Host_id     := '';
            Card_Name   := '';
            Card_no     := '';
            Auth_no     := '';
            Ref_no      := '';
          end;
        end;
        Result := true;
        if Assigned(FOnGetEDC) then
            FOnGetEDC(self, mCrd);
    end;
end;

function TEDCCom.SendToEDC2(TransType, Payment, StoreNo, PosNo, TrnNo: string;
  CardType: Integer): boolean;
const
    second = 0.00001157407;
var
    mRetBuf, mSndBuf, mTmpBuff: array[0..MAX_DATA] of AnsiChar;
    mPRetBuf: PChar;
    mCrd: TCreditCard;
    mPayNum, mTenantNo, mRetStr: string;
    SL: TStringList;
    mTemp:String;
    mBgTime:TDateTime;
    mPath, ApRunPath:String;
begin
    Result := false;
    FState := 0;
//�b�ϥΦ���󤧫e���}��....AFile
  Notification := TATFileNotification.Create(self);    //20171122 edc.exe
  Notification.OnChanged:= NotificationChanged;        //20171122 edc.exe
  ApRunPath:=ExtractFilePath(Application.ExeName);
  if FileExists(ApRunPath + EDCPath+'\'+FileName) then
  begin
    //1 �զ��r��
    FillChar(mTmpBuff, FEDCDataLength, AnsiChar(' '));
    case (FEdcType) of

      VEGA_9000:
        begin
          //1	ECR Card Type�d���O	10,20,30,40,99 [�H�Υd,���ĥd,����B�q,���p�d,EINVO_CARD]
          if CardType = 1 then
            StrCopy(mSndBuf +   0, '10')   //�H�Υd
          else
          if CardType = 2 then
            StrCopy(mSndBuf +   0, '20')   //���ĥd
          else
          if CardType = 3 then
            StrCopy(mSndBuf +   0, '30')   //����B�q
          else
          if CardType = 4 then
            StrCopy(mSndBuf +   0, '40')   //���p�d
          else
            StrCopy(mSndBuf +   0, '10');  //��L

          //2	ECR Transaction Type����O	(2, 2)   **���n
          StrCopy(mSndBuf +   2, '11');
          //3	Trans Trans Amount������B	(4, 12)  **���n
          StrCopy(mSndBuf +   4,  PAnsiChar(copy(trim(Payment), 1, 12)));
          //4	ROC�վ\�s��	(16, 12)
          StrCopy(mSndBuf +  16,  PAnsiChar(copy(mTmpBuff, 1, 12)));
          //5	Quantity �ƶq	(28, 10)
          StrCopy(mSndBuf +  28,  PAnsiChar(copy(mTmpBuff, 1, 10)));
          //6	Coupon Id�I��N (38, 10)
          StrCopy(mSndBuf +  38,  PAnsiChar(copy(mTmpBuff, 1, 10)));
          //7	Store Id�d��,����,�o�����X(48, 18)
          StrCopy(mSndBuf +  48,  PAnsiChar(copy(trim(StoreNo), 1, 6) + copy(trim(PosNO), 1, 6)+ copy(mTmpBuff, 1, 6)));
          //8	Condition�ثe���O�d��(66, 1)
          StrCopy(mSndBuf +  66,  PAnsiChar(copy(mTmpBuff, 1, 1)));
          //9	Continue_Flag �ثe���O�d��(67, 1)
          StrCopy(mSndBuf +  67,  PAnsiChar(copy(mTmpBuff, 1, 1)));
          //10	����ɶ�        (68, 12)  **���n (�a���ɪű�)
          StrCopy(mSndBuf +  68, PAnsiChar(FormatDateTime('MMDDhhmmss', now)+ copy('   ', 1, 2)));
        end;

      TaiHsin_400:
        begin
          //(  1,  2)Trans_Type ����O
          StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(copy(trim(TransType), 1, 2))));
          //(  3,  2)Host_ID �Ȧ�O
          StrCopy(mSndBuf +   2, '03');
          //(  5,  6)Receipt_No �ݥ���ñ��Ǹ�
          StrCopy(mSndBuf +   4, '000000');
          //( 11, 19)Card_No �H�Υd�d��(���a�k�ɪť�)
          StrCopy(mSndBuf +  10, '0000000000000000000');
          //( 30,  4)Reserve �O�d
          StrCopy(mSndBuf +  29, '0000');
          //( 34, 12)Trans_Amount ������B
          StrCopy(mSndBuf +  33, PAnsiChar(AnsiString(copy(trim(Payment), 1, 12))));
          //( 46,  6)Trans_Date ������
          StrCopy(mSndBuf +  45, PAnsiChar(AnsiString(FormatDateTime('YYMMDD', now))));
          //( 52,  6)Trans_ Time ����ɶ�
          StrCopy(mSndBuf +  51, PAnsiChar(AnsiString(FormatDateTime('hhmmss', now))));
          //( 58,  9)Approval_No ���v�X(���a�k�ɪť�)
          StrCopy(mSndBuf +  57, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 9))));
          //( 67, 12)Reserve �O�d
          StrCopy(mSndBuf +  66, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //( 79,  4)ECR_Response_Code �q�T�^���X
          StrCopy(mSndBuf +  78, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //( 83,  8)EDC_Terminal_ID EDC�ݥ����N��
          StrCopy(mSndBuf +  82, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //( 91, 12)Reserve �O�d
          StrCopy(mSndBuf +  90, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(103, 12)Reserve �O�d
          StrCopy(mSndBuf + 102, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 12))));
          //(115, 18)Store_Id/RRN or Order_No �d���B�����εo�����X
          StrCopy(mSndBuf + 114, PAnsiChar(AnsiString(copy(trim(StoreNo), 1, 6))+
                                           AnsiString(copy(trim(PosNO), 1, 6))+
                                           AnsiString(copy(trim(TrnNo), 1, 6))));
          //(133,  2)Reserve �O�d
          StrCopy(mSndBuf + 132, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(135,  2)Card_type �d�O
          StrCopy(mSndBuf + 134, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 2))));
          //(137,  1)CUP Indicator C:CUP, N:Default
          StrCopy(mSndBuf + 136, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(138,  7)Reserve �O�d
          StrCopy(mSndBuf + 137, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 7))));
          //(145, 63)Reserve �O�d
          StrCopy(mSndBuf + 144, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 63))));
          //(208,  8)Issuer ID �o�d���N�� (Smart Pay �^�e)
          StrCopy(mSndBuf + 207, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 8))));
          //(216, 44)Card No. Vehicle �q�l�o���d������
          StrCopy(mSndBuf + 215, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 44))));
          //(260,  1)Transaction finished Y = ������� (���~�򱵦����) N = ����|������ (�~�򱵦����)
          StrCopy(mSndBuf + 259, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(261,  4)Timer block reset �O�ɮɶ����m Transaction finished ����=N�APOS�ݧP�_�����
          //                           1. ���� = �� ���Ρ�0000���A�ɶ����ݭ��m
          //                           2. ���� = ��0060���APOS �ݭ��s�]�m�O�ɮɶ��ñq60 ��_�l
          StrCopy(mSndBuf + 260, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 4))));
          //(265,  1)Signature type �v��ñ�������O ��1�� = �qñ��� ��2�� = �ȥ���� ��3�� = �Kñ���
          StrCopy(mSndBuf + 264, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 1))));
          //(266,135)Reserve �O�d
          StrCopy(mSndBuf + 265, PAnsiChar(AnsiString(copy(mTmpBuff, 1, 135))));
          //TOTAL	400
        end;
    end;

    //2.�s���ɮ�
    SL := TStringList.Create;
    try
      SL.Add(mSndBuf);
      SL.SaveToFile(ApRunPath + EDCPath +'\'+ FileName);
      if FDebugMode then
        DebugLog('Snd>' + String(mSndBuf));                                   //�����e�XBUF
    finally
      SL.Free;
    end;

    //3.����SystexEcr
    FOnEDCMsg(self, 2, 30, FSuspend); //'POS��ƶǰe��,�еy��';
    NotifyFile(ApRunPath + ANotiEDCPath);                                       //20171122 edc.exe
    RUN_DOS_COMMAND(ApRunPath+EDCPath+ '\' + 'ECR.exe', ApRunPath+EDCPath, '', 0) ;

    //4.Ū�� ecr_out �^���T��
    FTotSec := 0;
    mBgTime := Now;
    while (FTotSec <= FEDCTimeOut) do
    begin
      if assigned(FOnEDCMsg) then
          FOnEDCMsg(self, 5, FTotSec, FSuspend); //'EDC��Ʊ�����!';

      FTotSec := StrToInt(Format('%3.0f', [(now - mBgTime) /second])) ;

      //�i�J�Ĥ@�q
      if FState = 1 then
        Break;

      //�o�Ͳ��`
      if (FState in [2, 3]) then
      begin
        Result := True ;
        break;
      end;
    end;

    //Timeout
    //20171207 modi by 01753 for Cust-20171207001 ��
    if (FTotSec <= FEDCTimeOut) and (FState  <> 1) then
    begin
      FOnEDCError(self, 5, FSuspend);//'EDC�W�X�s�u�ɶ�!'
      Result := False ;
      exit;
    end;
    //20171207 modi by 01753 for Cust-20171207001 ��

    if FState =1 then
    begin
      mBgTime := Now;
      while (FTotSec <= FEDCTimeOut) do
      begin
        if assigned(FOnEDCMsg) then
            FOnEDCMsg(self, 5, FTotSec, FSuspend); //'EDC��Ʊ�����!';

        FTotSec := StrToInt(Format('%3.0f', [(now - mBgTime) /second])) ;

        if (FState in [2, 3]) then
        begin
          Result := True ;
          break;
        end;
      end;
    end;
  end else begin
    if assigned(FOnEDCError) then
      FOnEDCError(self, 6, FSuspend); //Result := EDC���v�{�����s�b !'     //20171207 modi by 01753 for Cust-20171207001
  end;
end;

function TEDCCom.SendToEDCV3(TransType, Payment, HostID: AnsiString; FSLength,
        PaymentS,   //�I�ڰ_�l��m
        PaymentE,   //�I�ںI���m
        ResponseS,  //�q�T���Ұ_�l��m
        ResponseE,  //�q�T���ҺI���m
        AppS,       //���v�X�_�l��m
        AppE,       //���v�X�I���m
        CardS,      //�d���_�l��m
        CardE       //�d���I���m
        :integer ): boolean;
const
    second = 0.00001157407;
var

    mRetBuf, mSndBuf, mTmpBuff: array[0..MAX_DATA] of AnsiChar;
    mPRetBuf: PAnsiChar;
    mCrd: TCreditCard;
    mPayNum, mTenantNo, mRetStr: AnsiString;
begin

    if (not Connect) or (not Connected) then
    begin
        DisConnect;
        if assigned(FOnEDCError) then
            FOnEDCError(self, 1, FSuspend);//'��d���s�u����!';
        exit;
    end;


    Result := false;
    FState := 0;
    { TODO : 1.1 ����r�굹���ť� }
    FillChar(mTmpBuff, FSLength, AnsiChar(' '));
    { TODO : 1.2 �⥲������g���� ���B�B����O�B�������}
    StrCopy(mSndBuf +   0, PAnsiChar(AnsiString(trim(TransType))));
    StrCopy(mSndBuf +   2, PAnsiChar(AnsiString(trim(HostID))));
    StrCopy(mSndBuf +   4, PAnsiChar(AnsiString(copy(mTmpBuff, 1, PaymentS-4))));
    StrCopy(mSndBuf +   PaymentS -1 , PAnsiChar(AnsiString(copy(Payment, 1, PaymentE))));
    StrCopy(mSndBuf +   PaymentS+PaymentE -1 , PAnsiChar(AnsiString(copy(mTmpBuff, 1, FSLength - (PaymentE+PaymentS-1)))));
    { TODO : �o�ӭn�h��#0 �_�h�|�����D }
    { TODO : 2 �e�X }
    mPRetBuf := @mRetBuf;
    if not TxTransmit(mSndBuf, mPRetBuf) then
    begin
      DisConnect;
      if assigned(FOnEDCError) then
          FOnEDCError(self, 2, FSuspend);//'��d����ƶǰe���� !'
    end
    else
    begin
      DisConnect;
      mRetStr := String(mRetBuf);
      with mCrd do
      begin
        case (FEdcType) of
          UniversalEDC:
            begin
              Host_id     := trim(copy(mRetStr,  3,  2));
              Ref_no      := trim(copy(mRetStr,  5,  6));
              Card_Name   := '';
              Card_no     := trim(copy(mRetStr, CardS, CardE));
              Auth_no     := trim(copy(mRetStr, AppS,  AppE));
              Response    := trim(copy(mRetStr, ResponseS,  ResponseE));
            end;
        end;

        if Response <> '0000' then
        begin
          Host_id     := '';
          Card_Name   := '';
          Card_no     := '';
          Auth_no     := '';
          Ref_no      := '';
        end;
      end;
      Result := true;
      if Assigned(FOnGetEDC) then
          FOnGetEDC(self, mCrd);
    end;
end;

function TEDCCom.TxTransmit(const xBuffer: PAnsiChar; var xRetValue: PAnsiChar): boolean;
const
    second = 0.00001157407;
var
    mExit, mSendOk: boolean;
    mLoop: integer;
    mSndBuf, mRcvBuf: array[0..MAX_DATA] of Ansichar;
    mLrc, mGetChar: Ansichar;
    i, mCounter: integer;
    mSndLen, mRcvLen: integer;
    mBgTime: TDateTime;
    mTotSec: integer;
begin
    Result := false;
    //�NmSndBuf, mRcvBuf�M��;
    FillChar(mSndBuf, SizeOf(mSndBuf), #0);
    FillChar(mRcvBuf, SizeOf(mRcvBuf), #0);
    mSndLen := strlen(xBuffer);
    strcopy(mSndbuf + 1, xBuffer);
    mSndbuf[0] := STX;
    mSndbuf[mSndLen + 1] := ETX;
    mSndbuf[mSndLen + 2] := Cal_Lrc(mSndbuf + 1, mSndLen + 1);
    mSndbuf[mSndLen + 3] := NULL;
    mTotSec := 0; mLoop := 0;
    mSendOk := false;
    mBgTime := Now;

    if FDebugMode then DebugLog('Snd>' + AnsiString(mSndbuf));
    if assigned(FOnEDCMsg) then
        FOnEDCMsg(self, 1, mTotSec, FSuspend); //'�ХѨ�d����H�Υd!';

    while ((mTotSec <= FEDCTimeOut) and (mLoop < 5)) do
    begin
        if FSuspend then exit;

        if assigned(FOnEDCMsg) then
            FOnEDCMsg(self, 2, mTotSec, FSuspend); //'POS��ƶǰe��,�еy��';

        while (getc > NULL) do
          if FSuspend then exit;

        mSendOk := false;
        mRcvLen := puts(PAnsiChar(strpas(mSndBuf)), mSndLen + 3);    //�ǰe������
        if mRcvLen <> mSndLen + 3 then
        begin
            inc(mLoop);
            Continue;
        end;

        Result := false;
        if assigned(FOnEDCMsg) then
            FOnEDCMsg(self, 3, mTotSec, FSuspend); //'����EDC�^����!';

        mCounter  := 0;
        mExit     := false;

        while mTotSec <= FEDCTimeOut do
        begin
            if FSuspend then exit;

            mGetChar := getc;     //����EDC���
            case mGetChar of
               NAK:
                    begin
                        mSendOk := false;
                        if assigned(FOnEDCError) then
                            FOnEDCError(self, 3, FSuspend); //'EDC�^���������~!';
                    end;
               ACK, STX:
                    begin
                        mSendOk := true;
                        mCounter := 0;
                        if assigned(FOnEDCMsg) then
                            FOnEDCMsg(self, 4, mTotSec, FSuspend); //'EDC�^���������\!';
                    end;
               else
                    begin
                        mSendOk := true;
                        if assigned(FOnEDCMsg) then
                            FOnEDCMsg(self, 5, mTotSec, FSuspend); //'����EDC��Ƥ�!';
                    end;
            end;

            if not mSendOk then
            begin
                Inc(mLoop);
                break;
            end;

            if mGetChar <> char(-1) then
            begin
               if (mGetChar <> ACK) or (mCounter > 0) then
               begin
                   mRcvBuf[mCounter] := mGetChar;
                   if mGetChar = ETX then
                   begin
                      Inc(mCounter);
                      mGetChar := getc;
                      mRcvBuf[mCounter] := mGetChar;
                      mExit := true;
                      Result := true;
                   end;
                   Inc(mCounter);
               end;
            end;

            if mExit then
            begin
                mLrc := Cal_Lrc(mRcvBuf + 1, mCounter - 2);

                if mLrc <> mRcvBuf[mCounter - 1] then  //�YLrc���Ҥ��ūh���s�ǰe����
                begin
                    if assigned(FOnEDCError) then
                        FOnEDCError(self, 3, FSuspend); //'EDC�^���������~!';
                    FillChar(mRcvBuf, SizeOf(mRcvBuf), #0);
                    mExit    := false;
                    Result   := false;
                    mCounter := 0;

                    puts(NAK, 1);
                    puts(NAK, 1);
                end
                else
                begin
                    puts(ACK, 1);
                    puts(ACK, 1);
                    for i := 0 to mCounter - 3 do
                    begin
                        if mRcvBuf[i + 1] <> NULL then
                            xRetValue[i] := mRcvBuf[i + 1]
                        else
                            xRetValue[i] := ' ';
                    end;
                    xRetValue[mCounter - 3] := NULL;//'\0';

                    if FDebugMode then DebugLog('Rcv>' + AnsiString(xRetValue));
                    if assigned(FOnEDCMsg) then
                        FOnEDCMsg(self, 6, mTotSec, FSuspend); //'EDC��Ʊ������\!';

                    Result := true;
                    exit;
                end;
            end;
            mTotSec := StrToInt(Format('%3.0f', [(now - mBgTime) / second])) ;
        end;
        mTotSec := StrToInt(Format('%3.0f', [(now - mBgTime) /second])) ;
    end;

    if not Result then
      if assigned(FOnEDCError) then
          FOnEDCError(self, 4, FSuspend); //'EDC��Ʊ�������';
end;

function TEDCCom.Cal_Lrc(xValue: PAnsiChar; xLen: integer): Ansichar;
var
    i: integer;
begin
    Result := NULL;
    for i := 0 to xLen - 1 do
        Result :=  AnsiChar(integer(Result) xor integer(xValue[i]));
end;

procedure TEDCCom.DebugLog(Msg: AnsiString);
var
    mFileHandle, mLength: integer;
    mNoIndex: AnsiChar;
    mLogName: string;
    mToday: AnsiString;
    mStatment: PAnsiChar;
begin
    mToday := AnsiString(FormatDateTime('YYYY/mm/dd hh:nn:ss', Now()));
    mLogname := 'EDC.log';
    try
        if not FileExists(mLogName) then
            mFileHandle := FileCreate(mLogName)
        else
            mFileHandle := FileOpen(mLogName, fmOpenReadWrite);

        FileSeek(mFileHandle, 0, 2);
        mStatment := PAnsiChar(mToday + Msg + AnsiChar($0d) + AnsiChar($0a)) ;
        FileWrite(mFileHandle, mStatment^, Length(mStatment));
    finally
        FileClose(mFileHandle);
    end;
end;

//TKeyHook

function KeyboardHookProcWinKey(xCode: Integer; xWParam: WPARAM; xLParam: LPARAM): LRESULT; stdcall;
var
    mPassing: boolean;
    mLParam: PKBDLLHOOKSTRUCT;
    mShift: DWORD;
    function ConvertNum(xScanCode: LongInt): LongInt;
    begin
        Result := xScanCode;
        case mLParam^.vkCode of
            VK_TAB:     Result := 0;
            VK_LWIN:    Result := 23387;
            VK_RWIN:    Result := 0;
            VK_CAPITAL: Result := 14868;
            VK_APPS:    Result := 28901;
        end;
    end;

begin
  if xCode < 0 then
    Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam)
  else
  begin
    mPassing := false;
    case xWParam of
        WM_KEYDOWN, WM_SYSKEYDOWN, WM_KEYUP, WM_SYSKEYUP:
            begin
                mLParam := PKBDLLHOOKSTRUCT(xLParam);
                if (mLParam^.vkCode in [VK_LWIN, VK_RWIN, VK_APPS]) then
                    mPassing := true;
            end;
    end;
    if mPassing then
        Result := 1
    else
        Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam);
  end;
end;

procedure SetKeyState(xVirtualKey: integer);
var
  KBS : TKeyboardState;
begin
  GetKeyboardState(KBS);

  if xVirtualKey = VK_SHIFT then
  begin
    KeyHook.FShiftKey := not KeyHook.FShiftKey;
    KBS[VK_SHIFT] := strtoint(Booltostr(KeyHook.FShiftKey));
  end;

  if xVirtualKey = VK_CONTROL then
  begin
    KeyHook.FCtrlKey := not KeyHook.FCtrlKey;
    KBS[VK_CONTROL] := strtoint(Booltostr(KeyHook.FCtrlKey));
  end;

  if xVirtualKey = VK_MENU then
  begin
    KeyHook.FAltKey := not KeyHook.FAltKey;
    KBS[VK_MENU] := strtoint(Booltostr(KeyHook.FAltKey));
  end;

  SetKeyboardState(KBS);
end;

//function CheckKeyState(xVirtKey: integer): boolean; //20121017 mark by 3188
function CheckKeyState(xVirtKey: integer; xKBS: TKeyboardState): boolean;
begin
  Result := false;
//  if KeyHook.FPreSysKeyUp then
//  begin
//    case xVirtKey of
//      VK_SHIFT:
//        Result := KeyHook.FShiftKey;
//      VK_CONTROL:
//        Result := KeyHook.FCtrlKey;
//      VK_MENU:
//        Result := KeyHook.FAltKey;
//    end;
//  end
//  else
//  if (GetKeyState(xVirtKey) and $80) <> 0 then //20120807 mark by 3188 �󴫨���L���A��
//    Result := true;


  Result := False;                                       //20150709 modi by 04707
  if (xKBS[xVirtKey] and 128) <> 0 then Result := True;  //20150709 modi by 04707
end;


//20150519 add by 01753
function KeyboardHookProc(xCode: Integer; xWParam: WPARAM; xLParam: LPARAM): LRESULT; stdcall;
var
  mScanCode: Longint;
  mShift, mAlt, mCtrl, mCapital: DWORD;
  mKeyState: TShiftState;
  mKeyChar: Ansichar;
  mVKeyChar : array[1..2] of AnsiChar;
  KBS : TKeyboardState;

  procedure ResetKeyStateA;
  begin
    KeyHook.FShiftKey := false;
    KeyHook.FCtrlKey := false;
    KeyHook.FAltKey := false;
    //����W�����]�w�O�L�Ī��C�S����k�u�������L���]�w��  = -1 �N����U  ��L�Ϥ�
    KBS[VK_SHIFT] := 0;
    KBS[VK_CONTROL] := 0;
    KBS[VK_MENU] := 0;
    SetKeyboardState(KBS);
  end;
begin
  if xCode <= 0 then  //20150709 modi by 04707 �_�l�N�X���i���s  < �� <=
    Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam)
  else
  begin
    if ((xLParam and $80000000) = 0) then  //down
    begin
      if not(xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU, VK_CAPITAL]) then //20150709 add by 04707 ���Q��CapsLock�i�ӡA�]��KBS�N�i�H���o���A
      begin
        mKeyState := [];
        fillchar(mVKeyChar, SizeOf(mVKeyChar), #0);
        GetKeyboardState(KBS);

        if ToAsciiEx(xWParam, xLParam , KBS, @mVKeyChar, 0, GetKeyboardLayout(0)) <> 0 then
          mKeyChar := mVKeyChar[1]
        else
          mKeyChar := #0;

        mScanCode := ((xLParam and $00ff0000) div $100) + (xWParam and $ff);


        //20140407 mark by 01753 �n�ϵ��ѥ����ӻ��A�ڡC������n���o�q
        //�O�_�w���U SHIFT
        //if CheckKeyState(VK_SHIFT) then //20121017 mark by 3188
        if CheckKeyState(VK_SHIFT, KBS) then
        begin
          if ord(mKeyChar) in [48..57] then //�Ʀr��+SHIFT�ܧ󬰲Ÿ�
          begin
            case mKeyChar of
              '0':  mKeyChar := ')';
              '1':  mKeyChar := '!';
              '2':  mKeyChar := '@';
              '3':  mKeyChar := '#';
              '4':  mKeyChar := '$';
              '5':  mKeyChar := '%';
              '6':  mKeyChar := '^';
              '7':  mKeyChar := '&';
              '8':  mKeyChar := '*';
              '9':  mKeyChar := '(';
            end;
            mShift := 0;
          end
          else
          begin
            mShift := MapVirtualKey(VK_SHIFT, 0);
            Include(mKeyState, ssShift);
            //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
//            if ord(mKeyChar) in [65..90] then //A-Z
//              mKeyChar := AnsiChar(ord(mKeyChar) + 32)
//            else
//            if ord(mKeyChar) in [97..122] then //a-z
//              mKeyChar := AnsiChar(ord(mKeyChar) - 32);
            //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
          end;
        end
        else
          mShift := 0;
        //20140407 mark by 01753 �n�ϵ��ѥ����ӻ��A�ڡC������n���o�q


        //�O�_�w���U CONTROL
        //if CheckKeyState(VK_CONTROL) then //20121017 mark by 3188
        if CheckKeyState(VK_CONTROL, KBS) then
        begin
          mCtrl := MapVirtualKey(VK_CONTROL, 0);
          Include(mKeyState, ssCtrl);
        end
        else
          mCtrl := 0;

        //�O�_�w���U ALT
        //if CheckKeyState(VK_MENU) then  //20121017 mark by 3188
        if CheckKeyState(VK_MENU, KBS) then
        begin
          mAlt := MapVirtualKey(VK_MENU, 0);
          Include(mKeyState, ssAlt);
        end
        else
          mAlt := 0;
        //20150608 ADD BY 01753 �]�����H��cpaslock on �@�ߤj�g
        if Odd(GetKeyState(VK_CAPITAL)) then
        if ord(mKeyChar) in [97..122] then //A-Za
          mKeyChar := AnsiChar(ord(mKeyChar) - 32);
        //20150608 ADD BY 01753

        if Assigned(KeyHook.OnHookKeyUpEvent) then
          KeyHook.OnHookKeyUpEvent(KeyHook, mScanCode + mShift + mCtrl + mAlt, mKeyChar, mKeyState);
      end else
      if xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU, VK_CAPITAL] then  //20150709 add by 04707 �W�[ CapsLock
      begin
        if not KeyHook.FShiftState then
        begin
          KeyHook.FShiftState := True;
        end;
        SetKeyState(xWParam);
      end;

      case xWParam of
        VK_SHIFT,{ VK_RETURN,} VK_TAB, VK_BACK, VK_CLEAR, VK_SPACE, VK_ESCAPE,
        VK_PRIOR, VK_NEXT,   VK_LEFT,VK_RIGHT,VK_UP,    VK_DOWN,  VK_CANCEL,
        VK_CAPITAL, VK_MENU, //20150709 add by 04707 �W�[ CapsLock & Alt (���ⱱ�����M�n���U�ӹ_�l��)
        VK_ACCEPT, VK_CONTROL: //�W�[CTRL
        BEGIN
          Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam);
          KeyHook.FSendkey := false;
        END else
        begin
          //�D�S��\������L�� �����^�ǹ_�l�^�h���t�ΡC�������񵹤U�@�ӹ_�l����C
          if not KeyHook.FSendkey then
          BEGIN
            if Odd(GetKeyState(VK_CAPITAL)) then
            begin
              GetKeyboardState(KBS); //20150709 add by 04707 �����oKBS�~�i�]�wKBS,���M��L���A�|�j��
              KBS[VK_SHIFT] := 0;
              SetKeyboardState(KBS);
            end;
            Result := CallNextHookEx(KeyHook.FHook, xCode, xWParam, xLParam)
          end else
            result := 1;

          KeyHook.FSendkey := false;
        end;
      end;
    end else
    begin
      Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam);
      if ((GetKeyState(VK_SHIFT) and (1 shl 15)) <> 0) then
      begin
        //SetKeyState(xWParam);
        KeyHook.FShiftState := False;
      end;
    end;
  end;
end;
//20150519 add by 01753

//20150519 mark by 01753
//20150122 add by 01753
//function KeyboardHookProc(xCode: Integer; xWParam: WPARAM; xLParam: LPARAM): LRESULT; stdcall;
//var
//  mScanCode: Longint;
//  mShift, mAlt, mCtrl, mCapital: DWORD;
//  mKeyState: TShiftState;
//  mKeyChar: Ansichar;
//  mVKeyChar : array[1..2] of AnsiChar;
//  KBS : TKeyboardState;
//  procedure ResetKeyState;
//  begin
//    KeyHook.FShiftKey := false;
//    KeyHook.FCtrlKey := false;
//    KeyHook.FAltKey := false;
//    //����W�����]�w�O�L�Ī��C�S����k�u�������L���]�w��  = -1 �N����U  ��L�Ϥ�
//    KBS[VK_SHIFT] := 0;
//    KBS[VK_CONTROL] := 0;
//    KBS[VK_MENU] := 0;
//    SetKeyboardState(KBS);
//  end;
//begin
//  if xCode < 0 then
//    Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam)
//  else
//  begin
//    if ((xLParam and $80000000) = $80000000) then     //(lparam and $80000000= 0) -->down  <>0 --> up;
//    begin
//
//      if not(xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU]) then
//      begin
//        mKeyState := [];
//        fillchar(mVKeyChar, SizeOf(mVKeyChar), #0);
//        GetKeyboardState(KBS);
//
//        //if ToAsciiEx(xWParam, (xLParam and $00ff0000) , KBS, @mVKeyChar, 0, GetKeyboardLayout(0)) <> 0 then
//
//        if ToAsciiEx(xWParam, (xLParam and $00ff0000) shr 16 , KBS, @mVKeyChar, 0, GetKeyboardLayout(0)) <> 0 then
//          mKeyChar := mVKeyChar[1]
//        else
//          mKeyChar := #0;
//
//        mScanCode := ((xLParam and $00ff0000) div $100) + (xWParam and $ff);
//
//
//        //20140407 mark by 01753 �n�ϵ��ѥ����ӻ��A�ڡC������n���o�q
//        //�O�_�w���U SHIFT
//        //if CheckKeyState(VK_SHIFT) then //20121017 mark by 3188
//        if CheckKeyState(VK_SHIFT, KBS) then
//        begin
//          if ord(mKeyChar) in [48..57] then //�Ʀr��+SHIFT�ܧ󬰲Ÿ�
//          begin
//            case mKeyChar of
//              '0':  mKeyChar := ')';
//              '1':  mKeyChar := '!';
//              '2':  mKeyChar := '@';
//              '3':  mKeyChar := '#';
//              '4':  mKeyChar := '$';
//              '5':  mKeyChar := '%';
//              '6':  mKeyChar := '^';
//              '7':  mKeyChar := '&';
//              '8':  mKeyChar := '*';
//              '9':  mKeyChar := '(';
//            end;
//            mShift := 0;
//          end
//          else
//          begin
//            mShift := MapVirtualKey(VK_SHIFT, 0);
//            Include(mKeyState, ssShift);
//            //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
//            if ord(mKeyChar) in [65..90] then //A-Z
//              mKeyChar := AnsiChar(ord(mKeyChar) + 32)
//            else
//            if ord(mKeyChar) in [97..122] then //a-z
//              mKeyChar := AnsiChar(ord(mKeyChar) - 32);
//            //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
//          end;
//        end
//        else
//          mShift := 0;
//        //20140407 mark by 01753 �n�ϵ��ѥ����ӻ��A�ڡC������n���o�q
//
//
//        //�O�_�w���U CONTROL
//        //if CheckKeyState(VK_CONTROL) then //20121017 mark by 3188
//        if CheckKeyState(VK_CONTROL, KBS) then
//        begin
//          mCtrl := MapVirtualKey(VK_CONTROL, 0);
//          Include(mKeyState, ssCtrl);
//        end
//        else
//          mCtrl := 0;
//
//        //�O�_�w���U ALT
//        //if CheckKeyState(VK_MENU) then  //20121017 mark by 3188
//        if CheckKeyState(VK_MENU, KBS) then
//        begin
//          mAlt := MapVirtualKey(VK_MENU, 0);
//          Include(mKeyState, ssAlt);
//        end
//        else
//          mAlt := 0;
//
//        //ResetKeyState;
//        if Assigned(KeyHook.OnHookKeyUpEvent) then
//          KeyHook.OnHookKeyUpEvent(KeyHook, mScanCode + mShift + mCtrl + mAlt, mKeyChar, mKeyState);
//      end else
//      if xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU] then
//      begin
//        SetKeyState(xWParam);
//      end;
//
//      case xWParam of
//        VK_SHIFT, VK_RETURN, VK_TAB, VK_BACK, VK_CLEAR, VK_SPACE, VK_ESCAPE,
//        VK_PRIOR, VK_NEXT,   VK_LEFT,VK_RIGHT,VK_UP,    VK_DOWN,  VK_CANCEL,
//        VK_ACCEPT:
//        BEGIN
//          Result := CallNextHookEx(KeyHook.FHook, xCode, xWParam, xLParam);
//        END else
//        begin
//          //�D�S��\������L�� �����^�ǹ_�l�^�h���t�ΡC�������񵹤U�@�ӹ_�l����C
//          CallNextHookEx(KeyHook.FHook, xCode, xWParam, xLParam);
//          result := 1;
//        end;
//      end;
//    end else
//    begin
//      Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam);
//    end;
//  end;
//end;
//20150122 add by 01753

{
//mark by 4084 20131203
function KeyboardHookProc(xCode: Integer; xWParam: WPARAM; xLParam: LPARAM): LRESULT; stdcall;
var
  mScanCode: Longint;
  mShift, mAlt, mCtrl, mCapital: DWORD;
  mKeyState: TShiftState;
  mKeyChar: Ansichar;
  mVKeyChar : array[1..2] of AnsiChar;
  KBS : TKeyboardState;
  procedure ResetKeyState;
  begin
    KeyHook.FShiftKey := false;
    KeyHook.FCtrlKey := false;
    KeyHook.FAltKey := false;
  end;
begin
  if xCode <= 0 then
    Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam)
  else
  begin
    if ((xLParam and $80000000) = $80000000) then
    begin
      if not(xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU]) then
      begin
        mKeyState := [];
        fillchar(mVKeyChar, SizeOf(mVKeyChar), #0);
        GetKeyboardState(KBS);


        if ToAsciiEx(xWParam, (xLParam and $00ff0000), KBS, @mVKeyChar, 0, GetKeyboardLayout(0)) <> 0 then
          mKeyChar := mVKeyChar[1]
        else
          mKeyChar := #0;

        mScanCode := ((xLParam and $00ff0000) div $100) + (xWParam and $ff);


        //20140407 mark by 01753 �n�ϵ��ѥ����ӻ��A�ڡC������n���o�q
        //�O�_�w���U SHIFT
        //if CheckKeyState(VK_SHIFT) then //20121017 mark by 3188
        if CheckKeyState(VK_SHIFT, KBS) then
        begin
          if ord(mKeyChar) in [48..57] then //�Ʀr��+SHIFT�ܧ󬰲Ÿ�
          begin
            case mKeyChar of
              '0':  mKeyChar := ')';
              '1':  mKeyChar := '!';
              '2':  mKeyChar := '@';
              '3':  mKeyChar := '#';
              '4':  mKeyChar := '$';
              '5':  mKeyChar := '%';
              '6':  mKeyChar := '^';
              '7':  mKeyChar := '&';
              '8':  mKeyChar := '*';
              '9':  mKeyChar := '(';
            end;
            mShift := 0;
          end
          else
          begin
            mShift := MapVirtualKey(VK_SHIFT, 0);
            Include(mKeyState, ssShift);
            //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
            if ord(mKeyChar) in [65..90] then //A-Z
              mKeyChar := AnsiChar(ord(mKeyChar) + 32)
            else
            if ord(mKeyChar) in [97..122] then //a-z
              mKeyChar := AnsiChar(ord(mKeyChar) - 32);
            //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
          end;
        end
        else
          mShift := 0;
        //20140407 mark by 01753 �n�ϵ��ѥ����ӻ��A�ڡC������n���o�q

        //�O�_�w���U CONTROL
        //if CheckKeyState(VK_CONTROL) then //20121017 mark by 3188
        if CheckKeyState(VK_CONTROL, KBS) then
        begin
          mCtrl := MapVirtualKey(VK_CONTROL, 0);
          Include(mKeyState, ssCtrl);
        end
        else
          mCtrl := 0;

        //�O�_�w���U ALT
        //if CheckKeyState(VK_MENU) then  //20121017 mark by 3188
        if CheckKeyState(VK_MENU, KBS) then
        begin
          mAlt := MapVirtualKey(VK_MENU, 0);
          Include(mKeyState, ssAlt);
        end
        else
          mAlt := 0;

        ResetKeyState;
        if Assigned(KeyHook.OnHookKeyUpEvent) then
          KeyHook.OnHookKeyUpEvent(KeyHook, mScanCode + mShift + mCtrl + mAlt, mKeyChar, mKeyState);
      end
      else
      if xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU] then
      begin
        SetKeyState(xWParam);
      end;

      case xWParam of
          VK_SHIFT, VK_RETURN, VK_TAB, VK_BACK, VK_CLEAR, VK_SPACE, VK_ESCAPE,
          VK_PRIOR, VK_NEXT,   VK_LEFT,VK_RIGHT,VK_UP,    VK_DOWN,  VK_CANCEL,
          VK_ACCEPT:
            Result := CallNextHookEx(KeyHook.FHook, xCode, xWParam, xLParam);
      end;
    end else
    begin
      Result := CallNextHookEx(KeyHook.FHookWin, xCode, xWParam, xLParam);
    end;
  end;
end;
//mark by 4084 20131203
}

{
function KeyboardHookProc(xCode: Integer; xWParam: WPARAM; xLParam: LPARAM): LRESULT; stdcall;
var
  mScanCode: Longint;
  mShift, mAlt, mCtrl, mCapital: DWORD;
  mKeyState: TShiftState;
  mKeyChar: Ansichar;
  mVKeyChar : array[1..2] of AnsiChar;
  KBS : TKeyboardState;
  procedure ResetKeyState;
  begin
    KeyHook.FShiftKey := false;
    KeyHook.FCtrlKey := false;
    KeyHook.FAltKey := false;
  end;
begin
  Result:=0;
  if ((xLParam and $80000000) = $80000000) then
  begin
    if not(xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU]) then
    begin
      mKeyState := [];
      fillchar(mVKeyChar, SizeOf(mVKeyChar), #0);
      GetKeyboardState(KBS);
      if ToAsciiEx(xWParam, (xLParam and $00ff0000), KBS, @mVKeyChar, 0, GetKeyboardLayout(0)) <> 0 then
        mKeyChar := mVKeyChar[1]
      else
        mKeyChar := #0;
      mScanCode := ((xLParam and $00ff0000) div $100) + (xWParam and $ff);

      //�O�_�w���U SHIFT
      //if CheckKeyState(VK_SHIFT) then //20121017 mark by 3188
      if CheckKeyState(VK_SHIFT, KBS) then
      begin
        if ord(mKeyChar) in [48..57] then //�Ʀr��+SHIFT�ܧ󬰲Ÿ�
        begin
          case mKeyChar of
            '0':  mKeyChar := ')';
            '1':  mKeyChar := '!';
            '2':  mKeyChar := '@';
            '3':  mKeyChar := '#';
            '4':  mKeyChar := '$';
            '5':  mKeyChar := '%';
            '6':  mKeyChar := '^';
            '7':  mKeyChar := '&';
            '8':  mKeyChar := '*';
            '9':  mKeyChar := '(';
          end;
          mShift := 0;
        end
        else
        begin
          mShift := MapVirtualKey(VK_SHIFT, 0);
          Include(mKeyState, ssShift);
          //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
          if ord(mKeyChar) in [65..90] then //A-Z
            mKeyChar := AnsiChar(ord(mKeyChar) + 32)
          else
          if ord(mKeyChar) in [97..122] then //a-z
            mKeyChar := AnsiChar(ord(mKeyChar) - 32);
          //20110518 �W�[�YPreSysKeyUp,�h�۰ʱN�r���ഫ�j/�p�g��
        end;
      end
      else
        mShift := 0;

      //�O�_�w���U CONTROL
      //if CheckKeyState(VK_CONTROL) then //20121017 mark by 3188
      if CheckKeyState(VK_CONTROL, KBS) then
      begin
        mCtrl := MapVirtualKey(VK_CONTROL, 0);
        Include(mKeyState, ssCtrl);
      end
      else
        mCtrl := 0;

      //�O�_�w���U ALT
      //if CheckKeyState(VK_MENU) then  //20121017 mark by 3188
      if CheckKeyState(VK_MENU, KBS) then
      begin
        mAlt := MapVirtualKey(VK_MENU, 0);
        Include(mKeyState, ssAlt);
      end
      else
        mAlt := 0;

      ResetKeyState;
      if Assigned(KeyHook.OnHookKeyUpEvent) then
        KeyHook.OnHookKeyUpEvent(KeyHook, mScanCode + mShift + mCtrl + mAlt, mKeyChar, mKeyState);

    end
    else
    if xWParam in [VK_SHIFT, VK_CONTROL, VK_MENU] then
    begin
          SetKeyState(xWParam);
    end;
  end;
end;
 }

constructor TKeyHook.Create(AOwner: TComponent);
begin
    if not(csDesigning in ComponentState) then
        KeyHook := self;
    FPreSysKeyUp := false;
    FSendkey := False;   //add
    FShiftKey := false;
    FCtrlKey := false;
    FAltKey := false;
    FShiftState := False;
    inherited Create(AOwner);
end;

destructor TKeyHook.Destroy;
begin
    UnHookWindowsHookEx(FHook);
    UnHookWindowsHookEx(FHookWin);
    FHook := 0;
    FHookWin := 0;
    inherited Destroy;
end;

procedure TKeyHook.SetEnabled(Value: boolean);
begin
    if not(csDesigning in ComponentState) then
    begin
        if Value and not FEnabled then
        begin
            FHook    := SetWindowsHookEx(WH_KEYBOARD, @KeyboardHookProc, HInstance, GetCurrentThreadId);
            FHookWin := SetWindowsHookEx(WH_KEYBOARD_LL, @KeyboardHookProcWinKey, HInstance, 0);
        end
        else
        if not Value then
        begin
            ResetKeyoardState;
            UnHookWindowsHookEx(FHook);
            UnHookWindowsHookEx(FHookWin);
            FHook := 0;
            FHookWin := 0;
        end;
        FEnabled := (FHook <> 0);
    end;
end;

procedure TKeyHook.ResetKeyoardState;
var
  KBS: TKeyboardState;
begin
  fillchar(KBS, SizeOf(KBS), #0);
  SetKeyboardState(KBS);
end;

end.





