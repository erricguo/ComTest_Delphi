object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'COMTest'
  ClientHeight = 479
  ClientWidth = 458
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -16
  Font.Name = #24494#36575#27491#40657#39636
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 20
  object Label5: TLabel
    Left = 116
    Top = 28
    Width = 80
    Height = 20
    Alignment = taRightJustify
    AutoSize = False
    Caption = #36899#25509#26041#24335
  end
  object Label6: TLabel
    Left = 116
    Top = 58
    Width = 80
    Height = 20
    Alignment = taRightJustify
    AutoSize = False
    Caption = #30332#31080#27231#22411
  end
  object Label7: TLabel
    Left = 116
    Top = 87
    Width = 80
    Height = 20
    Alignment = taRightJustify
    AutoSize = False
    Caption = 'UniName'
  end
  object ckInvPrinterStatus: TPOSCheckBox
    Left = 9
    Top = 4
    Width = 144
    Height = 17
    Caption = #20597#28204#21360#34920#27231#29376#24907
    Checked = True
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 0
  end
  object cdInvPrinterDefaultReset: TPOSCheckBox
    Left = 161
    Top = 4
    Width = 183
    Height = 17
    Caption = #20659#36865'Default Reset'#25351#20196
    Checked = True
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 1
  end
  object Button1: TPOSButton
    Left = 8
    Top = 27
    Width = 100
    Height = 25
    CanFocus = True
    Caption = #36899'    '#25509
    Color = clSilver
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    HotTrackColor = clYellow
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -16
    HotTrackFont.Name = #32048#26126#39636
    HotTrackFont.Style = []
    ModalResult = 0
    ParentColor = False
    ParentFont = False
    Style = bsModern
    TabOrder = 2
    UseHotTrackFont = False
    OnClick = Button1Click
  end
  object Button2: TPOSButton
    Left = 8
    Top = 51
    Width = 100
    Height = 25
    CanFocus = True
    Caption = #20013#26039#36899#25509
    Color = clSilver
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    HotTrackColor = clYellow
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -16
    HotTrackFont.Name = #32048#26126#39636
    HotTrackFont.Style = []
    ModalResult = 0
    ParentColor = False
    ParentFont = False
    Style = bsModern
    TabOrder = 3
    UseHotTrackFont = False
  end
  object btFF: TPOSButton
    Left = 8
    Top = 75
    Width = 100
    Height = 25
    CanFocus = True
    Caption = #36339'    '#38913
    Color = clSilver
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    HotTrackColor = clYellow
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -16
    HotTrackFont.Name = #32048#26126#39636
    HotTrackFont.Style = []
    ModalResult = 0
    ParentColor = False
    ParentFont = False
    Style = bsModern
    TabOrder = 4
    UseHotTrackFont = False
  end
  object btLF: TPOSButton
    Left = 8
    Top = 99
    Width = 100
    Height = 25
    CanFocus = True
    Caption = #36339'    '#34892
    Color = clSilver
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    HotTrackColor = clYellow
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -16
    HotTrackFont.Name = #32048#26126#39636
    HotTrackFont.Style = []
    ModalResult = 0
    ParentColor = False
    ParentFont = False
    Style = bsModern
    TabOrder = 5
    UseHotTrackFont = False
  end
  object ComboBox2: TComboBox
    Left = 215
    Top = 24
    Width = 129
    Height = 28
    Hint = #36899#25509#26041#24335
    Style = csDropDownList
    ItemIndex = 1
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Text = 'ct_SerialPort'
    Items.Strings = (
      'ct_ParallelPort'
      'ct_SerialPort'
      'ct_PrintServer')
  end
  object ComboBox1: TComboBox
    Left = 346
    Top = 24
    Width = 103
    Height = 28
    Hint = #36899#25509#22496
    Style = csDropDownList
    ItemIndex = 0
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Text = 'cp_COM1 '
    Items.Strings = (
      'cp_COM1 '
      'cp_COM2'
      'cp_COM3'
      'cp_COM4'
      'cp_COM5'
      'cp_COM6'
      'lp_LPT1'
      'lp_LPT2'
      'lp_LPT3'
      'lp_LPT4'
      'lp_LPT5'
      'lp_LPT6'
      'cp_COM7'
      'cp_COM8'
      'cp_COM9')
  end
  object ComboBox3: TComboBox
    Left = 215
    Top = 54
    Width = 234
    Height = 28
    Hint = #30332#31080#27231#22411
    Style = csDropDownList
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
  end
  object edUniName: TEdit
    Left = 215
    Top = 85
    Width = 234
    Height = 24
    Hint = #32178#36335#30332#31080#27231#27231#21517#31281
    Color = clInfoBk
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
  end
  object btPrint: TPOSButton
    Left = 8
    Top = 154
    Width = 100
    Height = 25
    CanFocus = True
    Caption = #21015#21360#28204#35430
    Color = clSilver
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    HotTrackColor = clYellow
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -16
    HotTrackFont.Name = #32048#26126#39636
    HotTrackFont.Style = []
    ModalResult = 0
    ParentColor = False
    ParentFont = False
    Style = bsModern
    TabOrder = 10
    UseHotTrackFont = False
  end
  object btCutPaper: TPOSButton
    Left = 114
    Top = 154
    Width = 100
    Height = 25
    CanFocus = True
    Caption = #20999'    '#32025
    Color = clSilver
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #32048#26126#39636
    Font.Style = []
    HotTrackColor = clYellow
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -16
    HotTrackFont.Name = #32048#26126#39636
    HotTrackFont.Style = []
    ModalResult = 0
    ParentColor = False
    ParentFont = False
    Style = bsModern
    TabOrder = 11
    UseHotTrackFont = False
  end
  object Memo2: TMemo
    Left = 8
    Top = 185
    Width = 441
    Height = 259
    Lines.Strings = (
      'Memo2')
    ScrollBars = ssBoth
    TabOrder = 12
    WordWrap = False
  end
  object StatusBar: TStatusBar
    Left = 0
    Top = 450
    Width = 458
    Height = 29
    Panels = <
      item
        Width = 240
      end
      item
        Width = 50
      end>
    ExplicitLeft = -183
    ExplicitTop = 593
    ExplicitWidth = 641
  end
  object ckDSR: TPOSCheckBox
    Left = 350
    Top = 4
    Width = 100
    Height = 17
    Caption = 'DSR'#27298#26597
    Checked = True
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 14
  end
  object InvPrinter: TInvPrinter
    CheckDSR = False
    Port = cp_COM1
    WTimeOut = 300
    Delay = 0
    BaudRate = brt_9600
    DataBits = dbt_8
    StopBits = sbt_OneStopBit
    ParityCheck = pct_NoParity
    InbondBuffer = 2048
    OutbondBuffer = 2048
    PortHandle = 0
    DefaultReset = True
    CheckLPTStatus = False
    ErrorMsg = #30332#31080#27231#26377#21839#38988'!'#35531#27298#26597'!!'
    Active = False
    Direction = pd_Both
    PrinterType = TP_3688
    OutPutMode = om_Printer
    ConnectType = cm_SerialPort
    PrinterIndex = 2
    Left = 384
    Top = 136
  end
end
