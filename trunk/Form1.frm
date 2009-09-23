VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Visual Ansi 2008 alpha"
   ClientHeight    =   8325
   ClientLeft      =   4800
   ClientTop       =   240
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10800
   Tag             =   "0"
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   7560
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame3 
      Caption         =   "µeµ§"
      Height          =   3885
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   3030
      Begin VB.CommandButton Command5 
         Caption         =   "½s¿è"
         Height          =   300
         Left            =   1905
         TabIndex        =   46
         Top             =   1080
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "²Ó©úÅé"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "¢i"
         Top             =   1050
         Width           =   570
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2445
         Left            =   90
         TabIndex        =   4
         Top             =   1395
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   5
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame12"
         Height          =   1125
         Left            =   135
         TabIndex        =   78
         Top             =   225
         Width           =   2400
         Begin VB.OptionButton Option1 
            Caption         =   "¦r¤¸ªí"
            Height          =   210
            Left            =   0
            TabIndex        =   85
            Top             =   885
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "¦Û­q¤å¥y"
            Height          =   180
            Left            =   0
            TabIndex        =   84
            Top             =   315
            Width           =   1125
         End
         Begin VB.OptionButton Option3 
            Caption         =   "¦Û­q¦r¤¸"
            Height          =   180
            Left            =   0
            TabIndex        =   83
            Top             =   570
            Width           =   1140
         End
         Begin VB.TextBox Text6 
            Height          =   240
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   82
            Text            =   "§s"
            Top             =   540
            Width           =   375
         End
         Begin VB.CommandButton Command16 
            Caption         =   "¤º®e"
            Height          =   240
            Left            =   1200
            TabIndex        =   81
            Top             =   270
            Width           =   720
         End
         Begin VB.CheckBox Check7 
            Caption         =   "¨ú¥N"
            Height          =   195
            Left            =   1200
            TabIndex        =   80
            Top             =   15
            Width           =   705
         End
         Begin VB.CheckBox Check6 
            Caption         =   "³sÄò"
            Height          =   180
            Left            =   0
            TabIndex        =   79
            Top             =   15
            Value           =   1  '®Ö¨ú
            Width           =   675
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   645
            Picture         =   "Form1.frx":030A
            Top             =   -15
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   1920
            Picture         =   "Form1.frx":0694
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame13"
         Height          =   1215
         Left            =   120
         TabIndex        =   86
         Top             =   150
         Width           =   2760
         Begin VB.CheckBox Check10 
            Caption         =   "°»´úÂù¦ì¤¸¦r"
            Height          =   180
            Left            =   1230
            TabIndex        =   91
            Top             =   75
            Value           =   1  '®Ö¨ú
            Width           =   1425
         End
         Begin VB.CheckBox Check9 
            Caption         =   "ÂÐ»\"
            Height          =   180
            Left            =   45
            TabIndex        =   90
            Top             =   315
            Width           =   900
         End
         Begin VB.CheckBox Check8 
            Caption         =   "´å¼Ð²¾°Ê"
            Height          =   180
            Left            =   45
            TabIndex        =   89
            Top             =   75
            Value           =   1  '®Ö¨ú
            Width           =   1065
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   825
            MultiLine       =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   555
            Width           =   780
         End
         Begin VB.Label Label14 
            Caption         =   "Áä½LÂ^¨ú"
            Height          =   210
            Left            =   45
            TabIndex        =   88
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label13 
            Caption         =   "ÂI¨ú¦r¤¸ªí"
            Height          =   195
            Left            =   75
            TabIndex        =   87
            Top             =   990
            Width           =   1005
         End
      End
      Begin VB.Label Label4 
         Height          =   285
         Left            =   1605
         TabIndex        =   20
         Top             =   255
         Width           =   330
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "µe¥¬±±ªO"
      Height          =   1410
      Left            =   735
      TabIndex        =   7
      Top             =   4395
      Width           =   6570
      Begin VB.CommandButton Command21 
         Caption         =   "³]©w"
         Height          =   270
         Left            =   3960
         TabIndex        =   101
         Top             =   1050
         Width           =   645
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Appearance      =   0  '¥­­±
         Height          =   240
         Left            =   3120
         TabIndex        =   100
         Text            =   "1"
         Top             =   1080
         Width           =   390
      End
      Begin VB.CheckBox Check11 
         Caption         =   "debug"
         Height          =   195
         Left            =   5160
         TabIndex        =   93
         Top             =   720
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton Command14 
         Caption         =   "¼·©ñ"
         Height          =   285
         Left            =   4800
         TabIndex        =   52
         Top             =   1050
         Width           =   930
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  '¥­­±
         Caption         =   "³zµø"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   51
         ToolTipText     =   "¦¹¥\¯à¥¼§¹¦¨"
         Top             =   300
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton Command12 
         Caption         =   "GO!"
         Height          =   300
         Left            =   5925
         TabIndex        =   50
         Top             =   270
         Width           =   525
      End
      Begin VB.CommandButton Command11 
         Caption         =   "¤Û¼v"
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   1050
         Width           =   1005
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Appearance      =   0  '¥­­±
         Height          =   240
         Left            =   1800
         TabIndex        =   47
         Text            =   "1"
         Top             =   1065
         Width           =   450
      End
      Begin VB.CommandButton Command10 
         Caption         =   "·s¼W¤U¤@­¶"
         Height          =   300
         Left            =   2400
         TabIndex        =   45
         Top             =   705
         Width           =   1170
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Appearance      =   0  '¥­­±
         Height          =   240
         Left            =   1800
         TabIndex        =   43
         Text            =   "1"
         Top             =   705
         Width           =   450
      End
      Begin VB.CommandButton Command9 
         Caption         =   "¾ã­¶½Æ»s"
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   675
         Width           =   1020
      End
      Begin VB.CommandButton Command8 
         Caption         =   "²M°£µe¥¬"
         Height          =   300
         Left            =   285
         TabIndex        =   41
         Top             =   1500
         Width           =   1410
      End
      Begin VB.CommandButton Command7 
         Caption         =   "¤U¤@­¶"
         Height          =   285
         Index           =   1
         Left            =   3030
         TabIndex        =   39
         Top             =   300
         Width           =   825
      End
      Begin VB.CommandButton Command7 
         Caption         =   "¤W¤@­¶"
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   38
         Top             =   300
         Width           =   765
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '¥­­±
         Height          =   300
         Left            =   4905
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   37
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4890
         TabIndex        =   17
         Text            =   "14"
         Top             =   1605
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4035
         TabIndex        =   16
         Text            =   "28"
         Top             =   1635
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  '¥­­±
         Caption         =   "§ó·s"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1020
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear  Reset Size"
         Height          =   300
         Left            =   2190
         TabIndex        =   8
         Top             =   1695
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label15 
         Caption         =   "°±¯d®É¶¡          ¬í"
         Height          =   210
         Left            =   2400
         TabIndex        =   99
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "¨Ó·½"
         Height          =   240
         Left            =   1200
         TabIndex        =   49
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "¨Ó·½"
         Height          =   240
         Left            =   1200
         TabIndex        =   44
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "§Ö³t¸õ­¶:"
         Height          =   210
         Left            =   4080
         TabIndex        =   40
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "H:"
         Height          =   225
         Index           =   1
         Left            =   4620
         TabIndex        =   19
         Top             =   1650
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "W:"
         Height          =   225
         Index           =   0
         Left            =   3780
         TabIndex        =   18
         Top             =   1650
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "²¾°Ê"
      Height          =   3795
      Left            =   4800
      TabIndex        =   104
      Top             =   600
      Visible         =   0   'False
      Width           =   2940
      Begin VB.CheckBox Check12 
         Caption         =   "Ãä¬É­­¨î"
         Height          =   210
         Left            =   1605
         TabIndex        =   110
         Top             =   375
         Width           =   1080
      End
      Begin VB.CommandButton Command1 
         Caption         =   "¡ô"
         Height          =   330
         Index           =   0
         Left            =   705
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   109
         Top             =   300
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "¡õ"
         Height          =   330
         Index           =   1
         Left            =   705
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   108
         Top             =   1020
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "¡ö"
         Height          =   330
         Index           =   2
         Left            =   360
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   107
         Top             =   690
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Caption         =   "¡÷"
         Height          =   330
         Index           =   3
         Left            =   1065
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   106
         Top             =   645
         Width           =   360
      End
      Begin VB.Label Label16 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         BorderStyle     =   1  '³æ½u©T©w
         Caption         =   "²¾°Ê¿ï¨ú"
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   705
         TabIndex        =   105
         Top             =   645
         Width           =   390
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -330
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1156
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":188E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3FB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame11 
      Caption         =   "½Æ»s¡®¶K¤W"
      Height          =   3900
      Left            =   8280
      TabIndex        =   72
      Top             =   840
      Visible         =   0   'False
      Width           =   3000
      Begin VB.CommandButton Command23 
         Caption         =   "½Æ»s¨ì°Å¶KÃ¯"
         Height          =   300
         Left            =   345
         TabIndex        =   103
         Top             =   1395
         Width           =   1380
      End
      Begin VB.CommandButton Command19 
         Caption         =   "±q°Å¶KÃ¯Â^¨ú"
         Height          =   300
         Left            =   345
         TabIndex        =   92
         Top             =   1845
         Width           =   1380
      End
      Begin VB.CheckBox Check5 
         Caption         =   "¥h­I"
         Height          =   270
         Left            =   375
         TabIndex        =   76
         Top             =   2730
         Width           =   705
      End
      Begin VB.CommandButton Command18 
         Caption         =   "¶K¤W¼È¦s"
         Height          =   300
         Index           =   2
         Left            =   330
         TabIndex        =   75
         Top             =   2280
         Width           =   1380
      End
      Begin VB.CommandButton Command18 
         Caption         =   "°Å¤U¨ì¼È¦s"
         Height          =   300
         Index           =   1
         Left            =   345
         TabIndex        =   74
         Top             =   975
         Width           =   1380
      End
      Begin VB.CommandButton Command18 
         Caption         =   "½Æ»s¨ì¼È¦s°Ï"
         Height          =   315
         Index           =   0
         Left            =   345
         TabIndex        =   73
         Top             =   510
         Width           =   1380
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "¹Ï¤ù"
      Height          =   3795
      Left            =   6300
      TabIndex        =   94
      Top             =   2910
      Visible         =   0   'False
      Width           =   2940
      Begin VB.CommandButton Command20 
         Caption         =   "¸ü¤J¹Ï¤ù"
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   98
         Top             =   360
         Width           =   1530
      End
      Begin VB.CommandButton Command20 
         Caption         =   "ÀËµø/½Õ¾ã¹Ï¤ù"
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   97
         Top             =   855
         Width           =   1530
      End
      Begin VB.CommandButton Command20 
         Caption         =   "³]¬°­I´º"
         Height          =   330
         Index           =   2
         Left            =   210
         TabIndex        =   96
         Top             =   1365
         Width           =   1530
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Âà´«¹Ï¤ù¨ìµe¥¬"
         Height          =   330
         Index           =   3
         Left            =   210
         TabIndex        =   95
         Top             =   1860
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "§R°£"
      Height          =   3810
      Left            =   6360
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   2955
      Begin VB.CommandButton Command17 
         Caption         =   "²M°£"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   1425
         TabIndex        =   64
         Top             =   600
         Width           =   840
      End
      Begin VB.OptionButton Option4 
         Caption         =   "¿ï¨ú°Ï¶ô"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   63
         Top             =   645
         Width           =   1140
      End
      Begin VB.OptionButton Option4 
         Caption         =   "ÂI¿ï"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   62
         Top             =   360
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton Command15 
         Caption         =   "¾ã­¶²M°£"
         Height          =   315
         Left            =   270
         TabIndex        =   60
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label Label6 
         Height          =   1380
         Left            =   525
         TabIndex        =   27
         Top             =   1470
         Width           =   1830
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "´«¦â"
      Height          =   3810
      Left            =   6240
      TabIndex        =   53
      Top             =   2760
      Visible         =   0   'False
      Width           =   2955
      Begin VB.CommandButton Command22 
         Caption         =   "¦P­I´º"
         Height          =   300
         Left            =   2130
         TabIndex        =   102
         Top             =   1515
         Width           =   690
      End
      Begin VB.CommandButton Command17 
         Caption         =   "¹ê¦æ"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   71
         Top             =   795
         Width           =   840
      End
      Begin VB.OptionButton Option6 
         Caption         =   "¿ï¨ú°Ï¶ô"
         Height          =   270
         Index           =   1
         Left            =   300
         TabIndex        =   70
         Top             =   795
         Width           =   1140
      End
      Begin VB.OptionButton Option6 
         Caption         =   "ÂI¿ï"
         Height          =   270
         Index           =   0
         Left            =   300
         TabIndex        =   69
         Top             =   405
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.Label Lb_ExCh_Color 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   3
         Left            =   1605
         TabIndex        =   59
         Top             =   2295
         Width           =   435
      End
      Begin VB.Label Lb_ExCh_Color 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   2
         Left            =   840
         TabIndex        =   58
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Lb_ExCh_Color 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   1
         Left            =   1620
         TabIndex        =   57
         Top             =   1470
         Width           =   435
      End
      Begin VB.Label Lb_ExCh_Color 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '³æ½u©T©w
         Height          =   405
         Index           =   0
         Left            =   810
         TabIndex        =   56
         Top             =   1485
         Width           =   435
      End
      Begin VB.Label Label12 
         Caption         =   "«e´º               -->"
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   54
         Top             =   1575
         Width           =   2010
      End
      Begin VB.Label Label12 
         Caption         =   "­I´º               -->"
         Height          =   240
         Index           =   1
         Left            =   330
         TabIndex        =   55
         Top             =   2385
         Width           =   2010
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "¤W¦â"
      Height          =   3855
      Left            =   6840
      TabIndex        =   65
      Top             =   480
      Visible         =   0   'False
      Width           =   2925
      Begin VB.CommandButton Command17 
         Caption         =   "¹ê¦æ"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   1560
         TabIndex        =   68
         Top             =   780
         Width           =   840
      End
      Begin VB.OptionButton Option5 
         Caption         =   "¿ï¨ú°Ï¶ô"
         Height          =   210
         Index           =   1
         Left            =   465
         TabIndex        =   67
         Top             =   825
         Width           =   1140
      End
      Begin VB.OptionButton Option5 
         Caption         =   "ÂI¿ï"
         Height          =   210
         Index           =   0
         Left            =   450
         TabIndex        =   66
         Top             =   450
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Timer Timer1 
      Left            =   90
      Top             =   5550
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   -300
      Top             =   -180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4670
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "BBS"
      Height          =   1875
      Left            =   570
      TabIndex        =   14
      Top             =   6285
      Visible         =   0   'False
      Width           =   6675
      Begin VB.TextBox Text3 
         Appearance      =   0  '¥­­±
         Height          =   960
         Left            =   330
         MultiLine       =   -1  'True
         ScrollBars      =   3  '¨âªÌ¬Ò¦³
         TabIndex        =   15
         Top             =   660
         Width           =   5910
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "µe¥¬"
      Height          =   3465
      Left            =   720
      TabIndex        =   11
      Top             =   915
      Width           =   3300
      Begin VB.PictureBox Pic1 
         Appearance      =   0  '¥­­±
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "²Ó©úÅé"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   2820
         Left            =   75
         ScaleHeight     =   186
         ScaleMode       =   3  '¹³¯À
         ScaleWidth      =   202
         TabIndex        =   12
         Top             =   435
         Width           =   3060
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF00FF&
            BorderStyle     =   3  'ÂI½u
            DrawMode        =   4  'Mask Not Pen
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.Label Label2 
         Height          =   210
         Left            =   210
         TabIndex        =   13
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "½Õ¦â½L"
      Height          =   5445
      Left            =   45
      TabIndex        =   2
      Top             =   60
      Width           =   675
      Begin VB.PictureBox Pic3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawWidth       =   5
         FillStyle       =   4  '¥ª¤W¨ì¥k¤Uªº±×½u
         Height          =   1155
         Left            =   75
         MouseIcon       =   "Form1.frx":4A0C
         MousePointer    =   2  '¤Q¦r§Îª¬
         ScaleHeight     =   4
         ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
         ScaleWidth      =   1.6
         TabIndex        =   21
         Top             =   4140
         Width           =   510
      End
      Begin VB.PictureBox Pic2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawWidth       =   5
         FillStyle       =   4  '¥ª¤W¨ì¥k¤Uªº±×½u
         Height          =   2250
         Left            =   60
         MouseIcon       =   "Form1.frx":56D6
         MousePointer    =   2  '¤Q¦r§Îª¬
         ScaleHeight     =   8
         ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
         ScaleWidth      =   1.6
         TabIndex        =   6
         Top             =   1185
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "­I´º"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   23
         Top             =   3885
         Width           =   510
      End
      Begin VB.Label Label5 
         Caption         =   "«e´º"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   930
         Width           =   615
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -270
         TabIndex        =   10
         Top             =   3480
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '¤£³z©ú
         Height          =   360
         Left            =   45
         Top             =   270
         Width           =   390
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  '¤£³z©ú
         Height          =   375
         Left            =   225
         Top             =   390
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "¤u¨ã"
      Height          =   750
      Left            =   795
      TabIndex        =   1
      Top             =   75
      Width           =   3945
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   390
         Left            =   720
         TabIndex        =   77
         Top             =   240
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "µeµ§"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "¾ó¥ÖÀ¿"
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "¤W¦â¨ê¤l"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "ª«¥ó¦L³¹"
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "ÃC¦âÅÜ´«"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "½Æ»s&¶K¤W"
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "¿é¤J¼Ò¦¡"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "­I´º¹Ï¤ù"
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "²¾°Ê"
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox Check2 
         Caption         =   "­I´º"
         Height          =   180
         Left            =   30
         TabIndex        =   25
         Top             =   480
         Value           =   1  '®Ö¨ú
         Width           =   810
      End
      Begin VB.CheckBox Check1 
         Caption         =   "«e´º"
         Height          =   180
         Left            =   45
         TabIndex        =   24
         Top             =   210
         Value           =   1  '®Ö¨ú
         Width           =   915
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "ª«¥ó"
      Height          =   3870
      Left            =   4560
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   2925
      Begin VB.CheckBox Check4 
         Caption         =   "¥h­I"
         Height          =   255
         Left            =   1515
         TabIndex        =   61
         Top             =   285
         Width           =   780
      End
      Begin VB.CheckBox Check3 
         Caption         =   "DeBug"
         Height          =   180
         Left            =   195
         TabIndex        =   35
         Top             =   2205
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton Command6 
         Caption         =   "²¾°£"
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   34
         Top             =   1620
         Width           =   1200
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   615
         TabIndex        =   32
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "¿ï¾Ü¼Ò¦¡"
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "¶K¤W"
               ImageIndex      =   2
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command6 
         Caption         =   "±qÀÉ®×¶×¤J"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   31
         Top             =   1200
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "¥[¤Jª«¥ó>>"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   30
         Top             =   735
         Width           =   1200
      End
      Begin VB.ListBox List1 
         Appearance      =   0  '¥­­±
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00404040&
         Height          =   2550
         Left            =   1320
         TabIndex        =   29
         Top             =   660
         Width           =   1485
      End
      Begin VB.Label Label8 
         Height          =   240
         Left            =   1335
         TabIndex        =   36
         Top             =   3450
         Width           =   1020
      End
      Begin VB.Label Label7 
         Caption         =   "¼Ò¦¡"
         Height          =   210
         Left            =   165
         TabIndex        =   33
         Top             =   330
         Width           =   450
      End
   End
   Begin VB.Menu Me_File 
      Caption         =   "ÀÉ®×(&F)"
      Begin VB.Menu Me_New 
         Caption         =   "«Ø¥ß·sÀÉ   Ctrl+N"
      End
      Begin VB.Menu Me_OpenFile 
         Caption         =   "¶}±ÒÂÂÀÉ   Ctrl+O"
      End
      Begin VB.Menu Me_Save 
         Caption         =   "Àx¦sÀÉ®×   Ctrl+S"
      End
      Begin VB.Menu Me_SaveAs 
         Caption         =   "¥t¦s·sÀÉ"
      End
      Begin VB.Menu Me_line1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Me_View 
      Caption         =   "ÀËµø(&V)"
      Begin VB.Menu Me_Display 
         Caption         =   "¼·©ñ"
      End
      Begin VB.Menu Me_Refresh 
         Caption         =   "­«·s¾ã²z"
      End
   End
   Begin VB.Menu Me_Compile 
      Caption         =   "¿é¥X(&C)"
      Begin VB.Menu Me_Compile_bbs 
         Caption         =   "BBS±m¦â½X"
      End
      Begin VB.Menu Me_Compile_html 
         Caption         =   "HTMLºô­¶"
      End
   End
   Begin VB.Menu Me_Tool 
      Caption         =   "¤u¨ã(&T)"
      Begin VB.Menu Me_ImprortText 
         Caption         =   "¶×¤J¤å³¹/±m¦â½X"
      End
      Begin VB.Menu Me_AnsiListEd 
         Caption         =   "¦r¤¸ªí½s¿è¾¹"
      End
      Begin VB.Menu Me_Director 
         Caption         =   "°Êµe¾É¼·"
      End
      Begin VB.Menu Me_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu Me_SetOptions 
         Caption         =   "¿ï¶µ³]©w"
      End
   End
   Begin VB.Menu Me_About 
      Caption         =   "Ãö©ó(&A)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newf
Dim ForColor As Byte
Dim BacColor As Byte
Public Pic1MouseDown As Boolean
Dim AnsiFile As Integer
Dim OpFile As Integer
Dim FileOpMode As Integer
Dim NowAnsi As String
Dim ToolP As Frame
Public busyDrawing As Boolean
Public SL As New SelectLine
Public AD As New ApiDrawObject
Public FileSys As New FileSystemObject
Dim preX As Integer
Dim preY As Integer
Public CCR As New ColorCodeReader
'Public CDialog1 As New CDialogClass


Private Sub Check1_Click()
    Call CC_Update(C_Fore)
    Call Setoolbar
    'Debug.Print CC(C_Fore).value
End Sub

Private Sub Check10_Click()
    Text2.SetFocus
End Sub

Private Sub Check2_Click()
    Call CC_Update(C_BG)
    Call Setoolbar
End Sub

Private Sub Check4_Click()
    CC_Update C_OBJ_deBG
End Sub

Private Sub Check6_Click()
    If Check6.Value = 1 Then
        Check7.Value = 0
        SysInfo.EdMode = 4
        
    ElseIf Check6.Value = 0 And Check7.Value = 0 Then
        SysInfo.EdMode = 1
    End If
    Debug.Print "Check6_Click was envoked"
End Sub

Private Sub Check7_Click()
    If Check7.Value = 1 Then
        Check6.Value = 0
        SysInfo.EdMode = 3
    ElseIf Check6.Value = 0 And Check7.Value = 0 Then
        SysInfo.EdMode = 1
    End If
End Sub

Private Sub Check8_Click()
    Text2.SetFocus
End Sub

Private Sub Check9_Click()
    Text2.SetFocus
End Sub
Public Sub MoveArea(Index As Integer)

End Sub


Private Sub Command1_Click(Index As Integer)
    On Error GoTo out
    Dim tmpX1 As Integer, tmpY1 As Integer, tmpX2 As Integer, tmpY2 As Integer
    tmpX1 = SL.StartPoint_X
    tmpY1 = SL.StartPoint_Y
    tmpX2 = SL.EndPoint_X
    tmpY2 = SL.EndPoint_Y
    If tmpX2 > UBound(Arrf, 1) Then tmpX1 = UBound(Arrf, 1)
    If tmpY2 > UBound(Arrf, 2) Then tmpY2 = UBound(Arrf, 2)
    Select Case Index
        Case 0
            If Check12.Value = 1 Then
                If tmpY1 = 0 Then
                    Exit Sub
                End If
            Else
                If tmpY1 = 0 Then
                    If tmpY2 = tmpY1 Then
                    
                        For i = tmpX1 To tmpX2
                            Call DoErease(i, tmpY2)
                        Next i
                        Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
                        Exit Sub
                    Else
                        For i = tmpX1 To tmpX2
                            Call DoErease(i, tmpY2)
                        Next i
                        tmpY1 = tmpY1 + 1
                    End If
                End If
            End If
            Call Eff_Move_Area(tmpX1, tmpY1, tmpX2, tmpY2, OFP.CurrentPage, 0, -1)
            Call AD.ReDraw_Area(tmpX1 - 2, tmpY1 - 1, tmpX2 + 2, tmpY2)
            SL.StartPoint_Y = tmpY1 - 1
            SL.EndPoint_Y = tmpY2 - 1
            Call SL.DrawSelect
        Case 1
            If Check12.Value = 1 Then
                If tmpY2 = UBound(Arrf, 2) Then
                    Exit Sub
                End If
            Else
                If tmpY2 = UBound(Arrf, 2) Then
                    If tmpY2 = tmpY1 Then
                    
                        For i = tmpX1 To tmpX2
                            Call DoErease(i, tmpY2)
                        Next i
                        Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
                        Exit Sub
                    Else
                        For i = tmpX1 To tmpX2
                            Call DoErease(i, tmpY2)
                        Next i
                        tmpY2 = tmpY2 - 1
                    End If
                End If
            End If
            Call Eff_Move_Area(tmpX1, tmpY1, tmpX2, tmpY2, OFP.CurrentPage, 0, 1)
            Call AD.ReDraw_Area(tmpX1 - 2, tmpY1, tmpX2 + 2, tmpY2 + 1)
            SL.StartPoint_Y = tmpY1 + 1
            SL.EndPoint_Y = tmpY2 + 1
        Case 2
            If Check12.Value = 1 Then
                If tmpX1 = 0 Then
                    Exit Sub
                ElseIf tmpX1 = 1 Then
                    For j = tmpY1 To tmpY2
                        If Arrf(tmpX1, j, OFP.CurrentPage).Ansi = -1 Then Exit Sub
                    Next j
                End If
            Else
                If tmpX1 = 0 Then
                    If tmpX2 = tmpX1 Then
                    
                        For j = tmpY1 To tmpY2
                            Call DoErease(0, j)
                        Next j
                        Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
                        Exit Sub
                    Else
                        For j = tmpY1 To tmpY2
                            Call DoErease(tmpX1, j)
                        Next j
                        tmpX1 = tmpX1 + 1
                    End If
                ElseIf tmpX1 = 1 Then
                    For j = tmpY1 To tmpY2
                    
                        If Arrf(tmpX1, j, OFP.CurrentPage).Ansi = -1 Then Call DoErease(tmpX1, j)
                    Next j
                End If
            End If
            Call Eff_Move_Area(tmpX1, tmpY1, tmpX2, tmpY2, OFP.CurrentPage, -1, 0)
            Call AD.ReDraw_Area(tmpX1 - 3, tmpY1, tmpX2 + 2, tmpY2)
            SL.StartPoint_X = tmpX1 - 1
            SL.EndPoint_X = tmpX2 - 1
            Call SL.DrawSelect
        Case 3
            If Check12.Value = 1 Then
                If tmpX2 = UBound(Arrf, 1) Then
                    Exit Sub
                ElseIf tmpX2 = UBound(Arrf, 1) - 1 Then
                    For j = tmpY1 To tmpY2
                        If Arrf(tmpX2 + 1, j, OFP.CurrentPage).Ansi = -1 Then Exit Sub
                    Next j
                End If
            Else
                If tmpX2 = UBound(Arrf, 1) Then
                    If tmpX2 = tmpX1 Then
                    
                        For j = tmpY1 To tmpY2
                            Call DoErease(tmpX2, j)
                        Next j
                        Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
                        Exit Sub
                    Else
                        For j = tmpY1 To tmpY2
                            Call DoErease(tmpX2, j)
                        Next j
                        tmpX2 = tmpX2 - 1
                    End If
                ElseIf tmpX2 = UBound(Arrf, 1) - 1 Then
                    For j = tmpY1 To tmpY2
                    
                        If Arrf(tmpX2 + 1, j, OFP.CurrentPage).Ansi = -1 Then Call DoErease(tmpX2, j)
                    Next j
                End If
            End If
            Call Eff_Move_Area(tmpX1, tmpY1, tmpX2, tmpY2, OFP.CurrentPage, 1, 0)
            Call AD.ReDraw_Area(tmpX1 - 2, tmpY1, tmpX2 + 3, tmpY2)
            SL.StartPoint_X = tmpX1 + 1
            SL.EndPoint_X = tmpX2 + 1
            Call SL.DrawSelect
    End Select
    '¼Ð°O¤wÅÜ§ó
    OFP.IsChanged = True
    Call SetFormCaption
    Exit Sub
out:
    Debug.Print "MoveClick::Err:" & Err.Description
End Sub

Private Sub Command10_Click()
    Dim oriX As Integer
    Dim oriY As Integer
    Dim oriZ As Integer
    oriX = UBound(Arrf, 1)
    oriY = UBound(Arrf, 2)
    oriZ = UBound(Arrf, 3)
    
    ReDim Preserve Arrf(oriX, oriY, 1 To oriZ + 1)
    ReDim Preserve timeLine(1 To oriZ + 1)
    timeLine(oriZ + 1) = timeLine(oriZ) '·s¼W­¶­±ªº®É¶¡»P«e¤@­¶¦P
    
    Call Set_VAA_Combo
    Call Command7_Click(1)
    Call Input_Focus    '«O«ù¿é¤JÂ^¨úªºfocus
End Sub

Private Sub Command11_Click()
    Dim sp As Integer
    Dim oriP As Integer
    sp = Fix(Val(Text8.text))
    If sp < 1 Or sp > UBound(Arrf, 3) Then
        MsgBox "          ²Ä" & sp & "­¶¨Ã¤£¦s¦b µLªk§e²{¤Û¼v                  " & vbCrLf & vbCrLf & "     ½Ð¿ï¾Ü¦s¦bªº°Êµe­¶!!!" & vbCrLf & vbCrLf & "              ½d³ò: 1~" & UBound(Arrf, 3), 16, "³Þ!§O·d¯º"
        Exit Sub
    End If
    oriP = OFP.CurrentPage
    OFP.CurrentPage = sp
    Call AD.ReDraw
    OFP.CurrentPage = oriP
    Call Input_Focus    '«O«ù¿é¤JÂ^¨úªºfocus
End Sub

Private Sub Command12_Click()
    OFP.CurrentPage = Combo1.ListIndex + 1
    'Debug.Print "OFP.CurrentPage=" & OFP.CurrentPage
    Call AD.ReDraw
    Call VAA_SetButton
End Sub



Private Sub Command14_Click()
    Call DisplayVAA
    Call Input_Focus    '«O«ù¿é¤JÂ^¨úªºfocus
End Sub

Private Sub Command15_Click()
    Call DoEreaseAll
End Sub

Private Sub Command16_Click()
    Form13.Show
End Sub

Private Sub Command17_Click(Index As Integer)
'°Ï¶ô¹ê¦æ
    Dim tmpX1 As Integer, tmpY1 As Integer, tmpX2 As Integer, tmpY2 As Integer
    Dim tmpInt As Integer
    tmpX1 = SL.StartPoint_X
    tmpY1 = SL.StartPoint_Y
    tmpX2 = SL.EndPoint_X
    tmpY2 = SL.EndPoint_Y
    If tmpX1 > tmpX2 Then
        tmpInt = tmpX1
        tmpX1 = tmpX2
        tmpX2 = tmpInt
    End If
    If tmpY1 > tmpY2 Then
        tmpInt = tmpY1
        tmpY1 = tmpY2
        tmpY2 = tmpInt
    End If
    Select Case Index
        Case 0
        '§R°£
            Call AD.resetBG_Area(tmpX1, tmpY1, tmpX2, tmpY2)
            Call DoErease_Area(tmpX1, tmpY1, tmpX2, tmpY2)
            Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
        Case 1
        '¤W¦â
            For j = tmpY1 To tmpY2
                For i = tmpX1 To tmpX2
                     Call PaintColor_A(Arrf, i, j, OFP.CurrentPage, SysInfo.ForColor, SysInfo.BacColor, CC(C_Fore).Value, CC(C_BG).Value)
                    DoEvents
                Next i
            Next j
            Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
        Case 2
        '´«¦â
            For j = tmpY1 To tmpY2
                For i = tmpX1 To tmpX2
                     Call ExChColor_Draw_A(Arrf, i, j, OFP.CurrentPage, SysInfo.ExChColor.Color(0), SysInfo.ExChColor.Color(2), SysInfo.ExChColor.Color(1), SysInfo.ExChColor.Color(3), CC(C_Fore).Value, CC(C_BG).Value)
                    DoEvents
                Next i
            Next j
            Call AD.ReDraw_Area(tmpX1, tmpY1, tmpX2, tmpY2)
    End Select
    Pic1.Refresh
End Sub

Private Sub Command18_Click(Index As Integer)
    
    X1 = SL.StartPoint_X
    Y1 = SL.StartPoint_Y
    X2 = SL.EndPoint_X
    Y2 = SL.EndPoint_Y
    If X1 > X2 Then
        tempval = X1
        X1 = X2
        X2 = tempval
    End If
    If Y1 > Y2 Then
        tempval = Y1
        Y1 = Y2
        Y2 = tempval
    End If
    Select Case Index
        Case 0
            If X2 > UBound(Arrf, 1) Then X2 = UBound(Arrf, 1)
            If Y2 > UBound(Arrf, 2) Then Y2 = UBound(Arrf, 2)
            Call CLArrayCopy(Arrf, CPArr, X1, Y1, X2, Y2, OFP.CurrentPage)
        Case 1
            If X2 > UBound(Arrf, 1) Then X2 = UBound(Arrf, 1)
            If Y2 > UBound(Arrf, 2) Then Y2 = UBound(Arrf, 2)
            Call CLArrayCopy(Arrf, CPArr, X1, Y1, X2, Y2, OFP.CurrentPage)
            Call DoErease_Area_In(X1, Y1, X2, Y2)
        Case 2
            If Command18(2).Tag = "0" Then
                Command18(2).Tag = "1"
                Command18(2).Caption = "µ²§ô¶K¤W"
                Command18(0).Enabled = False
                Command18(1).Enabled = False
            Else
                Command18(2).Tag = "0"
                Command18(2).Caption = "¶K¤W"
                Command18(0).Enabled = True
                Command18(1).Enabled = True
            End If
    End Select

End Sub

Private Sub Command19_Click()
    Call GetClipboard_ByteArray(Me.hwnd, ByteArray)

    'Dim CCR As New ColorCodeReader
    Call CCR.AnalyzeCC_ByteArray(0, 0, UBound(Arrf, 1), UBound(Arrf, 2), 1)
    'Call Me.ad.Redraw
End Sub

Private Sub Command2_Click()
    Call AD.ReDraw
    Call Input_Focus
End Sub

Public Sub Command20_Click(Index As Integer)
    On Error GoTo out

    Select Case Index
        Case 0
            CDialog1.DialogTitle = "¸ü¤J¹Ï¤ù"
            CDialog1.Filter = "¤ä´©ªº¹Ï¤ù®æ¦¡(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|©Ò¦³®æ¦¡(*.*)|*.*"
            CDialog1.FileName = ""
            CDialog1.ShowOpen
            If Dir(CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) = "" Then
                MsgBox "ÀÉ®×¤£¦s¦b", 16, "¿ù»~"
            Else
                Call AD.LoadIP(CDialog1.FileName)
                Form21.Show
                Call Form21.ReSize_UI
                Call Form21.ShowPic
            End If
        Case 1
            Form21.Show
        Case 2
            If AD.HaveBG = 1 Then
                AD.HaveBG = 0
                Command20(2).Caption = "³]¬°­I´º"
            Else
                AD.HaveBG = 1
                Command20(2).Caption = "¨ú®ø­I´º"
            End If
            Call AD.ReDraw
        'Case 2
            '¹Ï¤ùÂà´«¨ìvaµe¥¬¤W
         '   Form17.Show
        'Case 4
    End Select
    Exit Sub
out:
End Sub



Private Sub Command21_Click()
    '³]©w°±¯d®É¶¡
    Call SetTimeLine(OFP.CurrentPage, OFP.CurrentPage, Val(Text9.text))
End Sub

Private Sub Command22_Click()
    SysInfo.ExChColor.Color(0) = SysInfo.ExChColor.Color(2)
    Lb_ExCh_Color(0).BackColor = QBColor(SysInfo.ExChColor.Color(0))
    SysInfo.ExChColor.Color(1) = SysInfo.ExChColor.Color(3)
    Lb_ExCh_Color(1).BackColor = QBColor(SysInfo.ExChColor.Color(1))
End Sub

Public Sub Command23_Click()
            Dim tmpX1 As Integer, tmpX2 As Integer, tmpY1 As Integer, tmpY2 As Integer, tmpInt As Integer, ansistr As String
            tmpX1 = SL.StartPoint_X
            tmpY1 = SL.StartPoint_Y
            tmpX2 = SL.EndPoint_X
            tmpY2 = SL.EndPoint_Y
            If tmpX1 > tmpX2 Then
                tmpInt = tmpX1
                tmpX1 = tmpX2
                tmpX2 = tmpInt
            End If
            If tmpY1 > tmpY2 Then
                tmpInt = tmpY1
                tmpY1 = tmpY2
                tmpY2 = tmpInt
            End If
            Call CreatAnsiTxt_Area(ansistr, tmpX1, tmpY1, tmpX2, tmpY2)
            Call SetClipboard_ByteArray(Me.hwnd, ByteArray)
            Unload Form12
            Debug.Print "copy was called"
End Sub

Private Sub Command3_Click()

    Call SetSize(Val(Text4.text), Val(Text5.text), 1, 1)
    Pic1.Cls
    AD.SetTraget
End Sub





Private Sub Command5_Click()
    Form3.Show vbModal
End Sub

Private Sub Command6_Click(Index As Integer)
    On Error GoTo out:
    Select Case Index
        Case Is = 0
            Dim ObjName As String
            Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
            ObjName = InputBox("½Ð¿é¤Jª«¥ó¦WºÙ", " ¥[¤J·sª«¥ó", "ª«¥ó1")
            If ObjName = "" Then Exit Sub
            X1 = SL.StartPoint_X
            Y1 = SL.StartPoint_Y
            X2 = SL.EndPoint_X
            Y2 = SL.EndPoint_Y
            If X1 > X2 Then
                tempval = X1
                X1 = X2
                X2 = tempval
            End If
            If Y1 > Y2 Then
                tempval = Y1
                Y1 = Y2
                Y2 = tempval
            End If
            'Debug.Print "(" & x1 & "," & Y1 & "," & X2 & "," & Y2 & ")"
            Call CopyToObjLib(Arrf(), ObjName, X1, Y1, X2, Y2)
            Dim fileinfo As ObjListDataType
            
            fileinfo.FilePath = App.Path & "\" & ObjName & ".vaf"
            fileinfo.ObjName = ObjName
            ObjList_Add fileinfo
        Case Is = 1
            CDialog1.DialogTitle = "¿ï¾Ü¶×¤Jªºª«¥ó"
            CDialog1.Filter = "*.VAF(³æ­¶µe)|*.vaf"
            CDialog1.FileName = ""
            CDialog1.ShowOpen
            Dim fileinfo1 As ObjListDataType
            fileinfo1.FilePath = CDialog1.FileName
            fileinfo1.ObjName = InputBox("½Ð¿é¤Jª«¥ó¦WºÙ", " ¥[¤J·sª«¥ó", FileSys.GetFileName(fileinfo1.FilePath))
            ObjList_Add fileinfo1
        Case Is = 2
            ObjList_Del List1.ListIndex
    End Select
    Exit Sub
out:
    Debug.Print "Command6_Click Error Out"
End Sub

Public Sub Command7_Click(Index As Integer)
On Error GoTo out
    If OFP.filetype = 2 Or OFP.filetype = 3 Then
        Select Case Index
            Case Is = 0
                If OFP.CurrentPage <= 1 Then
                    OFP.CurrentPage = 1
                Else
                    OFP.CurrentPage = OFP.CurrentPage - 1
                End If
            Case Is = 1
                If OFP.CurrentPage >= UBound(Arrf, 3) Then
                    OFP.CurrentPage = UBound(Arrf, 3)
                    'TimerFlag = 0
                Else
                
                    OFP.CurrentPage = OFP.CurrentPage + 1
                End If
        End Select
        Call AD.ReDraw
    End If
    Call VAA_SetButton
    
    Call Input_Focus    '«O«ù¿é¤JÂ^¨úªºfocus
Exit Sub
out:
End Sub
Public Sub VAA_SetButton()
    On Error GoTo out
    If OFP.CurrentPage = 1 Then
        Command7(0).Enabled = False
    Else
        Command7(0).Enabled = True
    End If
    If OFP.CurrentPage = UBound(Arrf, 3) Then
        Command7(1).Enabled = False
        Command10.Enabled = True
    Else
        Command7(1).Enabled = True
        Command10.Enabled = False
    End If
    If OFP.CurrentPage <> 1 Then
        Text7.text = OFP.CurrentPage - 1
        Text8.text = OFP.CurrentPage - 1
    End If
    Combo1.ListIndex = OFP.CurrentPage - 1
    Text9.text = timeLine(OFP.CurrentPage)
    Exit Sub
out:
    Debug.Print "VAA_SetButton Err::" & Err.Description
End Sub

Private Sub Command8_Click()
    Dim xlen As Integer
    Dim ylen As Integer
    For i = 0 To UBound(Arrf, 2)
        For j = 0 To UBound(Arrf, 1)
            Arrf(j, i, OFP.CurrentPage).Ansi = 0
            Arrf(j, i, OFP.CurrentPage).BColor = 0
            Arrf(j, i, OFP.CurrentPage).Color = 0
        Next j
    Next i
    Pic1.Cls
    AD.SetTraget
End Sub

Private Sub Command9_Click()
    Dim sp As Integer
    sp = Fix(Val(Text7.text))
    If sp < 1 Or sp > UBound(Arrf, 3) Then
        MsgBox "          ²Ä" & sp & "­¶¨Ã¤£¦s¦b                  " & vbCrLf & vbCrLf & "     ½Ð¿ï¾Ü¦s¦bªº°Êµe­¶!!!" & vbCrLf & vbCrLf & "              ½d³ò: 1~" & UBound(Arrf, 3), 16, "³Þ!§O·d¯º"
        Exit Sub
    End If
    For i = 0 To UBound(Arrf, 2)
        For j = 0 To UBound(Arrf, 1)
            Arrf(j, i, OFP.CurrentPage) = Arrf(j, i, sp)
            DoEvents
        Next j
    Next i
    
    Call AD.ReDraw
    Call Input_Focus    '«O«ù¿é¤JÂ^¨úªºfocus
End Sub

Private Sub Form_Click()
    If SysInfo.EdMode = 10 And OFP.Closed = False Then
        Text2.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HotKey_Handler(KeyCode, Shift)
End Sub



Private Sub Form_Load()

    For i = 0 To 9
         MSFlexGrid1.ColWidth(i) = 255
    Next i
    Call LoadAnsi(App.Path & "\Ansi.txt", Me.MSFlexGrid1)
'UIªì©l¤Æ
    Set Toolbar3.ImageList = ImageList1
    
    Toolbar3.Buttons(1).Image = 1
    Toolbar3.Buttons(2).Image = 2
    Toolbar3.Buttons(3).Image = 5
    Toolbar3.Buttons(4).Image = 6
    Toolbar3.Buttons(5).Image = 7
    Toolbar3.Buttons(6).Image = 8
    Toolbar3.Buttons(7).Image = 9
    Toolbar3.Buttons(8).Image = 10
    Toolbar3.Buttons(9).Image = 11
    

'=====¹w³]­Èªº³]©w====
    SysInfo.ForColor = 7
    SysInfo.Frontsize = 14
    SysInfo.EdMode = 1 '³]©w¤u¨ã
    Set ToolP = Frame3 '³]©w¤u¨ãÄÝ©ÊÄæ
    Toolbar3.Buttons(1).Value = tbrPressed
    Toolbar2.Buttons(1).Value = tbrPressed
    '¥æ´«ÃC¦âªº³]©w
    SysInfo.ExChColor.Color(0) = 7
    SysInfo.ExChColor.Color(1) = 7
    SysInfo.ExChColor.Color(2) = 7
    SysInfo.ExChColor.Color(3) = 7
    '«e´º¨Ó·½-¦h¦æ¤å¦r
    FString.StrLen(1) = 1
'=====================
    Call GetConfic
    Call SetToolP
    '³]©w½Õ¦â½L
    Call SetColorBoard
    Call SetNowAnsi
    'SysInfo.PPA = 285
    Call ForScreenSize(SysInfo.Frontsize)
    '³]©wµe¥¬ªì©l¤j¤p
    Call SetSize(28, 14, 1, 1)
    
    'ªì©l¤Æª«¥ó¼È¦s
    ReDim ObjCA(0, 0)
    'ªì©l¤Æ°Å¶KÃ¯
    ReDim CPArr(0, 0)
    Command18(2).Tag = "0"
    '³]©wÃ¸¹Ïª«¥óªº¥Ø¼Ð
    AD.Traget = Pic1
    '³]©w°Ï¶ô¿ï¾Üª«¥ó
    SL.TragetShape = Shape3
    If SysInfo.HideSelect = 1 Then Shape3.Visible = False

    'Åª¨úª«¥ó²M³æÀÉ
    Call ObjList_Read
    'Ãö³¬ª¬ºA³]¸m
    OFP.Closed = True
    CheckClose
    Unload Form8
    Me.Show
    If Command <> "" Then
        Call OpenFile_Command(Command)
        'MsgBox "got command", 64, "information"
    End If
    '³]©wÁä½LÂ^¨ú(text2)ªº¯S§Ocall back
    'prevtext2WndProc = GetWindowLong(Text2.hwnd, GWL_WNDPROC)
    'SetWindowLong Text2.hwnd, GWL_WNDPROC, AddressOf WndProc_Text2
    '³]©wµe¥¬ªº°T®§ÄdºI
    prevWndProc_Pic = GetWindowLong(Pic1.hwnd, GWL_WNDPROC)
    SetWindowLong Pic1.hwnd, GWL_WNDPROC, AddressOf WndProc_Pic
    
    'Control Collection
    Call CC_Init
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    KillTimerEnd Me.hwnd
    If SysInfo.CheckSave = 1 And OFP.Closed = False Then
        Call AskSave
    
    End If
    Unload Form21   '¨ø°£¹Ï¤ù¤u¨ã
    Unload Form7    '¨ø°£¶×¤J¤å³¹
    Unload Form6    '¨ø°£¿é¥X
    Unload Form13   '¨ø°£¤å¥y¤º®e³]©w
    Unload Form5    '¨ø°£½sÄ¶µøµ¡
    Unload Form22   '¨ø°£html½sÄ¶µøµ¡
        '³]©wµe¥¬ªº°T®§ÄdºI
    SetWindowLong Pic1.hwnd, GWL_WNDPROC, AddressOf WndProc
    'Àx¦sª«¥ó²M³æÀÉ
    Call ObjList_Save
    Call SetConfic
End Sub



Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)
Debug.Print "Drog"
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pic1MouseDown = False
End Sub

Private Sub Lb_ExCh_Color_Click(Index As Integer)
    Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BorderStyle = 0    '±N«e¤@­Ó§ï¦¨¥¼¿ï¨ú
    SysInfo.ExChColor.CurrentSel = Index                    '³]©w­n¿ï¨úÃC¦âªº¤è¶ô
    Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BorderStyle = 1    '±N¥Ø«eªº¤è¶ô§ï¦¨¿ï¨úª¬ºA
End Sub

Private Sub List1_Click()
    On Error GoTo out
    'If SysInfo.edmode = 7 Then
        Obj_ReadFile ObjList(List1.ListIndex).FilePath, ObjCA
        'Debug.Print ObjList(List1.ListIndex).FilePath
    
    'End If
        Label8.Caption = UBound(ObjCA, 1) + 1 & "X" & UBound(ObjCA, 2) + 1
    Exit Sub
out:
    Debug.Print "List1_Click Error Out"
End Sub

Private Sub Me_About_Click()
    Form9.Show vbModal
End Sub

Private Sub Me_AnsiListEd_Click()
    Form3.Show vbModal
End Sub

Private Sub Me_Compile_bbs_Click()
    If Form5.Visible = False Then
        Form5.Show vbModal
    Else
        Form5.Show
    End If
End Sub

Private Sub Me_Compile_html_Click()
    If Form5.Visible = False Then
        Form22.Show vbModal
    Else
        Form22.Show
    End If
End Sub

Private Sub Me_Director_Click()
    Form11.Show vbModal
End Sub

Private Sub Me_Display_Click()
'¼·©ñ°Êµe
    Call DisplayVAA
End Sub

Private Sub Me_ImprortText_Click()
    Form7.Show
End Sub

Public Sub Me_New_Click()
    If OFP.Closed = False Then Call AskSave
    Call OpenNewFile
    'close ¥Ñ¶}·sÀÉ®×ªº³¡¤À§PÂ_ ¨¾¤î¨ú®øªºbug
    'OFP.Closed = False
    Call CheckClose
    Call Set_FileType_Visual
    
    Call Set_VAA_Combo
    Call VAA_SetButton
    Call AD.ReDraw
    OFP.FilePath = ""
    OFP.IsChanged = True
    Call SetFormCaption
    AD.SetTraget
    
End Sub

Public Sub Me_OpenFile_Click()
    On Error GoTo out
    If OFP.Closed = False Then Call AskSave
    FileOpMode = 1
    
    CDialog1.DialogTitle = "¶}±ÒÂÂÀÉ"
    CDialog1.InitDir = App.Path
    CDialog1.Filter = "*.VAF(³æ­¶µe) *.VAA(°Êµe) *.VAM(¦h­¶µe)|*.vaf; *.vaa; *.vam|*.VAF(³æ­¶µe)|*.vaf| *.VAA(°Êµe)| *.vaa|*.VAM(¦h­¶µe)|*.vam"
    CDialog1.FileName = ""
    'CDialog1.FileName = ""
    CDialog1.ShowOpen
    If FileSys.FileExists(CDialog1.FileName) = False Then
    'If FileSys.FileExists(CDialog1.FileName) = False Then
        MsgBox "½Ð¿ï¾Ü¦s¦bªºÀÉ®×", vbOKOnly, "ÀÉ®×¤£¦s¦b"
        Exit Sub
    End If
    'OFP.FilePath = CDialog1.FilePath
    'Debug.Print OFP.FilePath
    OFP.FilePath = CDialog1.FileName
    Call SetFormCaption("¸ü¤J¤¤...")
    VA_ReadFile OFP.FilePath, Arrf, timeLine
    OFP.Closed = False
    OFP.IsChanged = False
    Call SetFormCaption
    Call CheckClose
    Call Set_FileType_Visual
    Call Set_VAA_Combo
    Call VAA_SetButton
    Call AD.ReDraw
    'Text9.text = timeLine(1)
    Exit Sub
out:
    Debug.Print "Me_OpenFile_Click Error Out"

End Sub
Private Sub OpenFile_Command(ByVal CommandString As String)
On Error GoTo out
    If OFP.Closed = False Then Call AskSave
    FileOpMode = 1
    If FileSys.FileExists(CommandString) = False Then
        MsgBox "½Ð¿ï¾Ü¦s¦bªºÀÉ®×", vbOKOnly, "ÀÉ®×¤£¦s¦b"
        Exit Sub
    End If
    OFP.FilePath = CommandString
    VA_ReadFile OFP.FilePath, Arrf, timeLine
    OFP.Closed = False
    OFP.IsChanged = False
    Call SetFormCaption
    Call CheckClose
    Call Set_FileType_Visual
    Call Set_VAA_Combo
    Call VAA_SetButton
    Call AD.ReDraw
    
Exit Sub
out:
    Debug.Print "OpenFile_Command Error Out"

End Sub

Private Sub Me_Refresh_Click()
    Call AD.ReDraw
End Sub

Public Sub Me_Save_Click()
On Error GoTo out
    CDialog1.DialogTitle = "Àx¦sÀÉ®×"
    Select Case OFP.filetype
        Case Is = 1
            CDialog1.Filter = "*.VAF(³æ­¶µe)|*.vaf"
        Case Is = 2
            CDialog1.Filter = "*.VAA(°Êµe)|*.vaa"
        Case Is = 3
            CDialog1.Filter = "*.VAM(¦h­¶µe)|*.vam"
    End Select
    If OFP.FilePath = "" Then
        CDialog1.FileName = ""
        CDialog1.ShowSave
        OFP.FilePath = CDialog1.FileName
        
    End If
    VA_SaveFile OFP.FilePath, Arrf, OFP, timeLine
    OFP.IsChanged = False
    Call SetFormCaption
Exit Sub
out:
    Debug.Print "Me_Save Error Out"
End Sub

Private Sub Me_SaveAs_Click()
On Error GoTo out

    CDialog1.DialogTitle = "¥t¦s·sÀÉ"
    Select Case OFP.filetype
        Case Is = 1
            CDialog1.Filter = "*.VAF(³æ­¶µe)|*.vaf"
        Case Is = 2
            CDialog1.Filter = "*.VAA(°Êµe)|*.vaa"
        Case Is = 3
            CDialog1.Filter = "*.VAM(¦h­¶µe)|*.vaa"
    End Select
    CDialog1.FileName = ""
    CDialog1.ShowSave
    
    
    
    If FileSys.FileExists(CDialog1.FileName) = True Then
        If MsgBox("³o­ÓÀÉ®×¤w¸g¦s¦b,§A½T©w­nÂÐ»\¥¦¶Ü?", vbOKCancel, "ÀÉ®×¤w¦s¦b") = vbNo Then Exit Sub
    End If
    VA_SaveFile CDialog1.FileName, Arrf, OFP, timeLine

Exit Sub
out:
Debug.Print "Me_SaveAs Error Out"
End Sub

Private Sub Me_SetOptions_Click()
Form10.Show vbModal
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print MSFlexGrid1.MouseRow
'Debug.Print MSFlexGrid1.MouseCol
'Debug.Print MSFlexGrid1.RowSel
'Debug.Print MSFlexGrid1.ColSel
    If SysInfo.EdMode = 10 Then
        Text2.text = MSFlexGrid1.text
    End If
    Text1.text = MSFlexGrid1.text
    Label4.Caption = Tlen(Text1.text)
    

End Sub

Private Sub MSFlexGrid1_RowColChange()
    'If SysInfo.EdMode = 10 Then
        'Text2.text = MSFlexGrid1.text
    'End If
    'Text1.text = MSFlexGrid1.text
    'Label4.Caption = Tlen(Text1.text)

End Sub

Private Sub Option1_Click()
Call SetNowAnsi
End Sub

Private Sub Option2_Click()
    CC_Update O_Pen_Text
    Call SetNowAnsi

End Sub

Private Sub Option3_Click()
    Call SetNowAnsi
End Sub

Private Sub Option4_Click(Index As Integer)
    Command17(0).Enabled = Option4(1).Value
End Sub

Private Sub Option5_Click(Index As Integer)
    Command17(1).Enabled = Option5(1).Value
End Sub

Private Sub Option6_Click(Index As Integer)
    Command17(2).Enabled = Option6(1).Value
End Sub



Private Sub Pic1_GotFocus()
    'SL.FocusStyle (True)
    'Debug.Print "µe¥¬¨ú±oµJÂI"
End Sub

Private Sub Pic1_KeyDown(KeyCode As Integer, Shift As Integer)
    If SysInfo.EdMode = 10 Then
        Text2.SetFocus
    
    End If
    
End Sub

Private Sub Pic1_LostFocus()
    'SL.FocusStyle (False)
    'Debug.Print "µe¥¬¥¢¥hµJÂI"
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
Dim intX As Integer
Dim intY As Integer
Pic1MouseDown = True
intX = Fix(X)
intY = Fix(Y)
'If Check11.Value = 1 Then Debug.Print "(" & intX & "," & intY & ") : " & Arrf(intX, intY, OFP.CurrentPage).Ansi
Debug.Print "draw mode=" & SysInfo.EdMode
Select Case SysInfo.EdMode
    Case 1  'µeµ§
        Debug.Print "pen fore=" & CC(C_Fore).Value
        'If Check1.value = 1 Then
        If CC(C_Fore).Value = 1 Then
            Debug.Print "pen fore"
            If CC(O_Pen_Text).Value = False Then
                Call DoDraw(intX, intY, NowAnsi)
            Else
                'Call DoMutiDraw(intX, intY, FString.str, 0)
                If busyDrawing = False Then
                    busyDrawing = True
                    Call DoMutiDraw(intX, intY, FString.str, 0)
                    busyDrawing = False
                End If
                'Pic1.Refresh
            End If
        End If

        If CC(C_BG).Value = 1 Then Call DoDrawBC(intX, intY)
    Case 2  '§R°£
        If Option4(0).Value = True Then
            Call AD.resetBG(intX, intY)
            If CC(C_BG).Value = 1 Then Call DoEreaseB(intX, intY)
            If CC(C_Fore).Value = 1 Then Call DoErease(intX, intY)
            
            Call AD.ShowIt(intX, intY)
            Pic1.Refresh
        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
    Case 3  'ÂÐ¼g
        If CC(O_Pen_Text).Value <> True Then
            Dim tmpInt As Integer
            '¥ý§R°£
            If CC(C_Fore).Value = 1 Then
                If Tlen(NowAnsi) = 2 Then
                    
                    tmpInt = DoErease(intX + 1, intY)
                    Call AD.ShowIt(intX + 1, intY)
                    If tmpInt = 1 Then
                        Call AD.ShowIt(intX + 2, intY)
                    End If
                End If
                tmpInt = DoErease(intX, intY)
            End If
            If tmpInt = -1 Then Call AD.ShowIt(intX - 1, intY)
            If CC(C_BG).Value = 1 Then
                If Tlen(NowAnsi) = 2 Then
                    tmpInt = DoEreaseB(intX + 1, intY)
                    Call AD.ShowIt(intX + 1, intY)
                    If tmpInt = 1 Then
                        Call AD.ShowIt(intX + 2, intY)
                    End If
                End If
                tmpInt = DoEreaseB(intX, intY)
                
            End If
            Call AD.ShowIt(intX, intY)
            If tmpInt = -1 Then Call AD.ShowIt(intX - 1, intY)
            '¦Aµe¤W
            If CC(C_Fore).Value = 1 Then
                Call DoDraw(intX, intY, NowAnsi)
            End If
            If CC(C_BG).Value = 1 Then
                Call DoDrawBC(intX, intY)
            End If
            Call AD.ShowIt(intX, intY)
            
        Else
                If busyDrawing = False Then
                    busyDrawing = True
                    Call DoMutiDraw(intX, intY, FString.str, 1)
                    busyDrawing = False
                End If
             'Call DoMutiDraw(intX, intY, FString.str, 1)
        End If

        Pic1.Refresh
    Case 4
        
        
            If CC(O_Pen_Text).Value = False Then
                'Call DoDraw(intX, intY, NowAnsi)
                If CC(C_Fore).Value = 1 Then Call DoDraw(intX, intY, NowAnsi)
                If CC(C_BG).Value = 1 Then Call DoDrawBC(intX, intY)
            Else
                If busyDrawing = False Then
                    busyDrawing = True
                    Call DoMutiDraw(intX, intY, FString.str, 0)
                    busyDrawing = False
                End If
                'Pic1.Refresh
            End If
            
        Call SelectAnsi(intX, intY)
    Case 5
        If Option5(0).Value = True Then
            Call PaintColor(intX, intY)

        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
    Case 6
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX
        SL.EndPoint_Y = intY
        SL.DrawSelect
    Case 7
        If Check4.Value = 1 Then
        '¥h­I
            Call CLArrayPaste_C(ObjCA(), Arrf(), intX, intY, OFP.CurrentPage)

        Else
            Call ObjLibPo(ObjCA(), Arrf(), intX, intY)
        End If
        Call AD.ReDraw_Area(intX - 1, intY, UBound(ObjCA, 1) + intX + 1, UBound(ObjCA, 2) + intY)
    Case 8
        If Option6(0).Value = True Then
            Call ExChColor_Draw(intX, intY)
        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
    Case 9 '½Æ»s&¶K¤W
        If Command18(2).Tag = "0" Then
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        Else
            If Check5.Value = 0 Then
                Call ObjLibPo(CPArr(), Arrf(), intX, intY)
            Else
                '¥h­I¼Ò¦¡
                Call CLArrayPaste_C(CPArr(), Arrf(), intX, intY, OFP.CurrentPage)
            
            End If
            Call AD.ReDraw_Area(intX, intY, UBound(CPArr, 1) + intX, UBound(CPArr, 2) + intY)
        End If
    Case 10 '¿é¤J¼Ò¦¡
        Call Input_Select(intX, intY)
        Text2.SetFocus
    Case 12 '®ÄªG
        If SL.IsInSel(intX, intY) Then
            SysInfo.Move_Draging = True
            SysInfo.MD_X = intX
            SysInfo.MD_X = intY
        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
End Select
'If SysInfo.EdMode <> 7 And SysInfo.EdMode <> 6 Then
'
'    Call SelectAnsi(intX, intY)
'End If
If SysInfo.EdMode <> 6 And SysInfo.EdMode <> 10 And CC(C_Fore).Value <> 0 And CC(C_BG).Value <> 0 Then
    OFP.IsChanged = True
    Debug.Print SysInfo.EdMode
    Call SetFormCaption
End If

'Pic1.Refresh
Exit Sub
out:
Debug.Print "Pic1_MouseDown Error Out"
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
    
    Dim intX As Integer
    Dim intY As Integer
    Dim area As Byte
    intX = Fix(X)
    intY = Fix(Y)
    If intX = preX And intY = preY Then
        Exit Sub
    Else
        preX = intX
        preY = intY
    End If
    
    If Pic1MouseDown Then
        Select Case SysInfo.EdMode
            Case 2
                If Option4(0).Value = True Then
                    Call AD.resetBG(intX, intY)
                    If CC(C_BG).Value = 1 Then Call DoEreaseB(intX, intY)
                    If CC(C_Fore).Value = 1 Then Call DoErease(intX, intY)
                    
                    Call AD.ShowIt(intX, intY)
                    Pic1.Refresh
                Else
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
            Case 4
                If CC(O_Pen_Text).Value = False Then
                    'Call DoDraw(intX, intY, NowAnsi)
                    'draw call
                    Debug.Print "C(C_Fore).value=" & CC(C_Fore).Value
                    If CC(C_Fore).Value = 1 Then Call DoDraw(intX, intY, NowAnsi)
                    If CC(C_BG).Value = 1 Then Call DoDrawBC(intX, intY)
                Else
                    
                    'Call DoMutiDraw(intX, intY, FString.str, 0)
                    If busyDrawing = False Then
                        busyDrawing = True
                        Call DoMutiDraw(intX, intY, FString.str, 0)
                        busyDrawing = False
                    End If
                    'Pic1.Refresh
                End If
    
            Case 5
                If Option5(0).Value = True Then
                    Call PaintColor(intX, intY)
                Else
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
            Case 6
                SL.EndPoint_X = intX
                SL.EndPoint_Y = intY
                SL.DrawSelect
            Case 8
            
                If Option6(0).Value = True Then
                    Call ExChColor_Draw(intX, intY)
                Else
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
            Case 9
                If Command18(2).Tag = "0" Then
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
                
            Case 10
                SL.EndPoint_X = intX
                SL.EndPoint_Y = intY
                SL.DrawSelect
            Case 12
                If SysInfo.Move_Draging Then
                    '©ì¦²²¾°Ê
                    
                
                Else
                    '©ì¦²¿ï¾Ü°Ï¶ô
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
        End Select
    
        Pic1.Refresh
    Else
        'IsInSel
        Select Case SysInfo.EdMode
            Case 12 '²¾°Ê
                If SL.IsInSel(intX, intY) Then
                    If Pic1.MousePointer <> 5 Then Pic1.MousePointer = 5
                Else
                    If Pic1.MousePointer <> 0 Then Pic1.MousePointer = 0
                End If
        End Select
    End If
    If (SysInfo.EdMode = 2 And Option4(0).Value = True) Then
            Call SelectAnsi(intX, intY) '¦Û°Ê¿ï¨úÂù¦r¤¸
    End If
    If (SysInfo.EdMode = 5 And Option5(0).Value = True) Or (SysInfo.EdMode = 8 And Option6(0).Value = True) Then
    
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX
        SL.EndPoint_Y = intY
        SL.DrawSelect
    End If
    
    
    If SysInfo.EdMode = 7 Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX + UBound(ObjCA, 1)
        SL.EndPoint_Y = intY + UBound(ObjCA, 2)
        SL.DrawSelect
    End If
    If SysInfo.EdMode = 9 And Command18(2).Tag = "1" Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX + UBound(CPArr, 1)
        SL.EndPoint_Y = intY + UBound(CPArr, 2)
        SL.DrawSelect
    End If
    If SysInfo.EdMode <> 7 And SysInfo.EdMode <> 6 And SysInfo.EdMode <> 9 And SysInfo.EdMode <> 2 And SysInfo.EdMode <> 5 And SysInfo.EdMode <> 8 And SysInfo.EdMode <> 10 And SysInfo.EdMode <> 12 Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        
        If CC(O_Pen_Text).Value = True Then
            SL.EndPoint_Y = intY + FString.StrLen(1) - 1
            SL.EndPoint_X = intX + FString.StrLen(0) - 1
        Else
            SL.EndPoint_Y = intY
            SL.EndPoint_X = intX + Tlen(NowAnsi) - 1
        End If
        
        SL.DrawSelect
    End If
    If SL.StartPoint_X <> SL.EndPoint_X Or SL.StartPoint_Y <> SL.EndPoint_Y Then
        Label2.Caption = "(" & intX & "," & intY & ")" & "  [(" & SL.StartPoint_X & "," & SL.StartPoint_Y & ")-(" & SL.EndPoint_X & "," & SL.EndPoint_Y & ")] " & Get_Char(intX, intY, OFP.CurrentPage)
    Else
        Label2.Caption = "(" & intX & "," & intY & ") " & Get_Char(intX, intY, OFP.CurrentPage)
    End If
Exit Sub
out:
Debug.Print "Pic1_MouseMove Error Out"

End Sub

Public Sub lentest()
For i = 8 To 2000
    If Pic1.Point(i, 150) = 0 Then
        Debug.Print i
        Exit For
    End If

DoEvents
Next i
End Sub



Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pic1MouseDown = False
    If Check3.Value = 1 Then
        Debug.Print "Arrf(" & Fix(X) & "," & Fix(Y) & ").Ansi=" & Arrf(Fix(X), Fix(Y), OFP.CurrentPage).Ansi
        Debug.Print "Arrf(" & Fix(X) & "," & Fix(Y) & ").Color=" & Arrf(Fix(X), Fix(Y), OFP.CurrentPage).Color
        Debug.Print QBCToAnsiC(Arrf(Fix(X), Fix(Y), OFP.CurrentPage).Color)
        Debug.Print "Arrf(" & Fix(X) & "," & Fix(Y) & ").BColor=" & Arrf(Fix(X), Fix(Y), OFP.CurrentPage).BColor
        Debug.Print QBCToAnsiBC(Arrf(Fix(X), Fix(Y), OFP.CurrentPage).BColor)
    End If
End Sub



Private Sub Pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intX As Integer, intY As Integer
    intX = Fix(X)
    intY = Fix(Y)
    
    If SysInfo.EdMode = 8 Then
        Call ExChColor_SetFColor(intX, intY)
    Else
        Call GetForColor(intX, intY)
    End If

    If SysInfo.EdMode = 10 And OFP.Closed = False Then '³B²z¶K¤W¼Ò¦¡ªº¤W¦â
        
        Call Command17_Click(1)
        Text2.SetFocus
    End If

End Sub

Private Sub Pic3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intX As Integer, intY As Integer
    intX = Fix(X)
    intY = Fix(Y)
    
    If SysInfo.EdMode = 8 Then
        Call ExChColor_SetBColor(intX, intY)
    Else
        Call GetBackColor(intX, intY)
    End If
    If SysInfo.EdMode = 10 And OFP.Closed = False Then '³B²z¶K¤W¼Ò¦¡ªº¤W¦â
        
        Call Command17_Click(1)
        Text2.SetFocus
    End If

End Sub

Private Sub Text1_Change()
    Call SetNowAnsi
End Sub

Public Sub Text2_Change()
    On Error GoTo out
    If SysInfo.CopyingFlag = 1 Then ' «ö¤UCtrl+V¶K¤W
        'If Text2.text <> "" Then
            Call GetClipboard_ByteArray(Me.hwnd, ByteArray)
            
            'Dim CCR As New ColorCodeReader
            Call CCR.AnalyzeCC_ByteArray(SL.StartPoint_X, SL.StartPoint_Y, UBound(Arrf, 1), UBound(Arrf, 2))
            Call AD.ReDraw
        'End If
        SysInfo.CopyingFlag = 0
        Text2.text = ""
    Else
    
        Dim tmpStr As String
        tmpStr = Text2.text
        If tmpStr <> "" And DelayFlag = 2 Then
        
            If Check9.Value = 1 Then
                DoMutiDraw SL.StartPoint_X, SL.StartPoint_Y, Text2.text, 1
            Else
                DoMutiDraw SL.StartPoint_X, SL.StartPoint_Y, Text2.text
            End If
            
            s = GetBiAsc(Asc(Left(tmpStr, 1)), 1)
            Text2.text = ""
            Pic1.Refresh
            Text2.SetFocus
            If Check8.Value = 1 Then
                Call Input_Select(SysInfo.cDrawPos_X, SysInfo.cDrawPos_Y - 1)
            Else
                Call Input_Select(SL.StartPoint_X, SL.StartPoint_Y)
            End If
            DelayFlag = 1
        ElseIf DelayFlag = 1 Then
            DelayFlag = 0
            Call SensDelay(Me.hwnd, 100)
        End If
    End If
Exit Sub
out:
    Debug.Print "Text2_Change::Err:" & Err.Description
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    SysInfo.InputFlag = 0
    If Shift = 2 Then
        If KeyCode = 86 Then   '¶K¤W
            SysInfo.CopyingFlag = 1
        
        ElseIf KeyCode = 67 Then
            Call CopyDelay(Me.hwnd, 200)
            Exit Sub
        ElseIf KeyCode = 37 Then
            Call Command1_Click(2)
        ElseIf KeyCode = 38 Then
            Call Command1_Click(0)
        ElseIf KeyCode = 39 Then
            Call Command1_Click(3)
        ElseIf KeyCode = 40 Then
            Call Command1_Click(1)
        End If
    Else
        Select Case KeyCode
            Case 8  '¦V¥ª§R°£
                SysInfo.cDrawPos_X = SL.StartPoint_X - 1
                SysInfo.cDrawPos_Y = SL.StartPoint_Y
                Call AD.resetBG(SysInfo.cDrawPos_X, SysInfo.cDrawPos_Y)
                If CC(C_BG).Value = 1 Then Call DoEreaseB(SysInfo.cDrawPos_X, SysInfo.cDrawPos_Y)
                If CC(C_Fore).Value = 1 Then Call DoErease(SysInfo.cDrawPos_X, SysInfo.cDrawPos_Y)
                
                Call AD.ShowIt(intX, intY)
                SysInfo.InputFlag = 1
                Pic1.Refresh
            Case 46 '­ì¦ì§R°£
                Call Command17_Click(0)
            Case 36 'home
                If Shift = 1 Then
                    Call Input_Select_Area(SL.StartPoint_X, SL.StartPoint_Y, 0, SL.EndPoint_Y)
                Else
                    SysInfo.cDrawPos_X = 0
                    SysInfo.cDrawPos_Y = SL.StartPoint_Y
                    SysInfo.InputFlag = 1
                End If
            Case 35 'end
                If Shift = 1 Then
                    Call Input_Select_Area(SL.StartPoint_X, SL.StartPoint_Y, UBound(Arrf, 1), SL.EndPoint_Y)
                Else
                    SysInfo.cDrawPos_X = UBound(Arrf, 1)
                    SysInfo.cDrawPos_Y = SL.StartPoint_Y
                    SysInfo.InputFlag = 1
                End If
            'replaced by global hotkey
            'Case 33 'pageup
            '    If Command7(0).Visible = True Then
            '        Call Command7_Click(0)
            '    End If
            
            'Case 34 'pagedown
            '    If Command7(1).Visible = True Then
            '        Call Command7_Click(1)
            '    End If
            Case 13 'Enter
                SysInfo.cDrawPos_X = 0
                SysInfo.cDrawPos_Y = SL.StartPoint_Y + 1
                SysInfo.InputFlag = 1
            Case 37 '¥ª
                If Shift = 1 Then
                    Call Input_Select_Area(SL.StartPoint_X, SL.StartPoint_Y, SL.EndPoint_X - 1, SL.EndPoint_Y)
                Else
                    SysInfo.cDrawPos_X = SL.StartPoint_X - 1
                    SysInfo.cDrawPos_Y = SL.StartPoint_Y
                    SysInfo.InputFlag = 1
                End If
            Case 38 '¤W
                If Shift = 1 Then
                    Call Input_Select_Area(SL.StartPoint_X, SL.StartPoint_Y, SL.EndPoint_X, SL.EndPoint_Y - 1)
                Else
                    SysInfo.cDrawPos_X = SL.StartPoint_X
                    SysInfo.cDrawPos_Y = SL.StartPoint_Y - 1
                    SysInfo.InputFlag = 1
                End If
            Case 39 '¥k
                If Shift = 1 Then
                    Call Input_Select_Area(SL.StartPoint_X, SL.StartPoint_Y, SL.EndPoint_X + 1, SL.EndPoint_Y)
                Else
                    SysInfo.cDrawPos_X = SL.EndPoint_X + 1
                    SysInfo.cDrawPos_Y = SL.EndPoint_Y
                    SysInfo.InputFlag = 1
                End If
            Case 40 '¤U
                If Shift = 1 Then
                    Call Input_Select_Area(SL.StartPoint_X, SL.StartPoint_Y, SL.EndPoint_X, SL.EndPoint_Y + 1)
                Else
                    SysInfo.cDrawPos_X = SL.StartPoint_X
                    SysInfo.cDrawPos_Y = SL.StartPoint_Y + 1
                    SysInfo.InputFlag = 1
                End If
            Case 45
                If Shift = 1 Then
                    SysInfo.CopyingFlag = 1
                Else
                    If Check9.Value = 1 Then
                        Check9.Value = 0
                    Else
                        Check9.Value = 1
                    End If
                End If
        End Select
    End If
    If SysInfo.InputFlag = 1 Then
        Call Input_Select(SysInfo.cDrawPos_X, SysInfo.cDrawPos_Y)
    End If
    Debug.Print KeyCode & "," & Shift
End Sub
Private Sub Input_Select(ByVal X As Integer, ByVal Y As Integer)
    
        If X > UBound(Arrf, 1) Then
            X = 0
            Y = Y + 1
        ElseIf X < 0 Then
            X = 0
        End If
        If Y > UBound(Arrf, 2) Then
            Y = UBound(Arrf, 2)
        ElseIf Y < 0 Then
            Y = 0
        End If
        If Check10.Value = 1 Then
            Call SelectAnsi(X, Y)
        Else
            SL.StartPoint_X = X
            SL.StartPoint_Y = Y
            SL.EndPoint_X = X
            SL.EndPoint_Y = Y
            SL.DrawSelect
        End If
End Sub
Private Sub Input_Select_Area(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    If X1 < 0 Then
        X1 = 0
    ElseIf X1 > UBound(Arrf, 1) Then
        X1 = UBound(Arrf, 1)
    End If
    If X2 < 0 Then
        X2 = 0
    ElseIf X2 > UBound(Arrf, 1) Then
        X2 = UBound(Arrf, 1)
    End If
    If Y1 < 0 Then
        Y1 = 0
    ElseIf Y1 >= UBound(Arrf, 2) Then
        Y1 = UBound(Arrf, 2)
    End If
    If Y2 < 0 Then
        Y1 = 0
    ElseIf Y2 >= UBound(Arrf, 2) Then
        Y2 = UBound(Arrf, 2)
    End If
    SL.StartPoint_X = X1
    SL.StartPoint_Y = Y1
    SL.EndPoint_X = X2
    SL.EndPoint_Y = Y2
    SL.DrawSelect
End Sub

Private Sub Input_Focus()
    If SysInfo.EdMode = 10 And OFP.Closed = False Then '³B²z¶K¤W¼Ò¦¡ªº¤W¦â
        Text2.SetFocus
    End If

End Sub

Private Sub Text2_LostFocus()
    'If SysInfo.EdMode = 10 Then
        'Text2.SetFocus
    'End If
End Sub



Private Sub Text6_Change()
    Call SetNowAnsi
End Sub

Private Sub Text6_GotFocus()
    Text_Sel_All Text6
End Sub

Private Sub Timer1_Timer()
    If Command7(1).Enabled = False Then
        Timer1.Interval = 0
        Command14.Caption = "¼·©ñ"
        Me_Display.Caption = "¼·©ñ"
    Else
        Call Command7_Click(1)
    End If
End Sub

Public Sub TimerFunc()
    On Error Resume Next
    If TimerFlag = 0 Then
        'Timer1.Interval = 0
        'SetTimerTime Me.hwnd, 0
        Call KillTimerEnd(Me.hwnd)
        Command14.Caption = "¼·©ñ"
        Me_Display.Caption = "¼·©ñ"
    Else
        If UBound(timeLine) > OFP.CurrentPage Then
            SetTimerTime Me.hwnd, timeLine(OFP.CurrentPage + 1) * 1000
        ElseIf UBound(timeLine) = OFP.CurrentPage Then
            TimerFlag = 0
        End If
        Call Command7_Click(1)
    End If
    'Debug.Print "timer.run"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Debug.Print Button.Index
    Pic1MouseDown = False
    If Button.Index >= 7 Then
        '¥Ñ©óª«¥ó´N¦û¤F¨âºØ¼Ò¦¡ ©Ò¥H¨ä«áªº¥\¯à³£­n¥[1
        SysInfo.EdMode = Button.Index + 1
        
    Else
        
        SysInfo.EdMode = Button.Index
    End If
    If SysInfo.HideSelect = 1 Then Shape3.Visible = False
    If SysInfo.EdMode = 6 Then
        Toolbar2.Buttons(1).Value = tbrPressed
        Command6(0).Enabled = True
        Shape3.Visible = True
    Else
        'Shape3.Visible = False
    End If
    Select Case Button.Index
    
    
    End Select
    Call SetToolP
End Sub

Public Sub DoDraw(ByVal X As Integer, ByVal Y As Integer, ByVal tstr As String)
On Error GoTo out
    Call DoDraw_A(Arrf, X, Y, OFP.CurrentPage, tstr, SysInfo.ForColor)
    Call AD.ReDraw_Area(X, Y, X, Y)
Exit Sub
out:
    Debug.Print "DoDraw Error Out"
End Sub

Public Sub DoDrawBC(ByVal X As Integer, ByVal Y As Integer)

On Error GoTo out
        Call DoDrawBC_A(Arrf, X, Y, OFP.CurrentPage, SysInfo.BacColor)
        Call AD.ShowIt(X, Y)
Exit Sub
out:
    Debug.Print "DoDrawBC Error Out"

End Sub
Public Function DoErease(ByVal X As Single, ByVal Y As Single) As Integer
On Error GoTo out
    Dim tmpInt As Integer
    DoErease = DoErease_A(Arrf, X, Y, OFP.CurrentPage)
    'Call AD.ShowIt(X, Y)
    'If tmpInt <> 0 Then Call AD.ShowIt(X + tmpInt, Y)
'Pic1.Refresh
Exit Function
out:
Debug.Print "DoErease Error Out"
End Function
Public Function DoEreaseB(ByVal X As Integer, ByVal Y As Integer) As Integer
    On Error GoTo out

    Dim tmpInt As Integer
    DoEreaseB = DoEreaseB_A(Arrf, X, Y, OFP.CurrentPage)
    'Call AD.ShowIt(X, Y)
    'If tmpInt <> 0 Then Call AD.ShowIt(X + tmpInt, Y)
    'Call AD.ShowIt(X, Y)
Exit Function
out:
Debug.Print "DoEreaseB Error Out"
End Function

Public Sub DoEreaseAll()
    Dim tmpCL As ColorLayer
    tmpCL.Ansi = 0
    tmpCL.BColor = 0
    tmpCL.Color = 7
    For i = 0 To UBound(Arrf, 2)
        For j = 0 To UBound(Arrf, 1)
            Arrf(j, i, OFP.CurrentPage) = tmpCL
            DoEvents
        Next j
    Next i
    Pic1.Cls
    
    AD.SetTraget
    Call AD.ReShow_BG(0, 0, UBound(Arrf, 1), UBound(Arrf, 2))
End Sub


Public Sub SetColorBoard()
    '³]¸m½Õ¦â½L
    For i = 0 To 15
        Pic2.Line (i Mod 2, Fix(i / 2))-((i Mod 2) + 1, Fix(i / 2) + 1), QBColor(i), BF
    Next i

    For i = 0 To 7
        Pic3.Line (i Mod 2, Fix(i / 2))-((i Mod 2) + 1, Fix(i / 2) + 1), QBColor(i), BF
    Next i
End Sub
Public Sub GetForColor(ByVal X As Integer, ByVal Y As Integer)

    SysInfo.ForColor = X + 2 * Y
    'Debug.Print "SysInfo.ForColor=" & SysInfo.ForColor
    Shape2.BackColor = QBColor(SysInfo.ForColor)
    Text1.ForeColor = QBColor(SysInfo.ForColor)
    
End Sub
Public Sub GetBackColor(ByVal X As Integer, ByVal Y As Integer)

    SysInfo.BacColor = X + 2 * Y
    'Debug.Print "SysInfo.BacColor=" & SysInfo.BacColor
    Shape1.BackColor = QBColor(SysInfo.BacColor)
    Text1.BackColor = QBColor(SysInfo.BacColor)

End Sub

Public Sub SetSize(W As Integer, H As Integer, Z As Integer, filetype As Byte)
    Pic1.ScaleMode = 1
    Pic1.Height = H * SysInfo.PPA + 30
    Pic1.Width = (W / 2) * SysInfo.PPA + 30
    'Debug.Print "H * SysInfo.PPA + 30=" & (H * SysInfo.PPA + 30)
    Frame5.Width = Pic1.Width + 240
    Frame5.Height = Pic1.Height + 500
    
    If Frame5.Top + Frame5.Height > ToolP.Top + ToolP.Height Then
        Frame4.Top = Frame5.Height + Frame5.Top + 100
        'Frame6.Top = Frame5.Height + Frame5.Top + 1500
        'Debug.Print "Frame4.Top=" & Frame4.Top
    Else
        Frame4.Top = ToolP.Top + ToolP.Height + 100
        'Frame6.Top = 5640
    End If
    '³]©w¤u¨ãÄÝ©Êªº¦ì¸m
    Call SetToolPPos
    
    If Me.WindowState = 0 Then
        Me.Height = Frame4.Top + Frame4.Height + 850
        Me.Width = ToolP.Left + ToolP.Width + 150
        Me.Left = (Screen.Width - Me.Width) \ 2
        Me.Top = (Screen.Height - Me.Height) \ 2
        
    End If

    Pic1.ScaleHeight = H
    Pic1.ScaleWidth = W
    
    ReDim Arrf(0 To W - 1, 0 To H - 1, 1 To Z) As ColorLayer
    ReDim timeLine(1 To Z)  '©w¸q®É¶¡¶b
    Call SetTimeLine(1, Z, 1)   '³]©w®É¶¡¶b¹w³]­È
    
    Call ArrfPreValue
    '³]©wÀÉ®×ÄÝ©Ê
    OFP.filetype = filetype
    '³]©wÃ¸¹Ïª«¥óªºscale³æ¦ì
    OFP.CurrentPage = 1
    'ªì©l¤ÆùØÃ¸¹Ïª«¥ó
    'If AD.Traget <> "" Then
    '    Call AD.resetInDC
    'End If
    'AD.TwipsPerScaleX = 285 / (2 * Screen.TwipsPerPixelX)
    'AD.TwipsPerScaleY = 285 / Screen.TwipsPerPixelY
End Sub
Public Sub CreatAnsiTxt_Area_Bak(ByRef Ansitxt As String, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)

    Dim AnsiLine As String
    Dim tmpCCC As New ColorCodeCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    On Error Resume Next
    Ansitxt = "[m"

        For i = Y1 To Y2
            AnsiLine = ""
            For j = X1 To X2
                If Arrf(j, i, OFP.CurrentPage).Ansi <> -1 Then
                    AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, OFP.CurrentPage).Ansi, Arrf(j, i, OFP.CurrentPage).Color, Arrf(j, i, OFP.CurrentPage).BColor)
                End If
                DoEvents
            Next j
            If tmpCCC.preBColor = 0 Then
                Ansitxt = Ansitxt & RTrim(AnsiLine) & vbCrLf
            Else
                Ansitxt = Ansitxt & AnsiLine & vbCrLf
            End If
            'tmpCCC.Clear
        Next i
            Ansitxt = Ansitxt & "[m"
            tmpCCC.Clear

    'Text3.text = Ansitxt
End Sub
Public Sub CreatAnsiTxt_Area(ByRef Ansitxt As String, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)

    Dim AnsiLine As String
    Dim tmpCCC As New ColorCodeCreater
    Dim tmpL As Long
    Dim tmpTotal As Long

    On Error Resume Next
    Call BA_SetDefault  'ªì©l¤Æbytearray°}¦C¼Ò²Õ¹w³]­È
    Call BA_Reset     'ªì©l¤Æbytearray°}¦C
    Ansitxt = "[m"
    tmpL = 0
    tmpTotal = Y2 - Y1 + 1
    Call BA_Put_Str(Ansitxt)
        For i = Y1 To Y2
            
            For j = X1 To X2
                '''''''
                If Arrf(j, i, OFP.CurrentPage).Ansi <> -1 Then
                    
                    If Arrf(j + 1, i, OFP.CurrentPage).Ansi = -1 Then
                        tmpbyte = GetBiAsc(Arrf(j, i, OFP.CurrentPage).Ansi, 1)
                        If tmpbyte <> 0 Then
                            Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, OFP.CurrentPage).Color, Arrf(j, i, OFP.CurrentPage).BColor))
                            Call BA_Put(tmpbyte)
                        End If
                    Else
                    'AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        Call BA_Put_Str(tmpCCC.GetCode(Arrf(j, i, OFP.CurrentPage).Ansi, Arrf(j, i, OFP.CurrentPage).Color, Arrf(j, i, OFP.CurrentPage).BColor))
                    End If
                Else
                    tmpbyte = GetBiAsc(Arrf(j - 1, i, OFP.CurrentPage).Ansi, 0)
                    If tmpbyte <> 0 Then
                        'tmpStr = tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, OFP.CurrentPage).Color, Arrf(j, i, OFP.CurrentPage).BColor))
                        Call BA_Put(tmpbyte)
                    End If
                End If
                DoEvents
            Next j
            If i <> Y2 Then
                Call BA_Put(13)
                Call BA_Put(10)
            End If
            'tmpCCC.Clear
            tmpL = tmpL + 1
            Form12.Label2.Caption = tmpL & "/" & tmpTotal
        Next i
            'Call BA_Put_Str("[m" & the23line)
            Call BA_Put_Str("[m")
            tmpCCC.Clear
            Call BA_CutTail

    'Text3.text = Ansitxt
End Sub
Public Sub CreatAnsiTxt_VAA_v5(ByVal fromPage As Integer, ByVal toPage As Integer, Optional ByVal AddLine As String, Optional withtimeline As Byte)
    '¦¹ª©¥»¥Îbytearray
    '¦¹ª©¥»¥i«ü©w³sÄòªº­¶¼Æ
    'Dim AnsiLine As String
    Dim the23line As String
    Call BA_SetDefault  'ªì©l¤Æbytearray°}¦C¼Ò²Õ¹w³]­È
    Call BA_Reset     'ªì©l¤Æbytearray°}¦C
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            If withtimeline = 1 Then
                the23line = ""
            Else
                the23line = "[m" & AddLine & vbCrLf
            End If
        Case 3
            the23line = ""
    End Select
    Dim tmpCCC As New ColorCodeCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    Dim tmpTotal As Long
    'Dim tmpCount As Long
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    '°£¥h¥i¯àªº¿ù»~
    If fromPage < 1 Then fromPage = 1
    If toPage > maxZ Then toPage = maxZ
    
    tmpTotal = (maxY + 1) * maxZ
    Dim tmpbyte As Byte
    On Error Resume Next
    Dim tmpL As Long
    Dim tmpStr As String
    Ansitxt = "[m"
    For k = fromPage To toPage
        If withtimeline = 1 Then Call BA_Put_Str("^L" & timeLine(k) & vbCrLf)
        For i = 0 To maxY
            'AnsiLine = ""
            For j = 0 To maxX
                If Arrf(j, i, k).Ansi <> -1 Then
                    
                    If Arrf(j + 1, i, k).Ansi = -1 Then
                        tmpbyte = GetBiAsc(Arrf(j, i, k).Ansi, 1)
                        If tmpbyte <> 0 Then
                            Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                            Call BA_Put(tmpbyte)
                        End If
                    Else
                    'AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        Call BA_Put_Str(tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                    End If
                Else
                    tmpbyte = GetBiAsc(Arrf(j - 1, i, k).Ansi, 0)
                    If tmpbyte <> 0 Then
                        'tmpStr = tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                        Call BA_Put(tmpbyte)
                    End If
                End If
                DoEvents
            Next j
            '´«¦æ
            Call BA_Put_Str(Ansitxt)
            tmpCCC.Clear
            Call BA_Put(13)
            Call BA_Put(10)
            'tmpCCC.Clear
            tmpL = tmpL + 1
            Form12.Label2.Caption = tmpL & "/" & tmpTotal
        Next i
            'Ansitxt = Ansitxt & "[m" & k & "-" & tmpL & the23line
            Call BA_Put_Str(the23line)
            tmpCCC.Clear
    Next k
    Call BA_CutTail
    'Text3.text = Ansitxt
End Sub
Public Sub CreatAnsiTxt_VAA_v4(Optional ByVal AddLine As String, Optional withtimeline As Byte)
    '¦¹ª©¥»¥Îbytearray
    
    'Dim AnsiLine As String
    Dim the23line As String
    Call BA_SetDefault  'ªì©l¤Æbytearray°}¦C¼Ò²Õ¹w³]­È
    Call BA_Reset     'ªì©l¤Æbytearray°}¦C
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            If withtimeline = 1 Then
                the23line = ""
            Else
                the23line = "[m" & AddLine & vbCrLf
            End If
        Case 3
            the23line = ""
    End Select
    Dim tmpCCC As New ColorCodeCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    Dim tmpTotal As Long
    'Dim tmpCount As Long
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    tmpTotal = (maxY + 1) * maxZ
    Dim tmpbyte As Byte
    On Error Resume Next
    Dim tmpL As Long
    Dim tmpStr As String
    Ansitxt = "[m"
    For k = 1 To maxZ
        If withtimeline = 1 Then Call BA_Put_Str("^L" & timeLine(k) & vbCrLf)
        For i = 0 To maxY
            'AnsiLine = ""
            For j = 0 To maxX
                If Arrf(j, i, k).Ansi <> -1 Then
                    
                    If Arrf(j + 1, i, k).Ansi = -1 Then
                        tmpbyte = GetBiAsc(Arrf(j, i, k).Ansi, 1)
                        If tmpbyte <> 0 Then
                            Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                            Call BA_Put(tmpbyte)
                        End If
                    Else
                    'AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        Call BA_Put_Str(tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                    End If
                Else
                    tmpbyte = GetBiAsc(Arrf(j - 1, i, k).Ansi, 0)
                    If tmpbyte <> 0 Then
                        'tmpStr = tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                        Call BA_Put(tmpbyte)
                    End If
                End If
                DoEvents
            Next j
            '´«¦æ
            Call BA_Put_Str(Ansitxt)
            tmpCCC.Clear
            Call BA_Put(13)
            Call BA_Put(10)
            'tmpCCC.Clear
            tmpL = tmpL + 1
            Form12.Label2.Caption = tmpL & "/" & tmpTotal
        Next i
            'Ansitxt = Ansitxt & "[m" & k & "-" & tmpL & the23line
            Call BA_Put_Str(the23line)
            tmpCCC.Clear
    Next k
    Call BA_CutTail
    'Text3.text = Ansitxt
End Sub

Public Sub CreatAnsiTxt_VAA_v3(ByRef Ansitxt As String, Optional ByVal AddLine As String)

    Dim AnsiLine As String
    Dim the23line As String
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            the23line = AddLine & vbCrLf
        Case 3
            the23line = ""
    End Select
    Dim tmpCCC As New ColorCodeCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    On Error Resume Next
    Dim tmpL As Long
    Ansitxt = "[m"
    For k = 1 To maxZ
        For i = 0 To maxY
            AnsiLine = ""
            For j = 0 To maxX
                If Arrf(j, i, k).Ansi <> -1 Then
                    AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                End If
                DoEvents
            Next j
            If tmpCCC.preBColor = 0 Then
                Ansitxt = Ansitxt & RTrim(AnsiLine) & vbCrLf
            Else
                Ansitxt = Ansitxt & AnsiLine & vbCrLf
            End If
            'tmpCCC.Clear
            tmpL = tmpL + 1
        Next i
            Ansitxt = Ansitxt & "[m" & k & "-" & tmpL & the23line
            tmpCCC.Clear
    Next k
    'Text3.text = Ansitxt
End Sub
Public Function CreatHTML_VAA(ByRef HTML As String, ByVal fromPage As Integer, ByVal toPage As Integer, ByVal title As String, ByVal autoplay_ctl As Integer, ByVal forcetime As Single, Optional ByVal font_size As String)
    '¦¹ª©¥»¥Îbytearray
    '¦¹ª©¥»¥i«ü©w³sÄòªº­¶¼Æ
    'Dim AnsiLine As String
    Dim the23line As String
    'Call BA_SetDefault  'ªì©l¤Æbytearray°}¦C¼Ò²Õ¹w³]­È
    'Call BA_Reset     'ªì©l¤Æbytearray°}¦C
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            If withtimeline = 1 Then
                the23line = ""
            Else
                the23line = "[m" & AddLine & vbCrLf
            End If
        Case 3
            the23line = ""
    End Select
    Dim tmpCCC As New HTMLCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    Dim tmpTotal As Long
    'Dim tmpCount As Long
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    '°£¥h¥i¯àªº¿ù»~
    If fromPage < 1 Then fromPage = 1
    If toPage > maxZ Then toPage = maxZ
    
    tmpTotal = (maxY + 1) * maxZ
    Dim tmpbyte As Byte
    On Error Resume Next
    Dim tmpL As Long
    Dim tmpStr As String
    'Ansitxt = "[m"
    
    HTML = "<div id='bbsmovie_info' pagecount='" & (toPage - fromPage + 1) & "' forcetime='" & forcetime & "' autoplay_ctl='" & autoplay_ctl & "' ></div>"
    For k = fromPage To toPage
        'If withTimeLine = 1 Then Call BA_Put_Str("^L" & timeLine(k) & vbCrLf)
        HTML = HTML & "<div id='p" & k - 1 & "' time='" & timeLine(k) & "' class='bbs'>" & vbCrLf
        For i = 0 To maxY
            tmpCCC.Clear
            AnsiLine = "<div>"
            
            For j = 0 To maxX
                If Arrf(j, i, k).Ansi <> -1 Then
                    
                    If Arrf(j + 1, i, k).Ansi = -1 Then
                        tmpbyte = GetBiAsc(Arrf(j, i, k).Ansi, 1)
                        If tmpbyte <> 0 Then
                            If Arrf(j, i, k).Color = Arrf(j + 1, i, k).Color And Arrf(j, i, k).BColor = Arrf(j + 1, i, k).BColor Then
                                AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                            Else
                                
                                
                                AnsiLine = AnsiLine & tmpCCC.linetail & tmpCCC.GetCode_Bi(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor, 0)
                                tmpCCC.Clear
                            End If
                            'Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                            
                            'Call BA_Put(tmpbyte)
                        End If
                    Else
                    AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        'Call BA_Put_Str(tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                    End If
                Else
                    tmpbyte = GetBiAsc(Arrf(j - 1, i, k).Ansi, 0)
                    If tmpbyte <> 0 Then
                        'tmpStr = tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        'AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                        If Arrf(j - 1, i, k).Color <> Arrf(j, i, k).Color Or Arrf(j - 1, i, k).BColor <> Arrf(j, i, k).BColor Then
                            AnsiLine = AnsiLine & tmpCCC.GetCode_Bi(Arrf(j - 1, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor, 1) & "<div>"
                        End If
                        
                        'Call BA_Put_Str(tmpCCC.GetCode_noChar(Arrf(j, i, k).Color, Arrf(j, i, k).BColor))
                        'Call BA_Put(tmpbyte)
                    End If
                End If
                DoEvents
            Next j
            '´«¦æ
            'Call BA_Put_Str(Ansitxt)
            
            'Call BA_Put(13)
            'Call BA_Put(10)
            'tmpCCC.Clear
            HTML = HTML & AnsiLine & tmpCCC.linetail & "<br />" & vbCrLf
            tmpL = tmpL + 1
            Form12.Label2.Caption = tmpL & "/" & tmpTotal
        Next i
            'Ansitxt = Ansitxt & "[m" & k & "-" & tmpL & the23line
            HTML = HTML & vbCrLf & "</div>" & vbCrLf
            'Call BA_Put_Str(the23line)
            tmpCCC.Clear
    Next k
    
    '¦X¦¨html
    If font_size = "" Then
        font_size = "18pt"
    End If
    HTML = Replace(Replace(LoadResString(101), "$title", title), "$font_size", font_size) & HTML & LoadResString(102)
    
    'Call BA_CutTail
    'Text3.text = Ansitxt
End Function
Public Sub CreatAnsiTxt_VAA(ByRef Ansitxt As String, Optional ByVal AddLine As String)  'ÂÂª©ªº½sÄ¶¨ç¦¡
    'Dim Ansitxt As String
    Dim AnsiLine As String
    Dim the23line As String
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            the23line = AddLine & vbCrLf
        Case 3
            the23line = ""
    End Select
On Error Resume Next
    For k = 1 To UBound(Arrf, 3)
        For i = 0 To UBound(Arrf, 2)
            AnsiLine = ""
            Select Case Arrf(0, i, k).Ansi
                Case Is = 0 And 32
                    Arrf(0, i, k).Color = 7
                    If Arrf(0, i, k).BColor = 0 Then
                        AnsiLine = AnsiLine & Chr(32)
                    Else
                        If Arrf(0, i, k).BColor = 0 Then
                            AnsiLine = AnsiLine & "[m" & Chr(32)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(0, i, k).BColor) & "m" & Chr(32)
                        End If
                    End If
                    
                Case Is = -1
            
                Case Else
                    If (Arrf(0, i, k).Color = 7) And (Arrf(0, i, k).BColor = 0) Then
                
                        AnsiLine = AnsiLine & Chr(Arrf(0, i, k).Ansi)
                    Else
                
                        If QBCToAnsiC(Arrf(0, i, k).Color) = "" Then
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(0, i, k).BColor) & QBCToAnsiC(Arrf(0, i, k).Color) & "m" & Chr(Arrf(0, i, k).Ansi)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(0, i, k).BColor) & ";" & QBCToAnsiC(Arrf(0, i, k).Color) & "m" & Chr(Arrf(0, i, k).Ansi)
                        End If
                    End If
            End Select
        For j = 1 To UBound(Arrf, 1)
            Select Case Arrf(j, i, k).Ansi
                Case Is = 0 And 32
                    Arrf(j, i, k).Color = 7
                    If Arrf(j, i, k).BColor = Arrf(j - 1, i, k).BColor Then
                        AnsiLine = AnsiLine & Chr(32)
                    Else
                        If Arrf(j, i, k).BColor = 0 Then
                            AnsiLine = AnsiLine & "[m" & Chr(32)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(j, i, k).BColor) & "m" & Chr(32)
                        End If
                    End If
                        
                Case Is = -1
                
                Case Else
                    If (Arrf(j, i, k).Color = Arrf(j - 1, i, k).Color) And (Arrf(j, i, k).BColor = Arrf(j - 1, i, k).BColor) And (Arrf(j - 1, i, k).Ansi <> 0) Then
                
                        AnsiLine = AnsiLine & Chr(Arrf(j, i, k).Ansi)
                    Else
                
                        If Arrf(j, i, k).Color = 7 And Arrf(j, i, k).BColor = 0 Then
                            AnsiLine = AnsiLine & "[m" & Chr(Arrf(j, i, k).Ansi)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiC(Arrf(j, i, k).Color) & ";" & QBCToAnsiBC(Arrf(j, i, k).BColor) & "m" & Chr(Arrf(j, i, k).Ansi)
                        End If
                    End If
            End Select
        'Debug.Print "Arrf(" & j & "," & i & ").Ansi=" & Arrf(j, i, k).Ansi
            DoEvents
        Next j
        Ansitxt = Ansitxt & "[m" & RTrim(AnsiLine) & vbCrLf
        DoEvents
        Next i
    '²Ä23¦æªº­«½Æ
    Ansitxt = Ansitxt & the23line
Next k
'Text3.text = Ansitxt

End Sub
Public Sub CreatAnsiTxt_NoColor(ByRef Ansitxt As String)
'Dim Ansitxt As String
    Dim AnsiLine As String
    Dim the23line As String
    Select Case OFP.filetype
        Case 3
            the23line = ""
        Case 2
            the23line = vbCrLf
        Case 3
            the23line = ""
    End Select
    On Error Resume Next
    For k = 1 To UBound(Arrf, 3)
        For i = 0 To UBound(Arrf, 2)
            AnsiLine = ""
        
            For j = 0 To UBound(Arrf, 1)
                Select Case Arrf(j, i, k).Ansi
                    Case Is = 0
                            AnsiLine = AnsiLine & Chr(32)
                    Case Is = -1
                    
                    Case Else
                            AnsiLine = AnsiLine & Chr(Arrf(j, i, k).Ansi)
                End Select
                'Debug.Print "Arrf(" & j & "," & i & ").Ansi=" & Arrf(j, i, k).Ansi
                DoEvents
            Next j
            Ansitxt = Ansitxt & RTrim(AnsiLine) & vbCrLf
            DoEvents
        Next i
        '²Ä23¦æªº­«½Æ
        Ansitxt = Ansitxt & the23line
    Next k
    'Text3.text = Ansitxt
    
End Sub
Public Sub SetNowAnsi()
    On Error GoTo out
    If Option1.Value = True Then
        NowAnsi = Text1.text
        'Toolbar3.Buttons(4).Enabled = True
        SysInfo.ForeSource = 1
        SysInfo.LastAnsi = Asc(NowAnsi)
        CC(O_Pen_Text).Value = False
        Option3.Value = False
    
    ElseIf CC(O_Pen_Text).Value = True Then
        'Toolbar3.Buttons(4).Enabled = False
        'If SysInfo.EdMode = 4 Then
        '    Toolbar3.Buttons(1).Value = tbrPressed
        '    SysInfo.EdMode = 1
        'End If
    
        SysInfo.ForeSource = 2
    

    ElseIf Option3.Value = True Then
        NowAnsi = Text6.text
        'Toolbar3.Buttons(4).Enabled = True
        
        SysInfo.ForeSource = 3
        SysInfo.SSAnsi = Asc(NowAnsi)
    End If
Exit Sub
out:
    Debug.Print "SetNowAnsi Error Out"
End Sub

Public Sub DoMutiDraw(ByVal X As Integer, ByVal Y As Integer, ByVal InAnsis As String, Optional ByVal Mode As Byte)
    
    Dim Pointer As Integer
    Dim tmpInt As Integer
    'Debug.Print "InAnsis=" & InAnsis
    Dim tmpstrA() As String
    tmpstrA = Split(InAnsis, vbCrLf)
    For j = 0 To UBound(tmpstrA)
        Pointer = 0
        If Mode = 1 Then
            tmpInt = Tlen(tmpstrA(j))
            Call AD.resetBG_Area(X + Pointer, Y + j, X + Pointer + tmpInt - 1, Y + j)
            Call DoErease_Area(X + Pointer, Y + j, X + Pointer + tmpInt - 1, Y + j)
            Call AD.ReDraw_Area(X + Pointer, Y + j, X + Pointer + tmpInt - 1, Y + j)
            
            'Call DoErease(X + Pointer, Y + j)
        End If
        For i = 1 To Len(tmpstrA(j))
            tstr = Mid(tmpstrA(j), i, 1)

             Call DoDraw(X + Pointer, Y + j, tstr)
             'Debug.Print j & "-" & i & ":" & Asc(tstr) & ";" & tstr
            If CC(C_BG).Value = 1 Then Call DoDrawBC(X + Pointer, Y + j)
            Pointer = Pointer + Tlen(tstr)
            DoEvents
        Next i
    Next j
    SysInfo.cDrawPos_X = X + Pointer
    SysInfo.cDrawPos_Y = Y + j
   
End Sub
Public Sub ShowIt_bak(ByVal X As Single, ByVal Y As Single)
On Error GoTo out
    
    If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
        'AD.PrintText "¢m", QBColor(Arrf(X, Y, OFP.CurrentPage).BColor), X - 1, Y
        AD.DrawRectangle X - 1, Y, X - 1, Y, Arrf(X, Y, OFP.CurrentPage).BColor
    End If
    'AD.PrintText "¢m", QBColor(Arrf(X, Y, OFP.CurrentPage).BColor), X, Y
    AD.DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
    If X <> UBound(Arrf, 1) Then
        If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
            'AD.PrintText "¢m", QBColor(Arrf(X, Y, OFP.CurrentPage).BColor), X + 1, Y
            AD.DrawRectangle X + 1, Y, X + 1, Y, Arrf(X, Y, OFP.CurrentPage).BColor
        End If
    End If
    If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
        If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X - 1, Y, OFP.CurrentPage).Color), X - 1, Y
        'If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), Arrf(X - 1, Y, OFP.CurrentPage).Color, X - 1, Y
    Else
        If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        'If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), Arrf(X, Y, OFP.CurrentPage).Color, X, Y
    End If

Exit Sub
out:

    Debug.Print "ShowIt::Error:" & Err.Description & "->In:" & "(" & X & "," & Y & ")"
End Sub
Public Sub ShowIt(ByVal X As Long, ByVal Y As Long)
On Error GoTo out
    '·sª©¥»ªºÅã¥Ü¨ç¦¡
    '¤ä´©Âù¦â¦r
    If X > UBound(Arrf, 1) Then Exit Sub
    If X <= UBound(Arrf, 1) Then
        If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
            AD.DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            AD.DrawRectangle X + 1, Y, X + 1, Y, Arrf(X + 1, Y, OFP.CurrentPage).BColor
            AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X + 1, Y, OFP.CurrentPage).Color), X, Y
            AD.PrintText_biByte_Left Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        ElseIf Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
            AD.DrawRectangle X - 1, Y, X - 1, Y, Arrf(X - 1, Y, OFP.CurrentPage).BColor
            AD.DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            AD.PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X - 1, Y
            AD.PrintText_biByte_Left Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X - 1, Y, OFP.CurrentPage).Color), X - 1, Y
        Else
            AD.DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        End If
    Else
        AD.DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
        AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
    End If
    
Exit Sub
out:
    Debug.Print "ShowIt::Error:" & Err.Description & "->In:" & "(" & X & "," & Y & ")"
End Sub

Public Sub Setoolbar()
    'If CC(C_BG).value = 1 And Check1.value = 0 Then
    '    Toolbar3.Buttons(3).Enabled = False
    'Else
    '    Toolbar3.Buttons(3).Enabled = True
    'End If
End Sub

Public Sub PaintColor(ByVal X As Integer, ByVal Y As Integer)
On Error GoTo out
    Call PaintColor_A(Arrf, X, Y, OFP.CurrentPage, SysInfo.ForColor, SysInfo.BacColor, CC(C_Fore).Value, CC(C_BG).Value)
    Call AD.ReDraw_Area(X, Y, X, Y)
    Pic1.Refresh
Exit Sub
out:

Debug.Print "Paint Color Error Out"
Resume Next
End Sub


Private Sub SetToolP()
    Dim preTP As Frame
    Set preTP = ToolP
    
    Select Case SysInfo.EdMode
    Case Is = 1
        Set ToolP = Frame3
        Frame3.Caption = "µeµ§"
        Frame12.Visible = True
        Frame13.Visible = False
    Case Is = 2
        Set ToolP = Frame8
    Case Is = 3
        Set ToolP = Frame3
        Frame3.Caption = "µeµ§"
        Frame12.Visible = True
        Frame13.Visible = False
    Case Is = 4
        Set ToolP = Frame3
        Frame3.Caption = "µeµ§"
        Frame12.Visible = True
        Frame13.Visible = False
    Case Is = 5
        Set ToolP = Frame10
    Case Is = 6
        Set ToolP = Frame9
    Case Is = 8
        Set ToolP = Frame7
    Case Is = 9
        Set ToolP = Frame11
    Case Is = 10
        Set ToolP = Frame3
        Frame3.Caption = "¿é¤J¼Ò¦¡"
        Frame12.Visible = False
        Frame13.Visible = True
    Case Is = 11
        Set ToolP = Frame14
        'C
    Case Is = 12
        Set ToolP = Frame15
    End Select
    preTP.Visible = False
    ToolP.Left = preTP.Left
    ToolP.Top = preTP.Top
    ToolP.Visible = True
    If SysInfo.EdMode = 10 And Frame3.Enabled = True And Me.Visible = True Then
        Text2.SetFocus
    End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

    Case Is = 1
        SysInfo.EdMode = 6
    Case Is = 2
        SysInfo.EdMode = 7
        If List1.ListCount <> 0 And List1.ListIndex = -1 Then List1.ListIndex = 0
    Case Is = 3
    

End Select
If SysInfo.EdMode = 7 Then
    Command6(0).Enabled = False
'    Command6(1).Enabled = False
'    Command6(2).Enabled = False

Else
    Command6(0).Enabled = True
End If

End Sub

Public Sub Set_FileType_Visual()
If OFP.filetype = 1 Then
    Command7(0).Visible = False
    Command7(1).Visible = False
    Command12.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command11.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Label9.Visible = False
    Combo1.Visible = False
Else
    Command7(0).Visible = True
    Command7(1).Visible = True
    Command12.Visible = True
    Command9.Visible = True
    Command10.Visible = True
    Command11.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Text7.Visible = True
    Text8.Visible = True
    Label9.Visible = True
    Combo1.Visible = True
End If
If OFP.filetype = 2 Then    '·íÀÉ®×Ãþ«¬¬°°Êµe®É¶}±Òªº¥\¯à
    Command14.Visible = True
    Me_Display.Enabled = True
    Me_Director.Enabled = True
    '®É¶¡¶b±±¨î
    Label15.Visible = True
    Text9.Visible = True
    Command9.Visible = True
    Command21.Visible = True
Else
    Command14.Visible = False
    Me_Display.Enabled = False
    Me_Director.Enabled = False
    '®É¶¡¶b±±¨î
    Label15.Visible = False
    Text9.Visible = False
    Command9.Visible = False
    Command21.Visible = False
End If

End Sub

Public Sub Set_VAA_Combo()
Combo1.Clear
For i = 1 To UBound(Arrf, 3)
    Combo1.AddItem "²Ä" & i & "­¶"


Next i

End Sub

Public Sub CheckClose()
On Error GoTo out
If OFP.Closed = True Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Me_Save.Enabled = False
    Me_SaveAs.Enabled = False
    Me_View.Enabled = False
    Me_Compile.Enabled = False
    Me_Tool.Enabled = False
    ToolP.Enabled = False
Else
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Me_Save.Enabled = True
    Me_SaveAs.Enabled = True
    Me_View.Enabled = True
    Me_Compile.Enabled = True
    Me_Tool.Enabled = True
    If SysInfo.EdMode = 10 Then Text2.SetFocus
    ToolP.Enabled = True
End If
Exit Sub
out:
Debug.Print "Error¡GCheckClose "

End Sub

Public Sub SetMoninter()

End Sub

Public Sub ForScreenSize(ByVal Frontsize As Byte)

Select Case Frontsize
    Case 12
        SysInfo.PPA = 225
    Case 13
        SysInfo.PPA = 255
    Case 14
        SysInfo.PPA = 285
    Case 15
        SysInfo.PPA = 315

End Select
Pic1.FontSize = Frontsize

'Pic1.FontSize = 13
'SysInfo.PPA = 255
'Pic1.FontSize = 12
'SysInfo.PPA = 225


'On Error Resume Next
'For Each object In me
'Debug.Print object.Name & ".FontSize=" & object.FontSize
'object.FontSize = 8
'Next

End Sub

Public Sub DisplayVAA()
    If TimerFlag = 0 Then
        'Timer1.Interval = 1000
        TimerFlag = 1
        SetTimerTime Me.hwnd, timeLine(OFP.CurrentPage) * 1000
        Command14.Caption = "°±¤î¼·©ñ"
        Me_Display.Caption = "°±¤î¼·©ñ"
        
    Else
        TimerFlag = 0
        Command14.Caption = "¼·©ñ"
        Me_Display.Caption = "¼·©ñ"
        Timer1.Interval = 0
    End If
End Sub

Public Sub SelectAnsi(ByVal X As Integer, ByVal Y As Integer)
    Dim dX As Integer
    dX = 0
    If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
        X = X - 1
        dX = 1
    End If
    If X < UBound(Arrf, 1) Then
        If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
            dX = 1
        End If
    End If
    SL.StartPoint_X = X
    SL.StartPoint_Y = Y
    SL.EndPoint_X = X + dX
    SL.EndPoint_Y = Y
    SL.DrawSelect
End Sub

Public Sub SetConfic()
Call WriteConfic(App.Path & "\confic.cfg", SysInfo)
End Sub
Public Sub GetConfic()
    Dim ConficDS As SysEnv
    ConficDS = ReadConfic(App.Path & "\confic.cfg")
    'Åª¨ú¤W¦¸Â÷¶}®É©Ò¿ï¾Üªº¤u¨ã
    Select Case ConficDS.EdMode
        Case 0
        Case 1
            Toolbar3.Buttons(1).Value = tbrPressed
        Case 2
            Toolbar3.Buttons(2).Value = tbrPressed
        Case 3
            Toolbar3.Buttons(1).Value = tbrPressed
            Check7.Value = 1
        Case 4
            Toolbar3.Buttons(1).Value = tbrPressed
            Check6.Value = 1
        Case 5
            Toolbar3.Buttons(3).Value = tbrPressed
        Case 6
            Toolbar3.Buttons(4).Value = tbrPressed
            Toolbar2.Buttons(1).Value = tbrPressed
        Case 7
            Toolbar3.Buttons(4).Value = tbrPressed
            Toolbar2.Buttons(2).Value = tbrPressed
        Case 8
            Toolbar3.Buttons(5).Value = tbrPressed
        Case 9
            Toolbar3.Buttons(6).Value = tbrPressed
        Case 10
            Toolbar3.Buttons(7).Value = tbrPressed
            DelayFlag = 1
        Case 11
            Toolbar3.Buttons(8).Value = tbrPressed
        Case 12
            Toolbar3.Buttons(9).Value = tbrPressed
    End Select
     SysInfo.EdMode = ConficDS.EdMode
    '¤W¦¸Â÷¶}®É©Ò¨Ï¥Î¤§µ§Ä²
    On Error GoTo engout
    If ConficDS.SSAnsi <> 0 Then
        Text6.text = Chr(ConficDS.SSAnsi)
    End If
    'Debug.Print "ConficDS.SSAnsi" & ConficDS.SSAnsi
    If ConficDS.LastAnsi <> 0 Then
        
        Text1.text = Chr(ConficDS.LastAnsi)
        'Debug.Print "ConficDS.LastAnsi=" & ConficDS.LastAnsi
    End If
    '¤W¦¸Â÷¶}®É©Ò¿ï¾Üªº«e´º¨Ó·½
    Select Case ConficDS.ForeSource
        Case 0
        Case 1
            Option1.Value = True
        Case 2
            CC(O_Pen_Text).Value = True
        Case 3
            Option3.Value = True
    End Select
    
    
    If ConficDS.Frontsize <> 0 Then SysInfo.Frontsize = ConficDS.Frontsize
    If ConficDS.ForColor <> 0 Then
        SysInfo.ForColor = ConficDS.ForColor
        'Debug.Print "SysInfo.ForColor=" & SysInfo.ForColor
        Shape2.BackColor = QBColor(SysInfo.ForColor)
        Text1.ForeColor = QBColor(SysInfo.ForColor)
    End If
    If ConficDS.BacColor <> 0 Then
        SysInfo.BacColor = ConficDS.BacColor
        'Debug.Print "SysInfo.BacColor=" & SysInfo.BacColor
        Shape1.BackColor = QBColor(SysInfo.BacColor)
        Text1.BackColor = QBColor(SysInfo.BacColor)
    End If
    'Ãö³¬¤@­ÓÀÉ®×®É½T»{Àx¦s
    SysInfo.CheckSave = ConficDS.CheckSave
    'ÁôÂÃ¿ï¾Ü®Ø
    SysInfo.HideSelect = ConficDS.HideSelect
    '¤u¨ãÄÝ©Ê¸m©³
    SysInfo.ToolPBoxDown = ConficDS.ToolPBoxDown
Exit Sub
engout:
    MsgBox "½Ð±N±zªºwindows¦a°Ï¿ï¶µ§ï¬°¤¤¤å(¥xÆW)¥»µ{¦¡¤~¯à¥¿±`¹B§@ (¯S§O·PÁÂsuperlubu§f¥¬´£¨Ñ°£¿ù)"
    Resume Next
End Sub

Public Sub AskSave()
If OFP.IsChanged = True Then
        If MsgBox("¬O§_­n¦bÃö³¬«eÀx¦s²{¦b³o­ÓÀÉ®×" & vbCrLf & OFP.FilePath, 36, "´£¿ô") = 6 Then
            Call Me_Save_Click
        End If
End If
End Sub
'>>>>>>>>>>³]©wµøµ¡¼ÐÃD
Public Sub SetFormCaption(Optional str As String)
Dim tmpStr As String
tmpStr = "Visual Ansi 2008 apha (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
If str <> "" Then

    Me.Caption = tmpStr & " - " & str
Else

    If OFP.FilePath = "" Then
        Me.Caption = tmpStr
    Else
        If OFP.IsChanged = True Then
            Me.Caption = tmpStr & " - " & FileSys.GetFileName(OFP.FilePath) & " * "
        Else
            Me.Caption = tmpStr & " - " & FileSys.GetFileName(OFP.FilePath)
        End If

    End If
End If
End Sub

Public Sub ArrfPreValue()
For k = 1 To UBound(Arrf, 3)
    For j = 0 To UBound(Arrf, 2)
        For i = 0 To UBound(Arrf, 1)
            Arrf(i, j, k).Color = 7
            DoEvents
        Next i
    Next j
Next k
End Sub
'>>>>>>>>>>ÃC¦â¸m´«¥\¯à
Public Sub ExChColor_SetFColor(ByVal X As Integer, ByVal Y As Integer)
If SysInfo.ExChColor.CurrentSel < 2 Then
    SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel) = X + 2 * Y
    Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BackColor = QBColor(SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel))
End If

End Sub
Public Sub ExChColor_SetBColor(ByVal X As Integer, ByVal Y As Integer)
If SysInfo.ExChColor.CurrentSel > 1 Then
    SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel) = X + 2 * Y
     Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BackColor = QBColor(SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel))
    
End If
End Sub
Public Sub ExChColor_Draw(ByVal X As Integer, ByVal Y As Integer)
On Error GoTo out
'Debug.Print "SysInfo.ExChColor.Color(0)=" & SysInfo.ExChColor.Color(0) & vbCrLf & "Arrf(X, Y, OFP.CurrentPage).Color=" & Arrf(X, Y, OFP.CurrentPage).Color
    Call ExChColor_Draw_A(Arrf, X, Y, OFP.CurrentPage, SysInfo.ExChColor.Color(0), SysInfo.ExChColor.Color(2), SysInfo.ExChColor.Color(1), SysInfo.ExChColor.Color(3), CC(C_Fore).Value, CC(C_BG).Value)
    Call AD.ReDraw_Area(X, Y, X, Y)
    Pic1.Refresh
Exit Sub
out:

Debug.Print "ExChColor_Draw  Error Out"
Resume Next
End Sub


Public Sub SetToolPPos()
On Error GoTo out
If SysInfo.ToolPBoxDown = 1 Then
    
    ToolP.Left = Frame4.Left + Frame4.Width
    ToolP.Top = Frame4.Top
    If Me.Height < ToolP.Top + ToolP.Height + 850 Then Me.Height = ToolP.Top + ToolP.Height + 850
Else
    If Frame5.Left + Frame5.Width >= Frame1.Left + Frame1.Width Then
        ToolP.Left = Frame5.Left + Frame5.Width
        'Debug.Print "ToolP.Left=" & ToolP.Left
    Else
        ToolP.Left = Frame1.Left + Frame1.Width

    End If
End If
Exit Sub
out:
End Sub

Public Sub DoErease_Area_In(ByVal tmpX1 As Integer, ByVal tmpY1 As Integer, ByVal tmpX2 As Integer, ByVal tmpY2 As Integer)
    On Error GoTo out
            For j = tmpY1 To tmpY2
                
                For i = tmpX1 To tmpX2
                    If Not (i = tmpX1 And Arrf(tmpX1, j, OFP.CurrentPage).Ansi = -1) And Not (i = tmpX2 And Arrf(tmpX2 + 1, j, OFP.CurrentPage).Ansi = -1) Then
                        Call DoErease_A(Arrf, i, j, OFP.CurrentPage)
                        Call DoEreaseB_A(Arrf, i, j, OFP.CurrentPage)
                        DoEvents
                    End If

                Next i
            Next j
            Call AD.ReDraw_Area(tmpX1 - 1, tmpY1, tmpX2 + 1, tmpY2)
Exit Sub
out:
End Sub
Public Sub DoErease_Area(ByVal tmpX1 As Integer, ByVal tmpY1 As Integer, ByVal tmpX2 As Integer, ByVal tmpY2 As Integer)
    On Error GoTo out
            For j = tmpY1 To tmpY2
                
                For i = tmpX1 To tmpX2
                        If CC(C_Fore).Value = 1 Then Call DoErease_A(Arrf, i, j, OFP.CurrentPage)
                        If CC(C_BG).Value = 1 Then Call DoEreaseB_A(Arrf, i, j, OFP.CurrentPage)
                        DoEvents
                Next i
            Next j
            Call AD.ReDraw_Area(tmpX1 - 1, tmpY1, tmpX2 + 1, tmpY2)
Exit Sub
out:
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Debug.Print Button.Index
    Pic1MouseDown = False

    If SysInfo.HideSelect = 1 Then Shape3.Visible = False

    Select Case Button.Index
        Case 1
            
            If Check6.Value = 1 Then
                SysInfo.EdMode = 4
            ElseIf Check7.Value = 1 Then
                SysInfo.EdMode = 3
            Else
                SysInfo.EdMode = 1
            End If
            
        Case 2
            SysInfo.EdMode = 2
        Case 3
            SysInfo.EdMode = 5
        Case 4
            SysInfo.EdMode = 6
            Toolbar2.Buttons(1).Value = tbrPressed
            Command6(0).Enabled = True
            Shape3.Visible = True
            
        Case 5
            SysInfo.EdMode = 8
        Case 6
            SysInfo.EdMode = 9
        Case 7
            SysInfo.EdMode = 10
            DelayFlag = 1
        Case 8
            SysInfo.EdMode = 11 '¹Ï¤ù
        Case 9
            SysInfo.EdMode = 12 '®ÄªG ²¾°Ê
    End Select
    Call SetToolP
End Sub
