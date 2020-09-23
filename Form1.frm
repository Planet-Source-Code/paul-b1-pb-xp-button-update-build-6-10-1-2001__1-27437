VERSION 5.00
Object = "{F49365FC-E8A5-4E38-9DBC-DAA7D889B8A3}#1.6#0"; "pbxpbutton.ocx"
Begin VB.Form Form11 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PB XP Button"
   ClientHeight    =   3930
   ClientLeft      =   7110
   ClientTop       =   4200
   ClientWidth     =   6255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin PB_XP_Button.PBXPButton PBXPButton4 
      Height          =   390
      Index           =   0
      Left            =   165
      TabIndex        =   32
      Top             =   3030
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   688
      Caption         =   "Options"
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":0E02
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIconOver=   13811126
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checked         =   -1  'True
      Value           =   -1  'True
      CheckedColor    =   14211029
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show Focus Rect"
      Height          =   240
      Left            =   1680
      TabIndex        =   28
      Top             =   2595
      Value           =   1  'Checked
      Width           =   1665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3765
      TabIndex        =   23
      Text            =   "2"
      Top             =   495
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Disable"
      Height          =   285
      Left            =   1695
      TabIndex        =   19
      Top             =   1980
      Value           =   1  'Checked
      Width           =   1170
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   0
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":119C
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton2 
      Height          =   435
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   1380
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      Caption         =   "&Checked"
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":1536
      BorderColorDown =   6956042
      BackColor       =   -2147483644
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIconOver=   13811126
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
      Checked         =   -1  'True
      Value           =   -1  'True
      CheckedColor    =   14211029
   End
   Begin PB_XP_Button.PBXPButton PBXPButton1 
      Height          =   450
      Left            =   4905
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   794
      Caption         =   "  Close "
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":18D0
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlignCaption    =   2
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   1
      Left            =   405
      TabIndex        =   3
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":25C2
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   2
      Left            =   795
      TabIndex        =   4
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":295C
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   3
      Left            =   1185
      TabIndex        =   5
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":344E
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   4
      Left            =   1575
      TabIndex        =   6
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":3F40
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   5
      Left            =   1965
      TabIndex        =   7
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":42DA
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   6
      Left            =   2355
      TabIndex        =   8
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":4674
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   7
      Left            =   2745
      TabIndex        =   9
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":4A0E
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   8
      Left            =   3135
      TabIndex        =   10
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":5500
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   9
      Left            =   3525
      TabIndex        =   11
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":5FF2
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   10
      Left            =   3915
      TabIndex        =   12
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":638C
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   11
      Left            =   4305
      TabIndex        =   13
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":6726
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   12
      Left            =   4695
      TabIndex        =   14
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":6AC0
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   13
      Left            =   5085
      TabIndex        =   15
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":75B2
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   390
      Index           =   14
      Left            =   5475
      TabIndex        =   16
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Caption         =   ""
      BorderColor     =   -2147483648
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":80A4
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   13811126
      BackColorIconDown=   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   510
      Index           =   15
      Left            =   135
      TabIndex        =   17
      Top             =   795
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   900
      Caption         =   "Any Colour"
      BorderColor     =   255
      BorderColorOver =   65280
      Icon            =   "Form1.frx":843E
      BorderColorDown =   6956042
      BackColor       =   52224
      BackColorOver   =   65535
      BackColorDown   =   49152
      BackColorIcon   =   16761024
      BackColorIconOver=   255
      BackColorIconDown=   32768
      IconSizeWidth   =   24
      IconSizeHeight  =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
   End
   Begin PB_XP_Button.PBXPButton PBXPButton2 
      Height          =   435
      Index           =   1
      Left            =   150
      TabIndex        =   18
      Top             =   1875
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":8BB8
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Disabled        =   -1  'True
      ShowFocus       =   -1  'True
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   510
      Index           =   16
      Left            =   3570
      TabIndex        =   21
      Top             =   855
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   900
      BorderColor     =   15856113
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":8F52
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   14737632
      BackColorIconDown=   13811126
      IconSizeWidth   =   24
      IconSizeHeight  =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rondalo"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowSize      =   2
      FontColor       =   65280
      AlignCaption    =   2
      FontColorOver   =   12583104
      FontColorDown   =   16711680
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   510
      Index           =   17
      Left            =   3585
      TabIndex        =   22
      Top             =   1395
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   900
      BorderColor     =   15856113
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":96CC
      BorderColorDown =   6956042
      BackColor       =   -2147483648
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   14737632
      BackColorIconDown=   13811126
      IconSizeWidth   =   24
      IconSizeHeight  =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rondalo"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ShadowSize      =   1
      FontColor       =   65280
      ShowShadowOver  =   -1  'True
      FontColorOver   =   49152
      FontColorDown   =   32768
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   510
      Index           =   18
      Left            =   3585
      TabIndex        =   25
      Top             =   1935
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   900
      Caption         =   $"Form1.frx":9E46
      BorderColor     =   15856113
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":9E61
      BorderColorDown =   6956042
      BackColor       =   -2147483648
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   14737632
      BackColorIconDown=   13811126
      IconSizeWidth   =   24
      IconSizeHeight  =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
   End
   Begin PB_XP_Button.PBXPButton PBXPButton3 
      Height          =   540
      Index           =   19
      Left            =   3585
      TabIndex        =   26
      Top             =   2475
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   953
      Caption         =   $"Form1.frx":A5DB
      BorderColor     =   15856113
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":A5F6
      BorderColorDown =   6956042
      BackColor       =   -2147483648
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIcon   =   13752539
      BackColorIconOver=   14737632
      BackColorIconDown=   13811126
      IconSizeWidth   =   24
      IconSizeHeight  =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      AlignCaption    =   2
   End
   Begin PB_XP_Button.PBXPButton PBXPButton2 
      Height          =   435
      Index           =   2
      Left            =   150
      TabIndex        =   29
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":AD70
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
   End
   Begin PB_XP_Button.PBXPButton PBXPButton2 
      Height          =   420
      Index           =   3
      Left            =   1890
      TabIndex        =   30
      Top             =   1380
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   741
      Caption         =   "UnChecked"
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":B10A
      BorderColorDown =   6956042
      BackColor       =   -2147483644
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIconOver=   13811126
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checked         =   -1  'True
      CheckedColor    =   14211029
   End
   Begin PB_XP_Button.PBXPButton PBXPButton4 
      Height          =   390
      Index           =   1
      Left            =   1290
      TabIndex        =   33
      Top             =   3030
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   688
      Caption         =   "Options"
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":B4A4
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIconOver=   13811126
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checked         =   -1  'True
      CheckedColor    =   14211029
   End
   Begin PB_XP_Button.PBXPButton PBXPButton4 
      Height          =   390
      Index           =   2
      Left            =   2415
      TabIndex        =   34
      Top             =   3030
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   688
      Caption         =   "Options"
      BorderColorOver =   6956042
      Icon            =   "Form1.frx":B83E
      BorderColorDown =   6956042
      BackColor       =   13752539
      BackColorOver   =   13811126
      BackColorDown   =   11899525
      BackColorIconOver=   13811126
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checked         =   -1  'True
      CheckedColor    =   14211029
   End
   Begin VB.Label Label3 
      Height          =   225
      Left            =   2130
      TabIndex        =   31
      Top             =   1095
      Width           =   1035
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   27
      Top             =   3570
      Width           =   4740
   End
   Begin VB.Label lblShadowSize 
      Caption         =   "ShadowSize:"
      Height          =   270
      Left            =   4335
      TabIndex        =   24
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   240
      Left            =   90
      TabIndex        =   20
      Top             =   465
      Width           =   3105
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
PBXPButton2(1).Disabled = (Check1.Value = vbChecked)
End Sub

Private Sub Check2_Click()
PBXPButton2(2).ShowFocus = (Check2.Value = vbChecked)
End Sub

Private Sub PBXPButton1_Click()
Unload Me
End Sub

Private Sub PBXPButton2_Click(Index As Integer)
Select Case Index
Case 0, 3
Label3.Caption = PBXPButton2(Index).Value
PBXPButton2(1).Disabled = PBXPButton2(Index).Value
If PBXPButton2(Index).Value Then
PBXPButton2(Index).Caption = "Checked"
Else
PBXPButton2(Index).Caption = "UnChecked"
End If

End Select
End Sub

Private Sub PBXPButton2_MouseOut(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = ""
If Index = 3 Then Label3.Caption = PBXPButton2(3).Value
End Sub

Private Sub PBXPButton2_MouseOver(Index As Integer)
Select Case Index
Case 0
Label2.Caption = "CheckBox Button"
Case 1
Label2.Caption = "Disable control"
Case 2
Label2.Caption = "Show / Hide Focus Rect"
Case 3
Label2.Caption = "CheckBox Button"
End Select
End Sub

Private Sub PBXPButton3_MouseOut(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = ""
Label1.Caption = "Mouse is left button " & Index
End Sub

Private Sub PBXPButton3_MouseOver(Index As Integer)
Label1.Caption = "Mouse is over button " & Index
Select Case Index
Case 0 To 14
Label2.Caption = "Create a XP Toolbar"
Case 15
Label2.Caption = "Change any of the colors"
Case 16
Label2.Caption = "Add a shadow to the caption"
Case 17
Label2.Caption = "Show shadow on mouse over"
Case 18
Label2.Caption = "Supports Multi Line Caption"
Case 19
Label2.Caption = "Align the caption Left, Center or Right"
End Select
End Sub

Private Sub PBXPButton4_Click(Index As Integer)
Dim xx As Long
For xx = 0 To PBXPButton4.Count - 1
PBXPButton4(xx).Value = False
Next
PBXPButton4(Index).Value = True
End Sub

Private Sub PBXPButton4_MouseOut(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = ""
End Sub

Private Sub PBXPButton4_MouseOver(Index As Integer)
Label2.Caption = "Option Button " & Index
End Sub

Private Sub Text1_Change()
On Error Resume Next
PBXPButton3(16).ShadowSize = Text1.Text
End Sub
