VERSION 5.00
Object = "{F431B48E-43DD-4783-9C34-EED68792E9D5}#2.1#0"; "MultiButtonControl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MultiButton Test Harness"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MultiButtonControl.MultiButton btnSearch 
      Height          =   555
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":0000
      BorderColor     =   -2147483633
      Caption         =   "Search"
      RedrawOnHover   =   0   'False
      Value           =   -1  'True
      CheckedBorderColor=   -2147483647
      CheckedFillColor=   -2147483628
      CheckedForeColor=   -2147483647
      ButtonMode      =   1
      OptionName      =   "view"
      CornerRadius    =   5
   End
   Begin MultiButtonControl.MultiButton pnlInfo 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1140
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":1D0A
      FillColor       =   -2147483624
      RedrawOnHover   =   0   'False
      Alignment       =   0
      BackColor       =   -2147483624
   End
   Begin MultiButtonControl.MultiButton btnCancel 
      Height          =   375
      Left            =   6420
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   4440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Picture         =   "Test.frx":215C
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
   Begin MultiButtonControl.MultiButton btnContainer 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      BorderColor     =   -2147483647
      Caption         =   "It's also a container control "
      RedrawOnHover   =   0   'False
      Alignment       =   1
      Begin MultiButtonControl.MultiButton btn 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   -2147483633
         Caption         =   "&File"
         HoverFillColor  =   14073525
         HoverBorderColor=   -2147483635
         ActiveForeColor =   -2147483647
         ActiveFillColor =   -2147483645
         ButtonMode      =   3
      End
      Begin MultiButtonControl.MultiButton btn 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   -2147483633
         Caption         =   "&Edit"
         HoverFillColor  =   14073525
         HoverBorderColor=   -2147483635
         ActiveForeColor =   -2147483647
         ActiveFillColor =   -2147483645
         ButtonMode      =   3
      End
      Begin MultiButtonControl.MultiButton btn 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Click Me"
         Top             =   30
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Test.frx":26F6
         BorderColor     =   -2147483633
         Caption         =   "&Help"
         HoverFillColor  =   14073525
         HoverBorderColor=   -2147483635
         Alignment       =   1
         PictureAlignment=   1
      End
   End
   Begin MultiButtonControl.MultiButton btnOk 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      ToolTipText     =   "Ok"
      Top             =   4440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Picture         =   "Test.frx":2850
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
   Begin MultiButtonControl.MultiButton btnFolders 
      Height          =   555
      Left            =   1380
      TabIndex        =   8
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":29AA
      BorderColor     =   -2147483633
      Caption         =   "Folders"
      RedrawOnHover   =   0   'False
      CheckedBorderColor=   -2147483647
      CheckedFillColor=   -2147483628
      CheckedForeColor=   -2147483647
      ButtonMode      =   1
      OptionName      =   "view"
      CheckedPicture  =   "Test.frx":533C
      CornerRadius    =   5
   End
   Begin MultiButtonControl.MultiButton MultiButton2 
      Height          =   1215
      Left            =   4020
      TabIndex        =   9
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Caption         =   "Every control on this form is a MultiButton! (Including the menu's)"
      RedrawOnHover   =   0   'False
   End
   Begin MultiButtonControl.MultiButton MultiButton1 
      Height          =   1515
      Left            =   4020
      TabIndex        =   10
      Top             =   2820
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2672
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BorderColor     =   -2147483635
      Caption         =   "Its a Group Box"
      RedrawOnHover   =   0   'False
      Alignment       =   0
      VerticalAlignment=   0
      ButtonMode      =   2
      CornerRadius    =   5
      Begin MultiButtonControl.MultiButton jslPanel3 
         Height          =   1035
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         FillColor       =   -2147483624
         Caption         =   "When ButtonMode is in mbnOption and no value is specified in the OptionName then the control acts like a checkbox"
         RedrawOnHover   =   0   'False
         Alignment       =   0
      End
   End
   Begin MultiButtonControl.MultiButton opt 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   2460
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":7CCE
      Caption         =   "Black Text"
      HoverForeColor  =   16711680
      Alignment       =   0
      Value           =   -1  'True
      CheckedBorderColor=   -2147483635
      CheckedFillColor=   -2147483624
      ButtonMode      =   1
      OptionName      =   "fc"
      CheckedPicture  =   "Test.frx":8268
      CornerRadius    =   20
   End
   Begin MultiButtonControl.MultiButton Info 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":8802
      BorderColor     =   -2147483635
      FillColor       =   -2147483624
      Caption         =   "Click ""Set Text"""
      RedrawOnHover   =   0   'False
      HoverBorderColor=   16777215
      Alignment       =   0
      PictureAlignment=   1
      ActiveBorderColor=   -2147483624
   End
   Begin MultiButtonControl.MultiButton btnSetText 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Change text"
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      MousePointer    =   99
      Caption         =   "Set Text"
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      PictureAlignment=   2
      ActiveFillColor =   14073525
      MouseIcon       =   "Test.frx":90DC
   End
   Begin MultiButtonControl.MultiButton btnAlignText 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Left Justify"
      Top             =   2460
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Picture         =   "Test.frx":923E
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      ActiveFillColor =   15263457
      Value           =   -1  'True
      CheckedBorderColor=   -2147483647
      CheckedFillColor=   15263457
      ButtonMode      =   1
      OptionName      =   "Justify"
   End
   Begin MultiButtonControl.MultiButton btnAlignText 
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   16
      ToolTipText     =   "Center Justify"
      Top             =   2460
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Picture         =   "Test.frx":9398
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      ActiveFillColor =   15263457
      CheckedBorderColor=   -2147483647
      CheckedFillColor=   15263457
      ButtonMode      =   1
      OptionName      =   "Justify"
   End
   Begin MultiButtonControl.MultiButton btnAlignText 
      Height          =   315
      Index           =   2
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "Right Justify"
      Top             =   2460
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Picture         =   "Test.frx":94F2
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      ActiveFillColor =   15263457
      CheckedBorderColor=   -2147483647
      CheckedFillColor=   15263457
      ButtonMode      =   1
      OptionName      =   "Justify"
   End
   Begin MultiButtonControl.MultiButton opt 
      Height          =   315
      Index           =   1
      Left            =   2580
      TabIndex        =   18
      Top             =   2460
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":964C
      Caption         =   "RedText"
      HoverForeColor  =   16711680
      Alignment       =   0
      CheckedBorderColor=   -2147483635
      CheckedFillColor=   -2147483624
      ButtonMode      =   1
      OptionName      =   "fc"
      CheckedPicture  =   "Test.frx":9BE6
      CornerRadius    =   20
   End
   Begin MultiButtonControl.MultiButton optValueA 
      Height          =   255
      Left            =   4020
      TabIndex        =   19
      Top             =   2100
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":A180
      BorderColor     =   -2147483633
      Caption         =   "Checkbox A"
      HoverForeColor  =   16711680
      HoverBorderColor=   33023
      Alignment       =   0
      VerticalAlignment=   2
      CheckedBorderColor=   -2147483633
      ButtonMode      =   1
      CheckedPicture  =   "Test.frx":A2DA
   End
   Begin MultiButtonControl.MultiButton optValueB 
      Height          =   255
      Left            =   5460
      TabIndex        =   20
      Top             =   2100
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "Test.frx":A434
      BorderColor     =   -2147483633
      Caption         =   "Checkbox B"
      HoverForeColor  =   16711680
      HoverBorderColor=   33023
      Alignment       =   0
      VerticalAlignment=   2
      CheckedBorderColor=   -2147483633
      ButtonMode      =   1
      CheckedPicture  =   "Test.frx":A58E
   End
   Begin VB.Menu FileMenu 
      Caption         =   "F"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "E"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnAlignText_Click(Index As Integer)
    Select Case Index
        Case 0
            Info.Alignment = vbLeftJustify
        Case 1
            Info.Alignment = vbCenter
        Case 2
            Info.Alignment = vbRightJustify
    End Select
End Sub

Private Sub btnSetText_Click()
    Dim sMsg As String
    
    sMsg = "MultiButtons that do not have text will not show as pressed when clicked." & vbCrLf & "Group relationships can be made by setting the 'OptionName' property. This gets over using a container control." & vbCrLf & "Enjoy..."
    Info.Caption = sMsg
End Sub

Private Sub Form_Load()
    Dim sMsg As String
    
    Show
    
    sMsg = "The MultiButton control is a multi function control that can emulate the following controls: Button, ToolButton, OptionButton, CheckBox, Frame and Menu.  This control can also be made to look like MS Explorer buttons."
    pnlInfo.Caption = sMsg
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0
            PopupMenu FileMenu, , btn(Index).Left, btn(Index).Top + btn(Index).Height
        Case 1
            PopupMenu EditMenu, , btn(Index).Left, btn(Index).Top + btn(Index).Height
        Case 2
            MsgBox "MultiButton Control by Paul Sanders", vbInformation
    End Select
End Sub


Private Sub btnOk_Click()
    Dim f As Form2
    
    Set f = New Form2
    f.Show vbModal
    Set f = Nothing
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub opt_Click(Index As Integer)
    If Index = 0 Then
        Info.ForeColor = vbWindowText
    Else
        Info.ForeColor = vbRed
    End If
End Sub
