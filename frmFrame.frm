VERSION 5.00
Object = "*\AMultiButtonControl.vbp"
Begin VB.Form frmFrame 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MultiButtonControl.MultiButton MultiButton2 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1620
      Width           =   735
      _ExtentX        =   1296
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
      ForeColor       =   0
   End
   Begin MultiButtonControl.MultiButton MultiButton1 
      Height          =   1275
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   2249
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
      Caption         =   "Bob the Builder"
      RedrawOnHover   =   0   'False
      Alignment       =   0
      VerticalAlignment=   0
      ButtonMode      =   2
      CornerRadius    =   8
      Begin MultiButtonControl.MultiButton btnOpt 
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "frmFrame.frx":0000
         Caption         =   "Wendy "
         Alignment       =   1
         CheckedBorderColor=   -2147483645
         ButtonMode      =   1
         OptionName      =   "Bob"
         CheckedPicture  =   "frmFrame.frx":059A
         CornerRadius    =   12
      End
      Begin MultiButtonControl.MultiButton btnOpt 
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "frmFrame.frx":0B34
         Caption         =   "Bob "
         Alignment       =   1
         Value           =   -1  'True
         CheckedBorderColor=   -2147483645
         ButtonMode      =   1
         OptionName      =   "Bob"
         CheckedPicture  =   "frmFrame.frx":10CE
         CornerRadius    =   12
      End
   End
End
Attribute VB_Name = "frmFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
