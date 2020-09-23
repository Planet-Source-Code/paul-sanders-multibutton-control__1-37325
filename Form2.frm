VERSION 5.00
Object = "{F431B48E-43DD-4783-9C34-EED68792E9D5}#2.1#0"; "MultiButtonControl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EditBar Example"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MultiButtonControl.MultiButton MultiButton1 
      Height          =   915
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1614
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
      Picture         =   "Form2.frx":0000
      BorderColor     =   -2147483647
      FillColor       =   65535
      Caption         =   "The EditBar allows a common GUI to manage data handling."
      RedrawOnHover   =   0   'False
      CornerRadius    =   20
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1575
   End
   Begin MultiButtonControl.EditBar MultiEditBar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   767
      BarBorderColor  =   -2147483647
   End
   Begin MultiButtonControl.MultiButton btnOk 
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      ToolTipText     =   "Ok"
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
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
      Picture         =   "Form2.frx":031A
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MultiEditBar1_Cancel()
    Text1.Text = "Cancel"
    pEnable False
End Sub

Private Sub MultiEditBar1_DeleteItem()
    Text1.Text = "Delete"
End Sub

Private Sub MultiEditBar1_EditItem(Cancel As Boolean)
    Text1.Text = "Edit item"
    Cancel = MsgBox("Edit?", vbQuestion + vbYesNo) = vbNo
    pEnable Cancel = False
End Sub


Private Sub MultiEditBar1_NewItem(Cancel As Boolean)
    Text1.Text = "New"
    Cancel = MsgBox("New?", vbQuestion + vbYesNo) = vbNo
    pEnable Cancel = False
End Sub

Private Sub MultiEditBar1_Ok(Cancel As Boolean)
    Text1.Text = "Ok"
    pEnable False
End Sub

Private Sub pEnable(b As Boolean)
    Text1.Enabled = b
    btnOk.Enabled = b
End Sub
