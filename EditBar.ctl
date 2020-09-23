VERSION 5.00
Begin VB.UserControl EditBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   57
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "EditBar.ctx":0000
   Begin MultiButtonControl.MultiButton cmd 
      Height          =   375
      Index           =   4
      Left            =   60
      TabIndex        =   4
      ToolTipText     =   "Ok"
      Top             =   30
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
      Picture         =   "EditBar.ctx":0532
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
   Begin MultiButtonControl.MultiButton cmd 
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Cancel"
      Top             =   30
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
      Picture         =   "EditBar.ctx":068C
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
   Begin MultiButtonControl.MultiButton cmd 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "New"
      Top             =   30
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
      Picture         =   "EditBar.ctx":0C26
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
   Begin MultiButtonControl.MultiButton cmd 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Edit"
      Top             =   30
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
      Picture         =   "EditBar.ctx":0D80
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
   Begin MultiButtonControl.MultiButton cmd 
      Height          =   375
      Index           =   3
      Left            =   900
      TabIndex        =   3
      ToolTipText     =   "Delete"
      Top             =   30
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
      Picture         =   "EditBar.ctx":0EDA
      Caption         =   ""
      HoverFillColor  =   14073525
      HoverBorderColor=   -2147483635
      Alignment       =   0
      PictureAlignment=   2
      ActiveFillColor =   14073525
   End
End
Attribute VB_Name = "EditBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-----------------------------------------'
' By Paul Sanders, pa_sanders@hotmail.com '
'--------------------------------------------------------------------------------------------
'            :
' Project    : MultiButtonControl
' Module     : EditBar
'            :
' Created    : 07-Jul-02 19:14
'            :
' Notes      :
'            :
' References : None
'            :
'--------------------------------------------------------------------------------------------

Private Const MODULENAME = "MultiEditBar::"

Private Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long)
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Event NewItem(ByRef Cancel As Boolean)
Event EditItem(ByRef Cancel As Boolean)
Event DeleteItem()
Event Ok(ByRef Cancel As Boolean)
Event Cancel()

'Default Property Values:
Const m_def_AllowDelete = True

'Property Variables:
Dim m_AllowDelete As Boolean
Dim m_BarBorderColor As OLE_COLOR

Const cNew = 0
Const cEdit = 1
Const cCancel = 2
Const cDelete = 3
Const cOk = 4
Const cMAX = 4

'--------------------------------------------------------------------------------------------
'Procedure : StartEdit
'Author    : Paul Sanders, pa_sanders@hotmail.com, 07-Jul-02 21:11
'Notes     : Starts the edit process
'--------------------------------------------------------------------------------------------
Public Sub StartEdit()
    Dim B As Boolean
    
    RaiseEvent EditItem(B)
    If B = False Then
        pEdit True
    End If
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : StopEdit
'Author    : Paul Sanders, pa_sanders@hotmail.com, 07-Jul-02 21:13
'Notes     : Stops any editing and resets controls
'--------------------------------------------------------------------------------------------
Public Sub StopEdit()
    pEdit False
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : pEdit
'Author    : Paul Sanders, pa_sanders@hotmail.com, 07-Jul-02 21:14
'Notes     : Sets up the controls to display correctly depending on current state
'--------------------------------------------------------------------------------------------
Private Sub pEdit(B As Boolean)
    Dim i As Integer
    
    For i = 0 To cMAX
        If i = cDelete And m_AllowDelete = False Then
            cmd(i).Visible = False
            cmd(i).Enabled = False
        Else
            If i = cNew Or i = cEdit Or i = cDelete Then
                cmd(i).Enabled = Not B
                cmd(i).Visible = Not B
            Else
                cmd(i).Visible = B
                cmd(i).Enabled = B
            End If
        End If
    Next
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Dim i As Integer
    
    For i = cmd.LBound To cmd.UBound
        cmd(i).Enabled = New_Enabled
    Next
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub cmd_Click(Index As Integer)
    Dim bCancel As Boolean
    
    Select Case Index
        Case cNew
            RaiseEvent NewItem(bCancel)
            If Not bCancel Then
                pEdit True
            End If
        
        Case cEdit
            StartEdit
        
        Case cDelete
            RaiseEvent DeleteItem
        
        Case cOk
            RaiseEvent Ok(bCancel)
            If Not bCancel Then
                pEdit False
            End If
            
        Case cCancel
            RaiseEvent Cancel
            pEdit False
    End Select
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AllowDelete = m_def_AllowDelete
    m_BarBorderColor = VBRUN.SystemColorConstants.vb3DDKShadow
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    ActiveBorderColor = PropBag.ReadProperty("ActiveBorderColor", &HFFFFFF)
    ActiveFillColor = PropBag.ReadProperty("ActiveFillColor", &HD6BEB5)
    ActiveForeColor = PropBag.ReadProperty("ActiveForeColor", &H80000008)
    BorderColor = PropBag.ReadProperty("BorderColor", &H80000010)
    HoverBorderColor = PropBag.ReadProperty("HoverBorderColor", &H8000000D)
    HoverFillColor = PropBag.ReadProperty("HoverFillColor", &HD6BEB5)
    HoverForeColor = PropBag.ReadProperty("HoverForeColor", &H80000008)
    m_AllowDelete = PropBag.ReadProperty("AllowDelete", m_def_AllowDelete)
    m_BarBorderColor = PropBag.ReadProperty("BarBorderColor", vb3DDKShadow)

    Set cmd(cNew).Picture = PropBag.ReadProperty("NewPicture", cmd(cNew).Picture)
    Set cmd(cEdit).Picture = PropBag.ReadProperty("EditPicture", cmd(cEdit).Picture)
    Set cmd(cCancel).Picture = PropBag.ReadProperty("CancelPicture", cmd(cCancel).Picture)
    Set cmd(cDelete).Picture = PropBag.ReadProperty("DeletePicture", cmd(cDelete).Picture)
    Set cmd(cOk).Picture = PropBag.ReadProperty("OkPicture", cmd(cOk).Picture)
End Sub

Private Sub UserControl_Resize()
    pDraw
End Sub

Private Sub UserControl_Show()
    pDraw
    pEdit False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ActiveBorderColor", cmd(0).ActiveBorderColor, &HFFFFFF)
    Call PropBag.WriteProperty("ActiveFillColor", cmd(0).ActiveFillColor, &HD6BEB5)
    Call PropBag.WriteProperty("ActiveForeColor", cmd(0).ActiveForeColor, &H80000008)
    Call PropBag.WriteProperty("BorderColor", cmd(0).BorderColor, &H80000010)
    Call PropBag.WriteProperty("HoverBorderColor", cmd(0).HoverBorderColor, &H8000000D)
    Call PropBag.WriteProperty("HoverFillColor", cmd(0).HoverFillColor, &HD6BEB5)
    Call PropBag.WriteProperty("HoverForeColor", cmd(0).HoverForeColor, &H80000008)
    Call PropBag.WriteProperty("AllowDelete", m_AllowDelete, m_def_AllowDelete)
    Call PropBag.WriteProperty("BarBorderColor", m_BarBorderColor, vb3DDKShadow)
    
    Call PropBag.WriteProperty("NewPicture", cmd(cNew).Picture, cmd(cNew).Picture)
    Call PropBag.WriteProperty("EditPicture", cmd(cEdit).Picture, cmd(cEdit).Picture)
    Call PropBag.WriteProperty("CancelPicture", cmd(cCancel).Picture, cmd(cCancel).Picture)
    Call PropBag.WriteProperty("DeletePicture", cmd(cDelete).Picture, cmd(cDelete).Picture)
    Call PropBag.WriteProperty("OkPicture", cmd(cOk).Picture, cmd(cOk).Picture)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,ActiveBorderColor
Public Property Get ActiveBorderColor() As OLE_COLOR
    ActiveBorderColor = cmd(0).ActiveBorderColor
End Property

Public Property Let ActiveBorderColor(ByVal New_ActiveBorderColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).ActiveBorderColor() = New_ActiveBorderColor
    Next
    PropertyChanged "ActiveBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,ActiveFillColor
Public Property Get ActiveFillColor() As OLE_COLOR
    ActiveFillColor = cmd(0).ActiveFillColor
End Property

Public Property Let ActiveFillColor(ByVal New_ActiveFillColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).ActiveFillColor() = New_ActiveFillColor
    Next
    PropertyChanged "ActiveFillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,ActiveForeColor
Public Property Get ActiveForeColor() As OLE_COLOR
    ActiveForeColor = cmd(0).ActiveForeColor
End Property

Public Property Let ActiveForeColor(ByVal New_ActiveForeColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).ActiveForeColor() = New_ActiveForeColor
    Next
    
    PropertyChanged "ActiveForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,BorderColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = cmd(0).BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).BorderColor() = New_BorderColor
    Next
    PropertyChanged "BorderColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,HoverBorderColor
Public Property Get HoverBorderColor() As OLE_COLOR
    HoverBorderColor = cmd(0).HoverBorderColor
End Property

Public Property Let HoverBorderColor(ByVal New_HoverBorderColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).HoverBorderColor() = New_HoverBorderColor
    Next
    PropertyChanged "HoverBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,HoverFillColor
Public Property Get HoverFillColor() As OLE_COLOR
    HoverFillColor = cmd(0).HoverFillColor
End Property

Public Property Let HoverFillColor(ByVal New_HoverFillColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).HoverFillColor() = New_HoverFillColor
    Next
    PropertyChanged "HoverFillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmd(0),cmd,0,HoverForeColor
Public Property Get HoverForeColor() As OLE_COLOR
    HoverForeColor = cmd(0).HoverForeColor
End Property

Public Property Let HoverForeColor(ByVal New_HoverForeColor As OLE_COLOR)
    Dim i As Integer
    
    For i = 0 To cMAX
        cmd(i).HoverForeColor() = New_HoverForeColor
    Next
    PropertyChanged "HoverForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AllowDelete() As Boolean
    AllowDelete = m_AllowDelete
End Property

Public Property Let AllowDelete(ByVal New_AllowDelete As Boolean)
    m_AllowDelete = New_AllowDelete
    If m_AllowDelete Then
        cmd(cDelete).Enabled = cmd(cEdit).Enabled
        cmd(cDelete).Visible = cmd(cEdit).Visible
    Else
        cmd(cDelete).Enabled = False
        cmd(cDelete).Visible = False
    End If
    PropertyChanged "AllowDelete"
End Property

Public Property Let BarBorderColor(NV As OLE_COLOR)
Attribute BarBorderColor.VB_Description = "Returns/sets the border color of the editbar."
    m_BarBorderColor = NV
    pDraw
End Property
Public Property Get BarBorderColor() As OLE_COLOR
    BarBorderColor = m_BarBorderColor
End Property

Private Sub pDraw()
    Dim hPen As Long
    Dim hPenOld As Long
    Dim x As Long, y As Long, w As Long, h As Long
    Dim PT As POINTAPI
    
    x = 0: y = 0
    w = ScaleWidth
    h = ScaleHeight
    
    hPen = CreatePen(0, 1, TranslateColor(m_BarBorderColor))
    hPenOld = SelectObject(hDC, hPen)

    Cls
    
    'Draw a standard box
    MoveToEx hDC, x + w - 1, y, PT
    LineTo hDC, x, y
    LineTo hDC, x, y + h - 1
    LineTo hDC, x + w - 1, y + h - 1
    LineTo hDC, x + w - 1, y
    
    'Clean up
    SelectObject hDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
    
    Refresh
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : TranslateColor
'Author    : Paul Sanders, pa_sanders@hotmail.com, 03-Apr-02 00:08
'Notes     :
'--------------------------------------------------------------------------------------------
Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function

Public Property Set NewPicture(NV As StdPicture)
    Set cmd(cNew).Picture = NV
    PropertyChanged "NewPicture"
End Property
Public Property Get NewPicture() As StdPicture
    Set NewPicture = cmd(cNew).Picture
End Property

Public Property Set EditPicture(NV As StdPicture)
    Set cmd(cEdit).Picture = NV
    PropertyChanged "EditPicture"
End Property
Public Property Get EditPicture() As StdPicture
    Set EditPicture = cmd(cEdit).Picture
End Property

Public Property Set CancelPicture(NV As StdPicture)
    Set cmd(cCancel).Picture = NV
    PropertyChanged "CancelPicture"
End Property
Public Property Get CancelPicture() As StdPicture
    Set CancelPicture = cmd(cCancel).Picture
End Property

Public Property Set DeletePicture(NV As StdPicture)
    Set cmd(cDelete).Picture = NV
    PropertyChanged "DeletePicture"
End Property
Public Property Get DeletePicture() As StdPicture
    Set DeletePicture = cmd(cDelete).Picture
End Property

Public Property Set OkPicture(NV As StdPicture)
    Set cmd(cOk).Picture = NV
    PropertyChanged "OkPicture"
End Property
Public Property Get OkPicture() As StdPicture
    Set OkPicture = cmd(cOk).Picture
End Property

