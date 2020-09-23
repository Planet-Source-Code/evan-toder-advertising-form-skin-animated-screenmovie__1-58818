VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl AdvertisingSkin 
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   ScaleHeight     =   2190
   ScaleWidth      =   4350
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   6300
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "AdvertisingSkin.ctx":0000
      Top             =   4275
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "AdvertisingSkin.ctx":017B
      Top             =   2790
      Visible         =   0   'False
      Width           =   825
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   1725
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   3795
      ExtentX         =   6694
      ExtentY         =   3043
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   5850
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "AdvertisingSkin.ctx":05EA
      Top             =   1260
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "AdvertisingSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long

 
Private WithEvents F As Form
Attribute F.VB_VarHelpID = -1

'Default Property Values:
Const m_def_FormOnTop = 0
Const m_def_Message = "your message "
Const m_def_BackColor = 0

'Property Variables:
Dim m_FormOnTop As Boolean
Dim m_Message As String
Dim m_BackColor As OLE_COLOR
 

Private Function SetFormZorder()
  
  On Error Resume Next
  
  If m_FormOnTop Then
     SetWindowPos F.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
  Else
     SetWindowPos F.hwnd, 1, 0, 0, 0, 0, &H1 Or &H2
  End If
  
End Function
'========================================================================================
Function ColorToHTML(ByVal color As Long) As String
    ' HTML color codes are in the format #RRGGBB (red, green, blue)
    ' while Hex(color) returns numbers in the format BBGGRR
    ' therefore we just have to invert the order of the
    ' hex values of red and blue
    Dim tmp As String
    tmp = Right$("00000" & Hex$(color), 6)
    ColorToHTML = "#" & Right$(tmp, 2) & Mid$(tmp, 3, 2) & Left$(tmp, 2)
End Function
'========================================================================================
' convert a VB color constant to a COLORREF
' accepts both RGB() values and system color constants
Function TranslateColor(ByVal clr As Long) As Long
    If OleTranslateColor(clr, 0, TranslateColor) Then
         TranslateColor = -1
    End If
End Function
'========================================================================================
Private Function htmlMessage() As String

  htmlMessage = "var message = '" & m_Message & " '" & vbCrLf

End Function
'========================================================================================
Private Function htmlCode1() As String
  
  Dim strcolor As String
  
  strcolor = ColorToHTML(TranslateColor(m_BackColor))

  htmlCode1 = "<body bgcolor='" & strcolor & _
              "' oncontextmenu='return false' scroll='no' onLoad='makesnake()'" & _
              "style='width:100%;overflow-x:hidden;overflow-y:scroll'>" & vbCrLf

End Function
'========================================================================================
Private Sub F_Resize()

  UserControl.Extender.Move 0, 0, F.Width + 100, F.Height
 
End Sub
'========================================================================================
Private Sub UserControl_Resize()
 
 On Error Resume Next
 WB1.Move -50, -50, (Width + 250), (Height + 100)
 UserControl.Extender.Move 0, 0, UserControl.Parent.Width, UserControl.Parent.Height
 
End Sub
'========================================================================================
Private Sub UserControl_Show()
   
   'lets make sure this controls parent is indeed the form
   If IsWindow(UserControl.Parent.hwnd) Then
       WB1.Navigate "about:blank"
       DoEvents
   End If
   
End Sub
'========================================================================================
Private Sub WB1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
 
 DoEvents
 WB1.Document.write Text1 & vbCrLf & htmlMessage & _
                    Text2 & vbCrLf & htmlCode1 & Text3
 WB1.Refresh
 
End Sub
'========================================================================================
Private Sub UserControl_Terminate()
  
  Set F = Nothing
  
End Sub
'BackColor==============================================================================
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Call UserControl_Show
End Property
'FormOnTop
Public Property Get FormOnTop() As Boolean
    FormOnTop = m_FormOnTop
End Property
Public Property Let FormOnTop(ByVal New_FormOnTop As Boolean)
    m_FormOnTop = New_FormOnTop
    PropertyChanged "FormOnTop"
    Call SetFormZorder
End Property
'Message=================================================================================
Public Property Get Message() As String
    Message = m_Message
End Property
Public Property Let Message(ByVal New_Message As String)
    m_Message = New_Message
    PropertyChanged "Message"
End Property
'========================================================================================
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_Message = m_def_Message
    m_FormOnTop = m_def_FormOnTop
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Message = PropBag.ReadProperty("Message", m_def_Message)
    m_FormOnTop = PropBag.ReadProperty("FormOnTop", m_def_FormOnTop)
    
    
    If Ambient.UserMode Then
      If IsWindow(UserControl.Parent.hwnd) Then
        Set F = UserControl.Extender.Parent
        Call SetFormZorder
      End If
    Else
      Set F = Nothing
    End If
    
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Message", m_Message, m_def_Message)
    Call PropBag.WriteProperty("FormOnTop", m_FormOnTop, m_def_FormOnTop)
End Sub
 


