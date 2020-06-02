VERSION 5.00
Begin VB.UserControl VPProgress 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DrawMode        =   6  'Mask Pen Not
   PropertyPages   =   "VPProgress.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "VPProgress.ctx":0023
   Begin VB.PictureBox picDisplay 
      BorderStyle     =   0  'None
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H80000002&
      Height          =   510
      Left            =   180
      ScaleHeight     =   510
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   2970
      Width           =   4425
      Begin VB.Line linRightBorder 
         Visible         =   0   'False
         X1              =   1620
         X2              =   1620
         Y1              =   1080
         Y2              =   2835
      End
      Begin VB.Line linMiddleBorder 
         Visible         =   0   'False
         X1              =   540
         X2              =   540
         Y1              =   1035
         Y2              =   2745
      End
      Begin VB.Line linLeftBorder 
         Visible         =   0   'False
         X1              =   225
         X2              =   225
         Y1              =   1035
         Y2              =   2700
      End
      Begin VB.Line linMiddle 
         X1              =   45
         X2              =   4680
         Y1              =   225
         Y2              =   225
      End
      Begin VB.Line linTop 
         X1              =   0
         X2              =   4725
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linBottom 
         X1              =   225
         X2              =   4905
         Y1              =   630
         Y2              =   630
      End
   End
End
Attribute VB_Name = "VPProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Const m_def_ShadowColor = vbButtonShadow
Const m_def_HighlightColor = vb3DHighlight
Const m_def_BorderStyle = 0

Dim m_ShadowColor As OLE_COLOR
Dim m_HighlightColor As OLE_COLOR
Dim m_Font As Font
Dim m_BorderStyle As Integer
Dim msMessage As String

Public Sub DisplayMessage(asMessage As String)
  'display text message asMessage
  'in picturebox picDisplay
  'if asMessage="" - clear display
  
  msMessage = asMessage
  
  picDisplay.Cls
  
  If asMessage <> "" Then
    picDisplay.AutoRedraw = True
    picDisplay.CurrentX = (picDisplay.ScaleWidth - picDisplay.TextWidth(asMessage)) / 2
    picDisplay.CurrentY = (picDisplay.ScaleHeight - picDisplay.TextHeight(asMessage)) / 2
    picDisplay.Print asMessage
    picDisplay.Refresh
  End If
  
End Sub

Public Sub DisplayProgress(alMax As Long, alCurrent As Long, Optional asMessage As String)
  'display progress in loading alMax record
  'alCurrent is number of currently loading record
  'asMessage - message to display before %
  'picDisplay - picture where progress is displayed
  
  Dim liPercent As Integer
  Dim lsMessage As String
  
  If alCurrent = 0 Then
    Cls
    picDisplay.Refresh
    Exit Sub
  End If
  
  liPercent = CInt(100 * alCurrent / alMax)
  picDisplay.Cls
  lsMessage = asMessage & " " & CStr(liPercent) & "%"
  'center and print the message
  picDisplay.AutoRedraw = True
  picDisplay.CurrentX = (picDisplay.ScaleWidth - picDisplay.TextWidth(lsMessage)) / 2
  picDisplay.CurrentY = (picDisplay.ScaleHeight - picDisplay.TextHeight(lsMessage)) / 2
  picDisplay.Print lsMessage
  
  'show progress bar
  picDisplay.Line (0, 0)-(liPercent * picDisplay.ScaleWidth / 100, picDisplay.ScaleHeight), vbHighlight, BF
  picDisplay.Refresh
End Sub


Private Sub Refresh()
   Dim liDelta As Integer

  liDelta = 30

  'set colors
  linTop.BorderColor = Me.ShadowColor
  linMiddle.BorderColor = Me.HighlightColor
  linBottom.BorderColor = Me.ShadowColor
  linLeftBorder.BorderColor = Me.ShadowColor
  linMiddleBorder.BorderColor = Me.HighlightColor
  linRightBorder.BorderColor = Me.ShadowColor
  linLeftBorder.Visible = BorderStyle
  linMiddleBorder.Visible = BorderStyle
  linRightBorder.Visible = BorderStyle
  linBottom.Visible = BorderStyle
  
  picDisplay.Move 0, 0, ScaleWidth, ScaleHeight
  
  'move lines
  With linTop
    .X1 = 0
    .X2 = picDisplay.ScaleWidth
    .Y1 = 0
    .Y2 = 0
    linMiddle.X1 = .X1
    linMiddle.X2 = .X2
    linMiddle.Y1 = .Y1 + 10
    linMiddle.Y2 = .Y2 + 10
    linBottom.X1 = .X1
    linBottom.X2 = .X2
    linBottom.Y1 = picDisplay.ScaleHeight - 10
    linBottom.Y2 = linBottom.Y1
  End With
  With linLeftBorder
    .X1 = 0
    .X2 = 0
    .Y1 = 10
    .Y2 = picDisplay.ScaleHeight - 10
  End With
  With linMiddleBorder
    .X1 = 10
    .X2 = 10
    .Y1 = 10
    .Y2 = picDisplay.ScaleHeight - 25
  End With
  With linRightBorder
    .X1 = picDisplay.ScaleWidth - 15
    .X2 = .X1
    .Y1 = 15
    .Y2 = picDisplay.ScaleHeight - 10
  End With
  
  Call DisplayMessage(msMessage)
  
End Sub


Public Property Get ShadowColor() As OLE_COLOR
  ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
  m_ShadowColor = New_ShadowColor
  PropertyChanged "ShadowColor"
End Property

Public Property Get HighlightColor() As OLE_COLOR
  HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal New_HighlightColor As OLE_COLOR)
  m_HighlightColor = New_HighlightColor
  PropertyChanged "HighlightColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ShadowColor = m_def_ShadowColor
  m_HighlightColor = m_def_HighlightColor
  Set m_Font = Ambient.Font
  m_BorderStyle = m_def_BorderStyle
End Sub

Private Sub UserControl_Paint()
  Call Refresh
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_ShadowColor = PropBag.ReadProperty("ShadowColor", vbButtonShadow)
  m_HighlightColor = PropBag.ReadProperty("HighlightColor", vb3DHighlight)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
  m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
  picDisplay.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
  picDisplay.FillColor = PropBag.ReadProperty("FillColor", vbActiveTitleBar)
End Sub

Private Sub UserControl_Resize()
  Call Refresh
End Sub

Private Sub UserControl_Terminate()
  Set m_Font = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("ShadowColor", m_ShadowColor, vbButtonShadow)
  Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, vb3DHighlight)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("ForeColor", picDisplay.ForeColor, vbButtonText)
  Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
  Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
  Call PropBag.WriteProperty("FillColor", picDisplay.FillColor, vbActiveTitleBar)
End Sub
'

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
  Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set m_Font = New_Font
  PropertyChanged "Font"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  m_BorderStyle = New_BorderStyle
  Call Refresh
  PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = picDisplay.BackColor
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = picDisplay.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  picDisplay.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property
