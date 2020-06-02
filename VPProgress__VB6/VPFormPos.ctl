VERSION 5.00
Begin VB.UserControl VPFormPos 
   BackColor       =   &H8000000D&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VPFormPos.ctx":0000
   PropertyPages   =   "VPFormPos.ctx":00FA
   ScaleHeight     =   240
   ScaleWidth      =   240
   ToolboxBitmap   =   "VPFormPos.ctx":0107
End
Attribute VB_Name = "VPFormPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'// Our parent form
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1

'// Parent Form's Width and Height
Private m_sngPWidth As Single
Private m_sngPHeight As Single

'Default Property Values:
Const m_def_CenterForm = True
Const m_def_RememberFormPosition = True
Const m_def_MinHeight = 3000
Const m_def_MinWidth = 3000
'Property Variables:
Dim m_CenterForm As Boolean
Dim m_RememberFormPosition As Boolean

Dim m_MinHeight As Long
Dim m_MinWidth As Long

Private Sub ParentForm_Unload(Cancel As Integer)
  With ParentForm
    
    SaveSetting App.Title, .Name, "WindowState", .WindowState
    
    If .WindowState = vbNormal Then
      If m_RememberFormPosition Then
        SaveSetting App.Title, .Name, "Left", .Left
        SaveSetting App.Title, .Name, "Top", .Top
        SaveSetting App.Title, .Name, "Width", .Width
        SaveSetting App.Title, .Name, "Height", .Height
      End If
    End If
  End With
End Sub


Private Sub ParentForm_Load()
    m_sngPWidth = 0
    
    With ParentForm
      If m_RememberFormPosition Then
        .WindowState = GetSetting(App.Title, .Name, "WindowState", vbNormal)
        If .WindowState = vbNormal Then
          .Width = GetSetting(App.Title, .Name, "Width", .Width)
          .Height = GetSetting(App.Title, .Name, "Height", .Height)
          .Left = GetSetting(App.Title, .Name, "Left", (Screen.Width - .Width) / 2)
          .Top = GetSetting(App.Title, .Name, "Top", (Screen.Height - .Height) / 2)
        End If
      End If
      If Me.CenterForm Then
        If .WindowState = vbNormal Then
          .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
        End If
      End If
    End With
End Sub

Private Sub ParentForm_Resize()
  If ParentForm.WindowState = 1 Then Exit Sub
  If ParentForm.Width < MinWidth Then ParentForm.Width = MinWidth
  If ParentForm.Height < MinHeight Then ParentForm.Height = MinHeight
End Sub

Private Sub UserControl_Resize()
  Width = 240
  Height = 230
End Sub

Private Sub UserControl_Terminate()
  Set ParentForm = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("MinHeight", m_MinHeight, m_def_MinHeight)
  Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
  Call PropBag.WriteProperty("CenterForm", m_CenterForm, m_def_CenterForm)
  Call PropBag.WriteProperty("RememberFormPosition", m_RememberFormPosition, m_def_RememberFormPosition)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  If Not Ambient.UserMode = False Then
    '// Store the parent form
    Set ParentForm = Parent
  End If
  
  m_MinHeight = PropBag.ReadProperty("MinHeight", m_def_MinHeight)
  m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
  m_CenterForm = PropBag.ReadProperty("CenterForm", m_def_CenterForm)
  m_RememberFormPosition = PropBag.ReadProperty("RememberFormPosition", m_def_RememberFormPosition)

End Sub
Public Property Get MinHeight() As Long
Attribute MinHeight.VB_ProcData.VB_Invoke_Property = "Resize"
  MinHeight = m_MinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
  m_MinHeight = New_MinHeight
  PropertyChanged "MinHeight"
End Property

Public Property Get MinWidth() As Long
Attribute MinWidth.VB_ProcData.VB_Invoke_Property = "Resize"
  MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
  m_MinWidth = New_MinWidth
  PropertyChanged "MinWidth"
End Property
'
Public Property Get CenterForm() As Boolean
Attribute CenterForm.VB_ProcData.VB_Invoke_Property = "Resize"
  CenterForm = m_CenterForm
End Property

Public Property Let CenterForm(ByVal New_CenterForm As Boolean)
  m_CenterForm = New_CenterForm
  PropertyChanged "CenterForm"
End Property

Public Property Get RememberFormPosition() As Boolean
Attribute RememberFormPosition.VB_ProcData.VB_Invoke_Property = "Resize"
  RememberFormPosition = m_RememberFormPosition
End Property

Public Property Let RememberFormPosition(ByVal New_RememberFormPosition As Boolean)
  m_RememberFormPosition = New_RememberFormPosition
  PropertyChanged "RememberFormPosition"
End Property

