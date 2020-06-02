VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin Project1.VPFormPos VPFormPos1 
      Left            =   1800
      Top             =   1620
      _ExtentX        =   423
      _ExtentY        =   397
      MinHeight       =   2655
      MinWidth        =   3480
      CenterForm      =   0   'False
   End
   Begin VB.Timer timProgress 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1305
      Top             =   1440
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Stop"
      Height          =   330
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   1575
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Start"
      Height          =   330
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   1575
      Width           =   1095
   End
   Begin Project1.VPProgress vpProgress 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   1965
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Label Label1 
      Caption         =   $"frmDemo.frx":0000
      Height          =   1455
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   3255
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlCount As Long

Private Sub cmdAction_Click(Index As Integer)
  Select Case Index
    Case 0  'start
      timProgress.Enabled = True
      
    Case 1  'stop
      timProgress.Enabled = False
      mlCount = 0
      Call vpProgress.DisplayMessage("")
  
  
  End Select
End Sub

Private Sub Form_Load()
  Call vpProgress.DisplayMessage("This is VPProgress.ctl")
End Sub

Private Sub timProgress_Timer()
  
  If mlCount = 10 Then
    mlCount = 0
  End If
  
  mlCount = mlCount + 1
  
  Call vpProgress.DisplayProgress(10, mlCount, "Test ")
  
End Sub
