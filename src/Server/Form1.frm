VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rot Server"
   ClientHeight    =   2568
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   2568
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   252
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   4380
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   252
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Waiting for client connection. . ."
      Top             =   840
      Width           =   4380
   End
   Begin VB.TextBox txtFolder 
      Height          =   288
      Left            =   252
      TabIndex        =   0
      Top             =   336
      Width           =   4380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oWorker    As cWorker
Attribute m_oWorker.VB_VarHelpID = -1
Private m_lTotal                As Long

Private Sub Form_Load()
    Set m_oWorker = New cWorker
    txtFolder.Text = Environ$("windir") & "\System32"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtFolder.Width = ScaleWidth - 2 * txtFolder.Left
    txtFileName.Width = ScaleWidth - 2 * txtFileName.Left
End Sub

Private Sub txtFolder_Change()
    m_oWorker.frTargetFolder = txtFolder.Text
End Sub

Private Sub m_oWorker_ReadFileComplete(FileName As String, ByVal FileSize As Long)
    txtFileName.Text = FileName
    m_lTotal = m_lTotal + FileSize
    txtTotal.Text = Format$((m_lTotal + 1023) \ 1024, "#,#0") & " KB"
End Sub

