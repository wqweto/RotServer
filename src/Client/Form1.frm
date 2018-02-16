VERSION 5.00
Object = "{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}#1.0#0"; "Inked.dll"
Begin VB.Form Form1 
   Caption         =   "Rot Client"
   ClientHeight    =   3732
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5532
   LinkTopic       =   "Form1"
   ScaleHeight     =   3732
   ScaleWidth      =   5532
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   600
      Left            =   252
      TabIndex        =   0
      Top             =   168
      Width           =   1440
   End
   Begin INKEDLibCtl.InkEdit txtLog 
      Height          =   2784
      Left            =   252
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   840
      Width           =   5136
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_VSCROLL                    As Long = &H115
Private Const SB_BOTTOM                     As Long = 7

Private Declare Function CLSIDFromString Lib "ole32" (ByVal szPtr As Long, clsid As Any) As Long
Private Declare Function GetActiveObject Lib "oleaut32" (lpRclsid As Any, ByVal pvReserved As Long, pUnk As IUnknown) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const STR_ROT_SERVER_GUID As String = "{82b006f4-ca87-423a-b048-a160373bea72}"

Private Sub cmdProcess_Click()
    Dim aGuid(0 To 3)   As Long
    Dim oRotServer      As Object
    
    On Error GoTo EH
    Call CLSIDFromString(StrPtr(STR_ROT_SERVER_GUID), aGuid(0))
    Call GetActiveObject(aGuid(0), 0, oRotServer)
    Process oRotServer
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "cmdProcess_Click"
End Sub

Private Sub Process(oRotServer As Object)
    Dim vElem           As Variant
    Dim baData()        As Byte
    
    On Error GoTo EH
    txtLog.Text = vbNullString
    For Each vElem In oRotServer.EnumFolder
        LogDebug vElem
        baData = oRotServer.ReadBinaryFile(CStr(vElem))
        LogDebug Format(UBound(baData) + 1, "#,#0") & " bytes"
    Next
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "Process"
End Sub

Private Sub LogDebug(ByVal sText As String)
    txtLog.SelStart = &H7FFFFFFF
    txtLog.SelText = sText & vbCrLf
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0&)
End Sub
