VERSION 5.00
Object = "{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}#1.0#0"; "Inked.dll"
Begin VB.Form Form1 
   Caption         =   "Rot Client"
   ClientHeight    =   4896
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4896
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTime 
      Height          =   348
      Left            =   252
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   756
      Width           =   5136
   End
   Begin VB.CommandButton cmdBounce 
      Caption         =   "Bounce time"
      Height          =   432
      Left            =   252
      TabIndex        =   0
      Top             =   168
      Width           =   1440
   End
   Begin VB.CommandButton cmdFetchFolder 
      Caption         =   "Fetch folder"
      Height          =   432
      Left            =   252
      TabIndex        =   2
      Top             =   1260
      Width           =   1440
   End
   Begin INKEDLibCtl.InkEdit txtLog 
      Height          =   2784
      Left            =   252
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   3
      Top             =   1848
      Width           =   5136
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const WM_VSCROLL                    As Long = &H115
Private Const SB_BOTTOM                     As Long = 7

Private Declare Function CLSIDFromString Lib "ole32" (ByVal szPtr As Long, clsid As Any) As Long
Private Declare Function GetActiveObject Lib "oleaut32" (lpRclsid As Any, ByVal pvReserved As Long, pUnk As IUnknown) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)

Private Type SYSTEMTIME
   wYear            As Integer
   wMonth           As Integer
   wDayOfWeek       As Integer
   wDay             As Integer
   wHour            As Integer
   wMinute          As Integer
   wSecond          As Integer
   wMilliseconds    As Integer
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ROT_SERVER_GUID As String = "{82b006f4-ca87-423a-b048-a160373bea72}"

Private m_bCancel           As Boolean

'=========================================================================
' Methods
'=========================================================================

Private Sub LogDebug(ByVal sText As String)
    txtLog.SelStart = &H7FFFFFFF
    txtLog.SelText = sText & vbCrLf
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0&)
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub cmdBounce_Click()
    Const LNG_COUNT     As Long = 10000
    Dim aGuid(0 To 3)   As Long
    Dim oRotServer      As Object
    Dim lIdx            As Long
    Dim lStart          As Long
    Dim uTime           As SYSTEMTIME
    Dim lResult         As Long
    
    On Error GoTo EH
    Call CLSIDFromString(StrPtr(STR_ROT_SERVER_GUID), aGuid(0))
    Call GetActiveObject(aGuid(0), 0, oRotServer)
    Call GetSystemTime(uTime)
    lStart = uTime.wMinute * 60000 + uTime.wSecond * 1000& + uTime.wMilliseconds
    For lIdx = 1 To LNG_COUNT
        Call GetSystemTime(uTime)
        lResult = oRotServer.BounceLong(uTime.wMinute * 60000 + uTime.wSecond * 1000& + uTime.wMilliseconds - lStart)
        txtTime.Text = lResult
        DoEvents
        If m_bCancel Then
            Exit Sub
        End If
    Next
Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "cmdBounce_Click"
End Sub

Private Sub cmdFetchFolder_Click()
    Dim aGuid(0 To 3)   As Long
    Dim oRotServer      As Object
    Dim vElem           As Variant
    Dim baData()        As Byte
    
    On Error GoTo EH
    Call CLSIDFromString(StrPtr(STR_ROT_SERVER_GUID), aGuid(0))
    Call GetActiveObject(aGuid(0), 0, oRotServer)
    txtLog.Text = vbNullString
    For Each vElem In oRotServer.EnumFolder
        LogDebug vElem
        baData = oRotServer.ReadBinaryFile(CStr(vElem))
        LogDebug Format(UBound(baData) + 1, "#,#0") & " bytes"
        DoEvents
        If m_bCancel Then
            Exit Sub
        End If
    Next
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "cmdFetchFolder_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_bCancel = True
End Sub
