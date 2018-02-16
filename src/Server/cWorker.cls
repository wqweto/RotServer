VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefObj A-Z

'=========================================================================
' Public events
'=========================================================================

Event ReadFileComplete(FileName As String, ByVal FileSize As Long)

'=========================================================================
' API
'=========================================================================

'Private Const ACTIVEOBJECT_STRONG       As Long = 0
Private Const ACTIVEOBJECT_WEAK         As Long = 1

Private Declare Function CLSIDFromString Lib "ole32" (ByVal szPtr As Long, clsid As Any) As Long
Private Declare Function CoDisconnectObject Lib "ole32" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32" (ByVal pUnk As IUnknown, rclsid As Any, ByVal dwFlags As Long, pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ROT_SERVER_GUID As String = "{82b006f4-ca87-423a-b048-a160373bea72}"

Private m_lInstance         As Long
Private m_sTargetFolder     As String

'=========================================================================
' Properties
'=========================================================================

Property Get frTargetFolder() As String
    frTargetFolder = m_sTargetFolder
End Property

Property Let frTargetFolder(sValue As String)
    m_sTargetFolder = sValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function EnumFolder(Optional FileMask As String) As Variant
    Dim cRetVal         As Collection
    Dim sFile           As String
    
    On Error GoTo EH
    Set cRetVal = New Collection
    sFile = Dir(PathCombine(m_sTargetFolder, IIf(LenB(FileMask) <> 0, FileMask, "*.*")))
    Do While LenB(sFile) <> 0
        If sFile <> "." And sFile <> ".." Then
            cRetVal.Add sFile
        End If
        sFile = Dir
    Loop
    EnumFolder = ToArray(cRetVal)
    Exit Function
EH:
    EnumFolder = Array()
End Function

Public Function ReadBinaryFile(ByVal sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    
    On Error GoTo EH
    sFile = PathCombine(m_sTargetFolder, sFile)
    baBuffer = vbNullString
    nFile = FreeFile
    If GetAttr(sFile) Then
        Open sFile For Binary Access Read Shared As nFile
        If LOF(nFile) > 0 Then
            ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
            Get nFile, , baBuffer
        End If
        Close nFile
    End If
    RaiseEvent ReadFileComplete(sFile, UBound(baBuffer) + 1)
    ReadBinaryFile = baBuffer
    Exit Function
EH:
    ReadBinaryFile = vbNullString
End Function

'= private ==============================================================

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Function ToArray(oCol As Collection) As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    If oCol.Count > 0 Then
        ReDim vRetVal(0 To oCol.Count - 1) As Variant
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = oCol(lIdx + 1)
        Next
        ToArray = vRetVal
    Else
        ToArray = Array()
    End If
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Dim aGuid(0 To 3)   As Long
    
    m_lInstance = 0
    Call CLSIDFromString(StrPtr(STR_ROT_SERVER_GUID), aGuid(0))
    Call RegisterActiveObject(Me, aGuid(0), ACTIVEOBJECT_WEAK, m_lInstance)
End Sub
 
Private Sub Class_Terminate()
    If m_lInstance <> 0 Then
        Call RevokeActiveObject(m_lInstance, 0)
    End If
    Call CoDisconnectObject(Me, 0)
End Sub
