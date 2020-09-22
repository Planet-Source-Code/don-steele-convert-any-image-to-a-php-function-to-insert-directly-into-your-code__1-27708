VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsB64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private mvB64Enc(63) As Byte
Private mvB64Dec(255) As Byte

Public Sub Load(ByVal strPathName As String, ByRef pbArrayOutput() As Byte)
   Dim lSize     As Long
   Dim iFreeFile As Integer
   lSize = FileLen(strPathName)
   iFreeFile = FreeFile
ReDim pbArrayOutput(lSize - 1)
Open strPathName For Binary As iFreeFile
      Get iFreeFile, , pbArrayOutput
   Close iFreeFile
End Sub

Public Sub ByteArrayToString(ByRef pbArrayInput() As Byte, ByRef strOut As String)
   strOut = StrConv(pbArrayInput, vbUnicode)
End Sub

Public Sub Encode(ByRef pbArrayInput() As Byte, ByRef pbArrayOutput() As Byte)
   Dim iSizeMod As Integer
   Dim lSizeIn  As Long
   Dim lSizeOut As Long
   Dim Index    As Long
   Dim lIndex2  As Long
   Dim lTotal   As Long
   Dim bBuffer(2) As Byte

   lSizeIn = UBound(pbArrayInput) + 1
   iSizeMod = lSizeIn Mod 3
   lSizeOut = ((lSizeIn - iSizeMod) \ 3) * 4
   If iSizeMod > 0 Then lSizeOut = lSizeOut + 4
   
   
   ReDim pbArrayOutput(lSizeOut - 1)
   
  
   If lSizeIn >= 3 Then
      
      lTotal = lSizeIn - iSizeMod - 1
      For Index = 0 To lTotal Step 3
        
         bBuffer(0) = pbArrayInput(Index)
         bBuffer(1) = pbArrayInput(Index + 1)
         bBuffer(2) = pbArrayInput(Index + 2)
pbArrayOutput(lIndex2) = mvB64Enc((bBuffer(0) And &HFC) \ 4)
         pbArrayOutput(lIndex2 + 1) = mvB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
         pbArrayOutput(lIndex2 + 2) = mvB64Enc((bBuffer(1) And &HF) * 4 Or (bBuffer(2) And &HC0) \ 64)
         pbArrayOutput(lIndex2 + 3) = mvB64Enc((bBuffer(2) And &H3F))
         lIndex2 = lIndex2 + 4
      Next
   End If

  
   Select Case iSizeMod
      Case 1
         bBuffer(0) = pbArrayInput(lSizeIn - 1)
         bBuffer(1) = 0
    pbArrayOutput(lIndex2) = mvB64Enc((bBuffer(0) And &HFC) \ 4)
         pbArrayOutput(lIndex2 + 1) = mvB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
         pbArrayOutput(lIndex2 + 2) = 61
         pbArrayOutput(lIndex2 + 3) = 61
      Case 2
         bBuffer(0) = pbArrayInput(lSizeIn - 2)
         bBuffer(1) = pbArrayInput(lSizeIn - 1)
         bBuffer(2) = 0
      pbArrayOutput(lIndex2) = mvB64Enc((bBuffer(0) And &HFC) \ 4)
         pbArrayOutput(lIndex2 + 1) = mvB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
         pbArrayOutput(lIndex2 + 2) = mvB64Enc((bBuffer(1) And &HF) * 4 Or (bBuffer(2) And &HC0) \ 64)
         pbArrayOutput(lIndex2 + 3) = 61
   End Select
End Sub


Private Sub Class_Initialize()
   Dim iIndex As Integer
For iIndex = 65 To 90
      mvB64Enc(iIndex - 65) = iIndex
   Next
For iIndex = 97 To 122
      mvB64Enc(iIndex - 71) = iIndex
   Next
For iIndex = 48 To 57
      mvB64Enc(iIndex + 4) = iIndex
   Next
   mvB64Enc(62) = 43
   mvB64Enc(63) = 47
End Sub
