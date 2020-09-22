VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Picture2Function"
   ClientHeight    =   5850
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox text1 
      Height          =   4335
      Left            =   540
      TabIndex        =   4
      Top             =   720
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7646
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   90
      Top             =   5370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   180
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   630
      TabIndex        =   0
      Top             =   150
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate!"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   210
      Width           =   615
   End
   Begin VB.Menu fre 
      Caption         =   "fre"
      Visible         =   0   'False
      Begin VB.Menu slcall 
         Caption         =   "Select All"
      End
      Begin VB.Menu cpy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim j
Dim g As String
Dim len32
Dim crlf
Dim u1
Dim u2
Dim Fname As String
Dim Ftype As String
Dim Mimetype As String
If Len(Trim(Text2.Text)) < 5 Then
MsgBox "Please select a valid file name"
Text2.SetFocus
Exit Sub
End If

'=
If file_exists(Trim(Text2.Text)) = False Then
MsgBox "Please select a valid file name"
Text2.SetFocus
Exit Sub
'=
End If
u1 = Split(Text2.Text, "\")
u2 = Split(Text2.Text, ".")
Fname = Trim(CStr(u1(UBound(u1))))
Ftype = Trim(CStr(u2(UBound(u2))))
Select Case LCase(Ftype)
Case "gif"
Mimetype = "Content-type: image/gif"
Case "jpg"
Mimetype = "Content-type: image/jpeg"
Case "jpeg"
Mimetype = "Content-type: image/jpeg"
Case "jpe"
Mimetype = "Content-type: image/jpeg"
Case "bmp"
Mimetype = "Content-type: image/bmp"
Case "png"
Mimetype = "Content-type: image/png"
Case "tiff"
Mimetype = "Content-type: image/tiff"
Case "tif"
Mimetype = "Content-type: image/tiff"
Case Else
MsgBox "Unable to determine file type!", vbCritical
Exit Sub
End Select

text1.Text = ""
crlf = Chr(13) & Chr(10)
Open Text2.Text For Input As #1
len32 = LOF(1)
Close #1
g = ""
g = "function " & Fname & "() {" & crlf
g = g & "header(" & Chr(34) & Mimetype & Chr(34) & ");" & crlf
g = g & "header(" & Chr(34) & "Content-length: " & len32 & Chr(34) & ");" & crlf
g = g & "echo base64_decode(" & crlf

Main
clsBase64.Load Text2.Text, pbBuffer1
clsBase64.Encode pbBuffer1, pbBuffer2
 clsBase64.ByteArrayToString pbBuffer2, ptContent
For j = 1 To Len(ptContent) Step 45
g = g & "'" & Mid(ptContent, j, 45) & "'." & crlf

Next
g = Mid(g, 1, Len(g) - 3)
g = g & ");" & crlf & "}"
text1.Text = g
End Sub

Private Sub Command2_Click()
Text2.SetFocus
cd.Filter = "Supported Files|*.gif; *.jpg; *.jpeg; *.jpe; *.bmp; *.png; *.tif; *.tiff|All Files|*.*|"
cd.ShowOpen
Text2.Text = cd.FileName
End Sub

Private Sub cpy_Click()
Clipboard.SetText CStr(text1.SelText)

End Sub

Private Sub slcall_Click()
text1.SetFocus
SendKeys Chr(1)
End Sub

Private Sub text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu fre
End Sub


Private Function file_exists(filepath As String) As Boolean
If Len(Dir(filepath)) > 0 Then
file_exists = True
Else
file_exists = False
End If

End Function
