VERSION 5.00
Begin VB.Form frmsearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find..."
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkcase 
      Caption         =   "Case Sensitive"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtreplacewith 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton cmdreplaceall 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdreplace 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtfindwhat 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblreplacewith 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   1290
   End
   Begin VB.Label lblfindwhat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find What ?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pos As Variant
Dim s As Integer
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdfind_Click()
Dim compare As Integer
frmmain.lblsearch.Caption = txtfindwhat.Text
frmsearch.Hide
If frmsearch.chkcase.Value = 1 Then
compare = 0
Else
compare = 1
End If
pos = InStr(pos + 1, frmmain.rtfword.Text, frmsearch.txtfindwhat.Text, compare)
If pos > 0 Then
cmdreplace.Enabled = True
cmdreplaceall.Enabled = True
frmmain.rtfword.SelStart = pos - 1
frmmain.rtfword.SelLength = Len(txtfindwhat.Text)
Else
s = MsgBox("String Not Found", vbCritical + vbOKOnly, "Search Result")
cmdreplace.Enabled = False
cmdreplaceall.Enabled = False
End If
frmsearch.Show 1
End Sub

Private Sub cmdreplace_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.Text = Replace(frmmain.rtfword.Text, txtfindwhat.Text, txtreplacewith.Text, , 1, vbTextCompare)
End Sub

Private Sub cmdreplaceall_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.Text = Replace(frmmain.rtfword.Text, txtfindwhat.Text, txtreplacewith.Text, , -1, vbTextCompare)
End Sub

Private Sub Form_Activate()
    txtfindwhat.SetFocus
End Sub

Private Sub Form_Load()
    frmsearch.Left = (Screen.Width - frmsearch.Width) / 2
    frmsearch.Top = (Screen.Height - frmsearch.Height) / 2
    frmsearch.Top = frmsearch.Top - 1000
End Sub

Private Sub txtfindwhat_Change()
    If txtfindwhat.Text = "" Then
        cmdfind.Enabled = False
    Else
        cmdfind.Enabled = True
    End If
End Sub
