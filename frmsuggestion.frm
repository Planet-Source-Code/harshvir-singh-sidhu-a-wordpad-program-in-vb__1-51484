VERSION 5.00
Begin VB.Form frmsuggestion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSP's Suggestion"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdreplace 
      Caption         =   "&Replace"
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
      Left            =   1920
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   4740
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suggestion"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Possible Errors"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "frmsuggestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdreplace_Click()
    If List1.ListIndex = -1 Then
        s = MsgBox("Please select a Word", vbInformation + vbOKOnly, "MSP's Word Replace Error")
        Exit Sub
    End If
    If List2.ListIndex = -1 Then
        s = MsgBox("Please select an Alternate Spelling", vbInformation + vbOKOnly, "MSP's Word Replace Error")
        Exit Sub
    End If
    frmmain.rtfword.TextRTF = Replace(frmmain.rtfword.TextRTF, List1.Text, List2.Text)
End Sub

Private Sub List1_Click()
    Screen.MousePointer = vbHourglass
    Set CorrectionsCollection = _
        AppWord.GetSpellingSuggestions(List1.Text)
    List2.Clear
    For iSuggWord = 1 To CorrectionsCollection.Count
        List2.AddItem CorrectionsCollection.Item(iSuggWord)
    Next
    Screen.MousePointer = vbDefault
End Sub
