VERSION 5.00
Begin VB.Form frmdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date & Time"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdinsertdate 
      Caption         =   "&Insert"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox lstdate 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdinsertdate_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelText = " " & lstdate.Text & " "
    frmword.saved = 0
End Sub

Private Sub Form_Activate()
    lstdate.AddItem Format(Date, "Long Date")
    lstdate.AddItem Format(Date, "mm/dd/yyyy")
    lstdate.AddItem Format(Date, "dd/mm/yyyy")
    lstdate.AddItem Format(Date, "mmm dd/yyyy")
    lstdate.AddItem Format(Date, "mmmm dd/yyyy")
    lstdate.AddItem Format(Date, "ddd (dd/mm/yyyy)")
    lstdate.AddItem Format(Date, "dd-mm-yyyy")
    lstdate.AddItem Format(Date, "mm-dd-yyyy")
    lstdate.AddItem Format(Date, "dd(ddd)-mm-yyyy")
    lstdate.AddItem Format(Date, "dd mmm ,yyyy")
    lstdate.AddItem Format(Date, "mmmm,ddd(dd)-yyyy")
    lstdate.AddItem Format(Time, "hh:mm:ss AMPM")
    lstdate.AddItem Format(Time, "hh:mm AMPM")
    lstdate.AddItem Format(Time, "hh:mm:ss")
    lstdate.AddItem Format(Time, "hh:mm")
End Sub

Private Sub Form_Load()
    frmdate.Top = (Screen.Height - frmdate.Height) / 2
    frmdate.Left = (Screen.Width - frmdate.Width) / 2
End Sub

Private Sub lstdate_DblClick()
    cmdinsertdate_Click
End Sub
