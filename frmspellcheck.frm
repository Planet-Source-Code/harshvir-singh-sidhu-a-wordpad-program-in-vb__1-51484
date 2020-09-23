VERSION 5.00
Begin VB.Form frmspellcheck 
   Caption         =   "MSP's SpellChecker Using Word"
   ClientHeight    =   390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MSP's Spell Checker"
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
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   2145
   End
End
Attribute VB_Name = "frmspellcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Variant

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim DRange As Range
    Me.Caption = "Initializing..."
    On Error Resume Next
    Set AppWord = GetObject(, "Word.Application")
    If AppWord Is Nothing Then
        Set AppWord = CreateObject("Word.Application")
        If AppWord Is Nothing Then
            s = MsgBox("Could not start Word's Spell Checker.", vbInformation + vbOKOnly, "MSP's Wordpad Error")
            Unload Me
        Else
            NewInstance = True
        End If
    Else
        NewInstance = False
    End If
    AppWord.Documents.Add
    Me.Caption = "Checking words..."
    Set DRange = AppWord.ActiveDocument.Range
    DRange.InsertAfter frmmain.rtfword.Text
    Set SpellCollection = DRange.SpellingErrors
    frmsuggestion.List1.Clear
    frmsuggestion.List2.Clear
    If SpellCollection.Count > 0 Then
        For iWord = 1 To SpellCollection.Count
            frmsuggestion.List1.AddItem SpellCollection.Item(iWord)
            If frmsuggestion.List1.List(frmsuggestion.List1.NewIndex) = frmsuggestion.List1.List(frmsuggestion.List1.NewIndex + 1) Then
                frmsuggestion.List1.RemoveItem frmsuggestion.List1.NewIndex
            End If
        Next
    End If
    frmsuggestion.Show 1
    Me.Caption = "Closing Word..."
    AppWord.ActiveDocument.Close False
    If NewInstance Then
        AppWord.Quit
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    Unload Me
    s = MsgBox("Error occured During Starting Word For Spell Checking.", vbCritical + vbOKOnly, "MSP's Wordpad Spellchecker Error")
End Sub
