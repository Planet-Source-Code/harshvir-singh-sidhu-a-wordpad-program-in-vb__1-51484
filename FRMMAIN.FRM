VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Document"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picdata 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbclip 
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmmain.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbtemp 
      Height          =   735
      Left            =   6840
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmmain.frx":00C9
   End
   Begin RichTextLib.RichTextBox rtfword 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmmain.frx":0192
   End
   Begin VB.Label lblsearch 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 On Error Resume Next
    rtfword.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfword.RightMargin = rtfword.Width - 400
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    rtfword.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfword.RightMargin = rtfword.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

Private Sub rtfword_Change()
    frmword.saved = 0
End Sub

Private Sub rtfword_KeyUp(KeyCode As Integer, Shift As Integer)
   If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyL
                frmmain.rtfword.Span ".?!", False, True
                SelectionStart = frmmain.rtfword.SelStart
                frmmain.rtfword.Span ".?!", True, True
                SelectionEnd = frmmain.rtfword.SelStart + frmmain.rtfword.SelLength
                frmmain.rtfword.SelStart = SelectionStart
                frmmain.rtfword.SelLength = SelectionEnd - SelectionStart
            Case vbKeyW
                frmmain.rtfword.Span " ,;:.?!", False, True
                SelectionStart = frmmain.rtfword.SelStart
                frmmain.rtfword.Span " ,;:.?!", True, True
                SelectionEnd = frmmain.rtfword.SelStart + frmmain.rtfword.SelLength
                frmmain.rtfword.SelStart = SelectionStart
                frmmain.rtfword.SelLength = SelectionEnd - SelectionStart
            End Select
    End If
    If Shift = (vbCtrlMask Or vbShiftMask) Then
        Select Case KeyCode
        Case vbKeyL
            frmmain.rtfword.UpTo ".?!", True, False
        Case vbKeyW
            frmmain.rtfword.UpTo " ,;:.?!", True, False
        End Select
    End If
End Sub

Private Sub rtfword_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Clipboard.GetFormat(vbCFBitmap) Then
        frmmain.rtfword.SelText = Clipboard.GetData
    End If
End Sub

Private Sub rtfword_SelChange()
    frmword.mspwordtool.Buttons("Bold").Value = IIf(rtfword.SelBold, tbrPressed, tbrUnpressed)
    frmword.mspwordtool.Buttons("Italic").Value = IIf(rtfword.SelItalic, tbrPressed, tbrUnpressed)
    frmword.mspwordtool.Buttons("Underline").Value = IIf(rtfword.SelUnderline, tbrPressed, tbrUnpressed)
    frmword.mspwordtool.Buttons("Align Left").Value = IIf(rtfword.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    frmword.mspwordtool.Buttons("Center").Value = IIf(rtfword.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    frmword.mspwordtool.Buttons("Align Right").Value = IIf(rtfword.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
    frmword.saved = 0
End Sub
