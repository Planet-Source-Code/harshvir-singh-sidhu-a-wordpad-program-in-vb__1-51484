VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmword 
   BackColor       =   &H80000004&
   Caption         =   "MSP's WORDPAD"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar mspwordstatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "MSP's Wordpad"
            TextSave        =   "MSP's Wordpad"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "1/8/03"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "9:36 AM"
            Key             =   "date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   987
            MinWidth        =   987
            TextSave        =   "CAPS"
            Key             =   "time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   987
            MinWidth        =   987
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   987
            MinWidth        =   987
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   987
            MinWidth        =   987
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2160
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar mspwordtool 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1905
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":077E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0890
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":09A2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0AB4
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0BC6
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmword.frx":0CD8
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu cmdnew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu cmdopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu cmdfileline1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdsave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu cmdsaveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu cmdfileline2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdpagesetup 
         Caption         =   "Page Set&up"
      End
      Begin VB.Menu cmdprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu cmdfileline3 
         Caption         =   "-"
      End
      Begin VB.Menu cmdexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu cmdundo 
         Caption         =   "&Undo                  Ctrl+Z"
      End
      Begin VB.Menu cmdeditline1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdcut 
         Caption         =   "Cu&t                     Ctrl+X"
      End
      Begin VB.Menu cmdcopy 
         Caption         =   "&Copy                  Ctrl+C"
      End
      Begin VB.Menu cmdpaste 
         Caption         =   "&Paste                 Ctrl+V"
      End
      Begin VB.Menu cmddelete 
         Caption         =   "&Delete                Del"
      End
      Begin VB.Menu cmdselectall 
         Caption         =   "&Select All            Ctrl +A"
      End
   End
   Begin VB.Menu mnuinsert 
      Caption         =   "&Insert"
      Begin VB.Menu cmddateandtime 
         Caption         =   "&Date Time"
      End
      Begin VB.Menu cmdinsertpicture 
         Caption         =   "&Picture"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu cmdspellcheck 
         Caption         =   "Check Spellin&g"
         Shortcut        =   ^G
      End
      Begin VB.Menu cmdcountwords 
         Caption         =   "Count &Words"
      End
      Begin VB.Menu cmdlockedit 
         Caption         =   "&Lock Editing"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu cmdfind 
         Caption         =   "&Find"
      End
      Begin VB.Menu cmdfindnext 
         Caption         =   "Fint &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu cmdtoolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu cmdstatusbar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "F&ormat"
      Begin VB.Menu cmdbold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu cmditalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu cmdunderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu cmdbullets 
         Caption         =   "Bull&ets"
         Shortcut        =   ^E
      End
      Begin VB.Menu cmduppercase 
         Caption         =   "&Upper Case"
      End
      Begin VB.Menu cmdlowercase 
         Caption         =   "&Lower Case"
      End
      Begin VB.Menu cmdpropercase 
         Caption         =   "&Proper Case"
      End
      Begin VB.Menu cmdpageolor 
         Caption         =   "&Page Color"
      End
      Begin VB.Menu cmdfont 
         Caption         =   "&Font"
         Begin VB.Menu cmdfontcolor 
            Caption         =   "&Color"
         End
         Begin VB.Menu cmdfonttype 
            Caption         =   "&Type"
         End
      End
   End
End
Attribute VB_Name = "frmword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As Integer
Public saved As Integer
Dim strtser As String
Dim fname As String
Dim pos As Variant

Private Sub cmdbold_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelBold = Not frmmain.rtfword.SelBold
End Sub

Private Sub cmdbullets_Click()
    On Error Resume Next
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelBullet = Not frmmain.rtfword.SelBullet
End Sub

Private Sub cmdcopy_Click()
    On Error Resume Next
    Clipboard.SetText frmmain.rtfword.SelRTF
End Sub

Private Sub cmdcountwords_Click()
On Error GoTo errorwordcount
    Dim strrtf As String
    Dim temp As String
    Dim words As Single, characters As Single, lines As Single, noc As Single, spaces As Single, i As Single
    Dim curchar As String
    Dim valid As Boolean
    Screen.MousePointer = vbHourglass
    strrtf = frmmain.rtfword.Text
    noc = Len(strrtf)
    i = 1
    Do While i <= noc
    curchar = Mid(strrtf, i, 1)
    If curchar = " " Then
        spaces = spaces + 1
        words = words + 1
        temp = Mid(strrtf, i + 1, 1)
            If temp = " " Then
            i = i + 1
                Do While temp = " "
                i = i + 1
                temp = Mid(strrtf, i, 1)
                spaces = spaces + 1
                Loop
            End If
    ElseIf curchar = Chr(13) Or curchar = "." Or curchar = "!" Or curchar = "?" Or curchar = ";" Then
        lines = lines + 1
        words = words + 1
    ElseIf curchar = ":" Or curchar = "," Then
        words = words + 1
    Else
        characters = characters + 1
    End If
        i = i + 1
    Loop
    lines = lines + 1
    words = words + 1
    Screen.MousePointer = vbNormal
    s = MsgBox("Number of Lines = " & lines & vbCrLf & "Number of Words= " & words & vbCrLf & "Number of Characters= " & characters & vbCrLf & "Number of spaces= " & spaces, vbOKOnly, "WordCount Result")
    frmmain.rtfword.SetFocus
    Exit Sub
errorwordcount:
    s = MsgBox("Unknown Error Occured While Counting Words", vbInformation + vbOKOnly, "Word Count Error")
End Sub

Private Sub cmdcut_Click()
    On Error Resume Next
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    Clipboard.SetText frmmain.rtfword.SelRTF
    frmmain.rtfword.SelText = vbNullString
End Sub

Private Sub cmddateandtime_Click()
    frmdate.Show 1
End Sub

Private Sub cmddelete_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    If frmmain.rtfword.SelText <> "" Then
        frmmain.rtfword.SelText = ""
    Else
        frmmain.rtfword.SelLength = 1
        frmmain.rtfword.SelText = ""
    End If
End Sub

Private Sub cmdexit_Click()
    If saved = 1 Then
        Clipboard.Clear
        Unload Me
        End
    ElseIf saved = 0 Then
        s = MsgBox("Document [" & frmmain.Caption & "] Not Saved, save it.", vbInformation + vbYesNoCancel, "MSP's Wordpad Error")
        If s = vbYes Then
            cmdsave_Click
            cmdexit_Click
        ElseIf s = vbNo Then
            saved = 1
            cmdexit_Click
        End If
    End If
End Sub

Private Sub cmdfind_Click()
    frmsearch.Show 1
End Sub

Private Sub cmdfindnext_Click()
    strtser = frmmain.lblsearch.Caption
    If strtser = "" Then
        cmdfind_Click
    Else
        pos = InStr(pos + 1, frmmain.rtfword.Text, strtser, vbTextCompare)
        If pos > 0 Then
            frmmain.rtfword.SelStart = pos - 1
            frmmain.rtfword.SelLength = Len(strtser)
        Else
            s = MsgBox("String Not Found", vbCritical + vbOKOnly, "Search Result")
            pos = 0
            frmmain.rtfword.SelStart = 1
            frmmain.rtfword.SelLength = 0
        End If
    End If
End Sub

Private Sub cmdfontcolor_Click()
    On Error GoTo errorfontcolor
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    cd.CancelError = True
    cd.DialogTitle = "MSP's Font Color Selector"
    cd.ShowColor
    frmmain.rtfword.SelColor = cd.Color
    Exit Sub
errorfontcolor:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Error Occured While Changing Color.", vbInformation + vbOKOnly, "MSP's Wordpad Error")
End Sub

Private Sub cmdfonttype_Click()
    On Error GoTo errorfonttype
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    cd.CancelError = True
    cd.Flags = cdlCFBoth Or cdlCFApply Or cdlCFEffects
    cd.DialogTitle = "MSP's Wordpad Font Selector"
    cd.ShowFont
    frmmain.rtfword.SelFontName = cd.FontName
    frmmain.rtfword.SelFontSize = cd.FontSize
    frmmain.rtfword.SelItalic = cd.FontItalic
    frmmain.rtfword.SelBold = cd.FontBold
    frmmain.rtfword.SelUnderline = cd.FontUnderline
    frmmain.rtfword.SelStrikeThru = cd.FontStrikethru
    frmmain.rtfword.SelColor = cd.Color
    Exit Sub
errorfonttype:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown error occured while setting font", vbInformation + vbOKOnly, "MSP' Wordpad Error")
End Sub

Private Sub cmdinsertpicture_Click()
Dim fname As String
On Error GoTo pictureinserterror
    cd.CancelError = True
    cd.Flags = cdlOFNFileMustExist
    cd.DefaultExt = "*.bmp"
    cd.DialogTitle = "Select Image"
    cd.Filter = "*.bmp|*.bmp|*.jpg|*.jpg|*.gif|*.gif"
    Clipboard.Clear
    cd.ShowOpen
    fname = cd.FileName
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.picdata.Picture = LoadPicture(fname)
    Clipboard.SetData frmmain.picdata.Picture
    frmmain.rtfword.SelRTF = Clipboard.GetData
    Exit Sub
pictureinserterror:
If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Erro Occured while Inerting Picture", vbInformation + vbOKOnly, "Insert Picture Error")
End Sub

Private Sub cmditalic_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelItalic = Not frmmain.rtfword.SelItalic
End Sub

Private Sub cmdlockedit_Click()
    If cmdlockedit.Caption = "&Lock Editing" Then
        cmdlockedit.Caption = "Un&lock Editing"
        frmmain.rtfword.Locked = True
    ElseIf cmdlockedit.Caption = "Un&lock Editing" Then
        cmdlockedit.Caption = "&Lock Editing"
        frmmain.rtfword.Locked = False
    End If
End Sub

Private Sub cmdlowercase_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelText = StrConv(frmmain.rtfword.SelText, vbLowerCase)
End Sub

Private Sub cmdnew_Click()
    If saved = 1 Then
        frmmain.rtfword.TextRTF = ""
        frmmain.Caption = "Document"
    Else
        s = MsgBox("Document [" & frmmain.Caption & "] Not Saved ,save it.", vbInformation + vbYesNo, "MSP's WordPad Error")
        If s = vbYes Then
            cmdsave_Click
            cmdnew_Click
        ElseIf s = vbNo Then
            saved = 1
            cmdnew_Click
        End If
    End If
End Sub

Private Sub cmdopen_Click()
    On Error GoTo openerror
    If saved = 1 Then
        cd.CancelError = True
        cd.Flags = cdlOFNFileMustExist
        cd.DefaultExt = "*.rtf"
        cd.Filter = "Rich Text Documents|*.rtf|Text Files|*.txt|All Files|*.*"
        cd.DialogTitle = "MSP's Wordpad File Opener"
        cd.FileName = ""
        cd.ShowOpen
        fname = cd.FileName
        frmmain.rtfword.LoadFile fname
        frmmain.Caption = fname
        saved = 1
        Exit Sub
    Else
        s = MsgBox("Document [" & frmmain.Caption & "] Not Saved ,save it.", vbInformation + vbYesNo, "MSP's WordPad Error")
        If s = vbYes Then
            cmdsave_Click
            cmdopen_Click
        ElseIf s = vbNo Then
            saved = 1
            cmdopen_Click
        End If
    End If
    saved = 1
    Exit Sub
openerror:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Error Occurd While Opening File", vbInformation + vbOKOnly, "MSP's WordPad Error")
End Sub

Private Sub cmdpageolor_Click()
    On Error GoTo errorpagecolor
    cd.CancelError = True
    cd.DialogTitle = "MSP's Page Color Selector"
    cd.ShowColor
    frmmain.rtfword.BackColor = cd.Color
    Exit Sub
errorpagecolor:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Error Occured While Changing Color.", vbInformation + vbOKOnly, "MSP's Wordpad Error")
End Sub

Private Sub cmdpagesetup_Click()
    On Error GoTo errorpagesetup
    cd.CancelError = True
    cd.DialogTitle = "MSP's Wordpad Page Setup"
    cd.ShowPrinter
    Exit Sub
errorpagesetup:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Error Occured While Page Setup for Printing", vbCritical + vbOKOnly, "MSP's Wordpad Error")
End Sub

Private Sub cmdpaste_Click()
    On Error Resume Next
    If Clipboard.GetFormat(vbCFText) Then
        frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
        frmmain.rtfword.SelRTF = Clipboard.GetText
    ElseIf Clipboard.GetFormat(vbCFRTF) Then
        frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
        frmmain.rtfword.SelRTF = Clipboard.GetData(vbCFRTF)
    ElseIf Clipboard.GetFormat(vbCFBitmap) Then
        frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
        frmmain.rtfword.SelRTF = Clipboard.GetData
    End If
End Sub

Private Sub cmdprint_Click()
    On Error GoTo errorprint
    cd.DialogTitle = "Print"
    cd.CancelError = True
    cd.Flags = cdlPDReturnDC + cdlPDNoPageNums
        If frmmain.rtfword.SelLength = 0 Then
            cd.Flags = cd.Flags + cdlPDAllPages
        Else
            cd.Flags = cd.Flags + cdlPDSelection
        End If
        cd.ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            frmmain.rtfword.SelPrint cd.hDC
        End If
Exit Sub
errorprint:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Error Occured while Printing", vbCritical + vbOKOnly, "MSP's Wordpad Error")
End Sub

Private Sub cmdpropercase_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelText = StrConv(frmmain.rtfword.SelText, vbProperCase)
End Sub

Private Sub cmdsave_Click()
    On Error GoTo filesaveerror
    If Trim(fname) = "" Then
        cmdsaveas_Click
    End If
    If Right(Trim(fname), 4) = ".txt" Then
        Open fname For Output As #1
        Print #1, frmmain.rtfword.Text
        Close #1
    ElseIf Right(Trim(fname), 4) = ".rtf" Then
        frmmain.rtfword.SaveFile fname
    End If
    saved = 1
    Exit Sub
filesaveerror:
    If Err.Number = cdlCancel Then Exit Sub
    s = MsgBox("Unknown Error Occured While Saving File: [" & fname & "]", vbCritical + vbOKOnly, "MSP's Wordpad Error")
End Sub

Private Sub cmdsaveas_Click()
    On Error GoTo filesaveerror
    cd.Flags = cdlOFNOverwritePrompt
    cd.DefaultExt = "*.rtf"
    cd.Filter = "Rich Text Documents|*.rtf|Text Files|*.txt|All Files|*.*"
    cd.DialogTitle = "MSP's Wordpad File Saver"
    cd.FileName = ""
    cd.ShowSave
    fname = cd.FileName
    If Right(Trim(fname), 4) = ".txt" Then
        Open fname For Output As #1
        Print #1, frmmain.rtfword.Text
        Close #1
    ElseIf Right(Trim(fname), 4) = ".rtf" Then
        frmmain.rtfword.SaveFile fname
    End If
    saved = 1
    Exit Sub
filesaveerror:
    If Err.Number = cdlCancel Then
        saved = 0
        Exit Sub
    End If
    s = MsgBox("UnKnown Error Occured While Saving File: " & cd.FileName, vbCritical + vbOKOnly, "MSP's Wordpad Error")
End Sub

Private Sub cmdselectall_Click()
    frmmain.rtfword.SelStart = 0
    frmmain.rtfword.SelLength = Len(frmmain.rtfword.Text)
End Sub

Private Sub cmdspellcheck_Click()
    frmspellcheck.Show
End Sub

Private Sub cmdstatusbar_Click()
    If cmdstatusbar.Checked = True Then
        cmdstatusbar.Checked = False
        mspwordstatus.Visible = False
    ElseIf cmdstatusbar.Checked = False Then
        cmdstatusbar.Checked = True
        mspwordstatus.Visible = True
    End If
End Sub

Private Sub cmdtoolbar_Click()
    If cmdtoolbar.Checked = True Then
        cmdtoolbar.Checked = False
        mspwordtool.Visible = False
    ElseIf cmdtoolbar.Checked = False Then
        cmdtoolbar.Checked = True
        mspwordtool.Visible = True
    End If
End Sub

Private Sub cmdunderline_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelUnderline = Not frmmain.rtfword.SelUnderline
End Sub

Private Sub cmdundo_Click()
    frmmain.rtbtemp.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.TextRTF = frmmain.rtbclip.TextRTF
    frmmain.rtbclip.TextRTF = frmmain.rtbtemp.TextRTF
End Sub

Private Sub cmduppercase_Click()
    frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
    frmmain.rtfword.SelText = StrConv(frmmain.rtfword.SelText, vbUpperCase)
End Sub

Private Sub MDIForm_Load()
    saved = 1
End Sub

Private Sub mnuedit_Click()
    If frmmain.rtfword.SelLength > 0 Then
        cmdcut.Enabled = True
        cmdcopy.Enabled = True
    Else
        cmdcut.Enabled = False
        cmdcopy.Enabled = False
    End If
    If Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(vbCFBitmap) Then
        cmdpaste.Enabled = True
    Else
        cmdpaste.Enabled = False
    End If
End Sub

Private Sub mspwordtool_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            cmdnew_Click
        Case "Open"
            cmdopen_Click
        Case "Save"
            cmdsave_Click
        Case "Print"
            cmdprint_Click
        Case "Cut"
            cmdcut_Click
        Case "Copy"
            cmdcopy_Click
        Case "Paste"
            cmdpaste_Click
        Case "Bold"
            frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
            frmmain.rtfword.SelBold = Not frmmain.rtfword.SelBold
            Button.Value = IIf(frmmain.rtfword.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
            frmmain.rtfword.SelItalic = Not frmmain.rtfword.SelItalic
            Button.Value = IIf(frmmain.rtfword.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            frmmain.rtbclip.TextRTF = frmmain.rtfword.TextRTF
            frmmain.rtfword.SelUnderline = Not frmmain.rtfword.SelUnderline
            Button.Value = IIf(frmmain.rtfword.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            frmmain.rtfword.SelAlignment = rtfLeft
        Case "Center"
            frmmain.rtfword.SelAlignment = rtfCenter
        Case "Align Right"
            frmmain.rtfword.SelAlignment = rtfRight
    End Select
End Sub
