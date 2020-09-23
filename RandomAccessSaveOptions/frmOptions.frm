VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "User Options"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   240
      MaxLength       =   40
      TabIndex        =   8
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Frame fraToolbar 
      Caption         =   "Toolbar:"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
      Begin VB.OptionButton optNone 
         Caption         =   "&None"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optSmall 
         Caption         =   "Sma&ll"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optFull 
         Caption         =   "F&ull-Sized"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CheckBox chkEmails 
      Caption         =   "&Format as HTML links"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CheckBox chkSplash 
      Caption         =   "&Show splash screen when program starts"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtDefNum 
      Height          =   285
      Left            =   240
      MaxLength       =   4
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblDefaultNum 
      Caption         =   "Default number to create:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload frmOptions
    Set frmOptions = Nothing
End Sub

Private Sub cmdSave_Click()
Dim udtUserOptions As OptionsStruc

'clear frmMain's captions
Call ClearMain

'get values & change frmMain's captions, etc.
udtUserOptions.intNumCreate = CInt(txtDefNum.Text)
frmMain.lblNumCreate.Caption = _
"Default number to create: " & udtUserOptions.intNumCreate

If chkSplash = Checked Then
    udtUserOptions.blnShowSplash = True
Else
    udtUserOptions.blnShowSplash = False
End If

frmMain.lblSplash.Caption = "Show splash screen? " & udtUserOptions.blnShowSplash

If chkEmails = Checked Then
    udtUserOptions.blnFormatHTML = True
Else
    udtUserOptions.blnFormatHTML = False
End If

frmMain.lblFormatHTML.Caption = "Format as HTML? " & udtUserOptions.blnFormatHTML

If optFull = True Then
    udtUserOptions.intToolBar = 1   'full-sized
    frmMain.lblToolbar.Caption = "Toolbar Size = Full-Sized"
    Call PopUpCheck(1)  ' #1 argument = full-sized
    Call frmMain.LargeToolbar
ElseIf optSmall = True Then
    udtUserOptions.intToolBar = 2   'small
    frmMain.lblToolbar.Caption = "Toolbar Size = Small"
    Call PopUpCheck(2)  ' #2 argument = small
    Call frmMain.SmallToolbar
ElseIf optNone = True Then
    udtUserOptions.intToolBar = 0   ' no toolbar
    frmMain.lblToolbar.Caption = "Toolbar Off"
    Call PopUpCheck(0)  ' #0 argument = no toolbar
Else
    udtUserOptions.intToolBar = 9   'diagnostic
    frmMain.lblToolbar.Caption = "Something went wrong when saving toolbar size"
End If

If txtUserName.Text <> "" Then
    udtUserOptions.strName = txtUserName.Text
    frmMain.lblUserName.Caption = "User Name: " & Trim(udtUserOptions.strName)
End If

Dim intFile As Integer  ' Free file number
Dim intCtr As Integer   ' Loop counter

'open random access file and put all the user's settings into it
intFile = FreeFile
Open App.Path & "\options.dat" For Random As intFile Len = Len(udtUserOptions)

Put #intFile, 1, udtUserOptions.intNumCreate
Put #intFile, 2, udtUserOptions.blnShowSplash
Put #intFile, 3, udtUserOptions.blnFormatHTML
Put #intFile, 4, udtUserOptions.intToolBar
Put #intFile, 5, udtUserOptions.strName

Close #intFile

Unload frmOptions
Set frmOptions = Nothing

End Sub

Private Sub Form_Load()
Dim intFileNumber As Integer
Dim udtUserOptions As OptionsStruc

' open random access file and get all the user's settings
intFileNumber = FreeFile
Open App.Path & "\options.dat" For Random As intFileNumber _
Len = Len(udtUserOptions)

Get #intFileNumber, 1, udtUserOptions.intNumCreate      '<-- Get [file number], [record number], [user-defined type variable]
Get #intFileNumber, 2, udtUserOptions.blnShowSplash
Get #intFileNumber, 3, udtUserOptions.blnFormatHTML
Get #intFileNumber, 4, udtUserOptions.intToolBar
Get #intFileNumber, 5, udtUserOptions.strName

' set things up in application based on user's settings
txtDefNum.Text = udtUserOptions.intNumCreate

If udtUserOptions.blnShowSplash = True Then
    chkSplash.Value = Checked
Else
    chkSplash.Value = Unchecked
End If

If udtUserOptions.blnFormatHTML = True Then
    chkEmails.Value = Checked
Else
    chkEmails.Value = Unchecked
End If

Select Case udtUserOptions.intToolBar
    Case 1
        optFull.Value = True
    Case 2
        optSmall.Value = True
    Case 0
        optNone.Value = True
End Select

txtUserName.Text = Trim(udtUserOptions.strName)

Close #intFileNumber    '<-- close the file
End Sub

Private Sub ClearMain()
    frmMain.lblNumCreate = ""
    frmMain.lblSplash = ""
    frmMain.lblFormatHTML = ""
    frmMain.lblToolbar = ""
    frmMain.lblUserName = ""
End Sub

Private Sub txtDefNum_KeyPress(KeyAscii As Integer)
' Don't let use type anything other than numbers or the backspace key
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    Beep
End If

End Sub

Private Sub txtDefNum_Validate(Cancel As Boolean)
    ' Make sure user hasn't cut and pasted any non-numbers in
    If Not IsNumeric(txtDefNum.Text) Then
        MsgBox _
        "Only numbers may be entered into the default number to create field", _
        vbCritical, "Error"
        txtDefNum.SetFocus
        txtDefNum.SelStart = 0
        txtDefNum.SelLength = Len(txtDefNum.Text)
    End If
End Sub
