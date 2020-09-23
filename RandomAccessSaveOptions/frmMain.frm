VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Main Page"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlSmallPics 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":039A
            Key             =   "newfile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0734
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ACE
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E68
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1202
            Key             =   "emails"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":135C
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPics 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14B6
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1908
            Key             =   "newfile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C22
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F3C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2256
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2570
            Key             =   "emails"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":288A
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open a File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newfile"
            Object.ToolTipText     =   "New File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "emails"
            Object.ToolTipText     =   "Emails"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Get Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&User Options"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblUserName 
      Caption         =   "User Name = Gumboots McGanglepang"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblToolbar 
      Caption         =   "Toolbar = Full-Sized/Small/None"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblFormatHTML 
      Caption         =   "Format as HTML = True/False"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblSplash 
      Caption         =   "Splash screen = True/Fales"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblNumCreate 
      Caption         =   "Default number to create: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "&PopUp"
      Begin VB.Menu mnuPopUpFullTools 
         Caption         =   "&Full-Sized Toolbar"
      End
      Begin VB.Menu mnuPopUpSmallTools 
         Caption         =   "&Small Toolbar"
      End
      Begin VB.Menu mnuPopUpNoTools 
         Caption         =   "&No Toolbar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Author: Tom Meinen
' E-mail: tomsoftware@vegsource.com
' Save user options to a random access file so
' you don't have to fool with the registry.  This
' app saves each user option as its own record
' in a random access file.  Sometimes it only accesses
' and changes one record as is the case with its
' right click toolbar choice.  It also shows you
' how to give the user the choice of seeing the
' splash screen or not by using Sub Main.  You
' may use this code as you wish.  If it has helped
' you or if you have other comments, you may e-mail
' me.  You must have the
' original icons that came with this app's zip file
' in order for the right-click/ resize toolbar function
' to run correctly.  Please only redistribute this code
' with my original zip file so that the recipient gets
' all the files.
'------------------------------------------------------------


Option Explicit

Private Sub cmdOptions_Click()
    frmOptions.Show
End Sub

Private Sub Form_Load()
Dim intFileNumber As Integer
Dim udtUserOptions As OptionsStruc

' hide popup menu
mnuPopUp.Visible = False    '<-- I prefer hiding it in Form_Load instead of in the
                                 'menu editor so that the popup menu is always visible
                                 'at design time

intFileNumber = FreeFile

'Open random access file and get all the user's settings:
Open App.Path & "\options.dat" For Random As intFileNumber _
Len = Len(udtUserOptions)

Get #intFileNumber, 1, udtUserOptions.intNumCreate      '<-- Get [file number], [record number], [user-defined type variable]
Get #intFileNumber, 2, udtUserOptions.blnShowSplash
Get #intFileNumber, 3, udtUserOptions.blnFormatHTML
Get #intFileNumber, 4, udtUserOptions.intToolBar
Get #intFileNumber, 5, udtUserOptions.strName

' Change captions based on user's settings:
lblNumCreate.Caption = "Default number to create: " & udtUserOptions.intNumCreate
lblSplash.Caption = "Show splash screen? " & udtUserOptions.blnShowSplash
lblFormatHTML.Caption = "Format as HTML? " & udtUserOptions.blnFormatHTML

lblToolbar.Caption = "Toolbar Size = "
Select Case udtUserOptions.intToolBar
    Case 1
        lblToolbar.Caption = lblToolbar.Caption & "Full-Sized"
        'set popup menu:
        Call PopUpCheck(1)  ' #1 argument = full-sized toolbar
    Case 2
        lblToolbar.Caption = lblToolbar.Caption & "Small"
        Call PopUpCheck(2)  ' #2 argument = small toolbar
        Call SmallToolbar
    Case 0
        lblToolbar.Caption = "Toolbar Off"
        Call PopUpCheck(0)  ' #0 argument = no toolbar
        tbrTools.Visible = False
    Case 9
        lblToolbar.Caption = "Something went wrong when saving toolbar size"
    Case Else
        lblToolbar.Caption = "Something didn't work"
End Select

lblUserName.Caption = "User Name: " & Trim(udtUserOptions.strName)

Close #intFileNumber    ' <-- close the file

Call AdjustOptionsButton
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' show popup menu
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub lblFormatHTML_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' show popup menu
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub lblNumCreate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' show popup menu
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub
Private Sub lblSplash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' show popup menu
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub lblToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' show popup menu
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub lblUserName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' show popup menu
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub mnuPopUpFullTools_Click()   'make toolbar full-sized
    Call LargeToolbar
    Call PopUpCheck(1)  ' argument #1 = large
    Call SaveToolbarSetting(1)  ' #1 arg = large
End Sub

Private Sub mnuPopUpNoTools_Click()     'hide toolbar
    Call PopUpCheck(0)  ' 0 = no toolbar
    Call SaveToolbarSetting(0)  '#0 arg = no toolbar
End Sub

Private Sub mnuPopUpSmallTools_Click()  'make toolbar small
    Call SmallToolbar
    Call PopUpCheck(2)  ' argument #2 = small
    Call SaveToolbarSetting(2)  ' #2 = small
End Sub

Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
' If user clicks on toolbar, execute code from whichever button:
    Select Case Button.Key
        Case "open"
            MsgBox "Open a file code would be here."
        Case "newfile"
            MsgBox "Start a new file code would be here."
        Case "cut"
            MsgBox "Cut text code would be here."
        Case "copy"
            MsgBox "Copy text code would be here."
        Case "paste"
            MsgBox "Paste text code would be here."
        Case "emails"
            MsgBox "Email code would be here."
        Case "help"
            MsgBox "Help code would be here."
    End Select
End Sub

Public Sub SmallToolbar()       ' must be public so frmOptions can access it
' Make toolbar small subroutine:

    Dim clsNewImage             As MSComctlLib.ListImage
    Dim clsNewButton            As MSComctlLib.Button

    '---------------------------
    ' Can't modify image list control
    ' if it's referenced by other controls.
    '---------------------------
    Set Me.tbrTools.ImageList = Nothing

    '---------------------------
    ' Add images to the image list control
    '---------------------------
    imlSmallPics.ListImages.Clear
    imlSmallPics.ImageHeight = 16
    imlSmallPics.ImageWidth = 16

    With Me.imlSmallPics.ListImages
        Set clsNewImage = .Add(, "Image1", _
            LoadPicture(App.Path & "\openfold_sm.ico"))
        Set clsNewImage = .Add(, "Image2", _
            LoadPicture(App.Path & "\newfile_sm.ico"))
        Set clsNewImage = .Add(, "Image3", _
            LoadPicture(App.Path & "\cut_sm.ico"))
        Set clsNewImage = .Add(, "Image4", _
            LoadPicture(App.Path & "\copy_sm.ico"))
        Set clsNewImage = .Add(, "Image5", _
            LoadPicture(App.Path & "\paste_sm.ico"))
        Set clsNewImage = .Add(, "Image6", _
            LoadPicture(App.Path & "\emails_sm.ico"))
        Set clsNewImage = .Add(, "Image7", _
            LoadPicture(App.Path & "\help2_sm.ico"))
    End With

    '---------------------------
    ' Add toolbar buttons and ToolTipText
    '---------------------------
        Set Me.tbrTools.ImageList = Me.imlSmallPics

        With tbrTools.Buttons
            .Clear

            Set clsNewButton = .Add(1, "open", , , "Image1")
                clsNewButton.ToolTipText = "Open a File"
            Set clsNewButton = .Add(2, "newfile", , , "Image2")
                clsNewButton.ToolTipText = "New File"
            Set clsNewButton = .Add(3, "cut", , , "Image3")
                clsNewButton.ToolTipText = "Cut"
            Set clsNewButton = .Add(4, "copy", , , "Image4")
                clsNewButton.ToolTipText = "Copy"
            Set clsNewButton = .Add(5, "paste", , , "Image5")
                clsNewButton.ToolTipText = "Paste"
            Set clsNewButton = .Add(6, "emails", , , "Image6")
                clsNewButton.ToolTipText = "Emails"
            Set clsNewButton = .Add(7, "help", , , "Image7")
                clsNewButton.ToolTipText = "Get Help"
        End With
    
End Sub

Public Sub LargeToolbar()   ' must be public so frmOptions can access it
' make toolbar large subroutine

    Dim clsNewImage             As MSComctlLib.ListImage
    Dim clsNewButton            As MSComctlLib.Button

    '---------------------------
    ' Can't modify image list control
    ' if it's referenced by other controls.
    '---------------------------
    Set Me.tbrTools.ImageList = Nothing

    '---------------------------
    ' Add images to the image list control
    '---------------------------
    imlPics.ListImages.Clear
    imlPics.ImageHeight = 32
    imlPics.ImageWidth = 32

    With Me.imlPics.ListImages
        Set clsNewImage = .Add(, "Image1", _
            LoadPicture(App.Path & "\openfold.ico"))
        Set clsNewImage = .Add(, "Image2", _
            LoadPicture(App.Path & "\newfile.ico"))
        Set clsNewImage = .Add(, "Image3", _
            LoadPicture(App.Path & "\cut.ico"))
        Set clsNewImage = .Add(, "Image4", _
            LoadPicture(App.Path & "\copy.ico"))
        Set clsNewImage = .Add(, "Image5", _
            LoadPicture(App.Path & "\paste.ico"))
        Set clsNewImage = .Add(, "Image6", _
            LoadPicture(App.Path & "\emails.ico"))
        Set clsNewImage = .Add(, "Image7", _
            LoadPicture(App.Path & "\help2.ico"))
    End With

    '---------------------------
    ' Add toolbar buttons and ToolTipText
    '---------------------------

        Set Me.tbrTools.ImageList = Me.imlPics

        With tbrTools.Buttons
            .Clear

            Set clsNewButton = .Add(1, "open", , , "Image1")
                clsNewButton.ToolTipText = "Open a File"
            Set clsNewButton = .Add(2, "newfile", , , "Image2")
                clsNewButton.ToolTipText = "New File"
            Set clsNewButton = .Add(3, "cut", , , "Image3")
                clsNewButton.ToolTipText = "Cut"
            Set clsNewButton = .Add(4, "copy", , , "Image4")
                clsNewButton.ToolTipText = "Copy"
            Set clsNewButton = .Add(5, "paste", , , "Image5")
                clsNewButton.ToolTipText = "Paste"
            Set clsNewButton = .Add(6, "emails", , , "Image6")
                clsNewButton.ToolTipText = "Emails"
            Set clsNewButton = .Add(7, "help", , , "Image7")
                clsNewButton.ToolTipText = "Get Help"
        End With
    
End Sub
Private Sub SaveToolbarSetting(ByVal intToolNum As Integer)
' Save toolbar setting to record #4 in random access file

Dim udtUserOptions As OptionsStruc
Dim intFile As Integer  ' Free file number

Select Case intToolNum
    Case 1
        udtUserOptions.intToolBar = 1   'full-sized
        lblToolbar.Caption = "Toolbar Size = Full-Sized"
    Case 2
        udtUserOptions.intToolBar = 2   'small
        lblToolbar.Caption = "Toolbar Size = Small"
    Case 0
        udtUserOptions.intToolBar = 0   ' no toolbar
        lblToolbar.Caption = "Toolbar Off"
End Select

intFile = FreeFile
Open App.Path & "\options.dat" For Random As intFile Len = Len(udtUserOptions)

Put #intFile, 4, udtUserOptions.intToolBar

Close #intFile

End Sub

Private Sub AdjustOptionsButton()
    cmdOptions.Top = frmMain.Height - cmdOptions.Height - 435
End Sub
