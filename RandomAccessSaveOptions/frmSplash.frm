VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   1800
      Left            =   6000
      Top             =   360
   End
   Begin VB.Frame fraSplash 
      BackColor       =   &H8000000A&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2055
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Top             =   795
         Width           =   2775
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Splash"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3360
         TabIndex        =   1
         Top             =   1560
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
End Sub



Private Sub fraSplash_DragDrop(Source As Control, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub tmrSplash_Timer()
    frmMain.Show
    Unload frmSplash
    Screen.MousePointer = vbDefault
End Sub
