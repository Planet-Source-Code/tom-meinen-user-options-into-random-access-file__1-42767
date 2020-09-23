Attribute VB_Name = "modX"
Option Explicit

Type OptionsStruc
' put other variables in here as udt, see if dupe variables not needed
    strName As String * 40
    intNumCreate As Integer
    blnShowSplash As Boolean
    blnFormatHTML As Boolean
    intToolBar As Integer   ' 1 = full-sized, 2 = small, 0 = none
End Type

Sub main()
' find out if splash screen should be shown, and either
'show it or go on to frmMain

Dim intFileNumber As Integer
Dim udtUserOptions As OptionsStruc
intFileNumber = FreeFile
Open App.Path & "\options.dat" For Random As intFileNumber _
Len = Len(udtUserOptions)

Get #intFileNumber, 2, udtUserOptions.blnShowSplash

If udtUserOptions.blnShowSplash = True Then
    frmSplash.Show
Else
    frmMain.Show
End If

Close #intFileNumber

End Sub



Public Sub PopUpCheck(ByVal intToolNum As Integer)
' make sure correct item on right click popup is checked
Select Case intToolNum
    Case 1
        frmMain.mnuPopUpFullTools.Checked = True    '<--- right click popup menu item checked
        frmMain.mnuPopUpSmallTools.Checked = False
        frmMain.mnuPopUpNoTools.Checked = False
        frmMain.tbrTools.Visible = True
    Case 2
        frmMain.mnuPopUpFullTools.Checked = False
        frmMain.mnuPopUpSmallTools.Checked = True   '<--- right click popup menu item checked
        frmMain.mnuPopUpNoTools.Checked = False
        frmMain.tbrTools.Visible = True
    Case 0
        frmMain.mnuPopUpFullTools.Checked = False
        frmMain.mnuPopUpSmallTools.Checked = False
        frmMain.mnuPopUpNoTools.Checked = True      '<--- right click popup menu item checked
        frmMain.tbrTools.Visible = False
End Select

End Sub
