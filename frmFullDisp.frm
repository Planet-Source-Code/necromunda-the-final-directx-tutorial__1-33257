VERSION 5.00
Begin VB.Form frmFullDisp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   4080
   ClientTop       =   2370
   ClientWidth     =   4545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4545
End
Attribute VB_Name = "frmFullDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the form where we display in full-screen mode.
'The coding below checks for the user pressing the
'escape key.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    With frmMain
    
        Select Case KeyCode
            Case Is = vbKeyEscape
                Call unloadObjects
            Case Is = vbKeyR
                .chkRotate.Value = IIf(.chkRotate.Value = 1, 0, 1)
            Case Is = vbKeyS
                .chkScale.Value = IIf(.chkScale.Value = 1, 0, 1)
            Case Is = vbKeyT
                .chkMove.Value = IIf(.chkMove.Value = 1, 0, 1)
            Case Is = vbKeyX
                .chkX.Value = IIf(.chkX.Value = 1, 0, 1)
            Case Is = vbKeyY
                .chkY.Value = IIf(.chkY.Value = 1, 0, 1)
            Case Is = vbKeyZ
                .chkZ.Value = IIf(.chkZ.Value = 1, 0, 1)
        End Select
    
    End With

End Sub

