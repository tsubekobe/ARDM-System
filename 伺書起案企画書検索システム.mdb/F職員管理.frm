Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdíœ_Click()
On Error GoTo Err_cmdíœ_Click

    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdíœ_Click:
    Exit Sub

Err_cmdíœ_Click:
    MsgBox Err.Description
    Resume Exit_cmdíœ_Click
    
End Sub

Private Sub cmd–ß‚é_Click()

    DoCmd.OpenForm "Fƒƒjƒ…["
    DoCmd.Close acForm, "FEˆõŠÇ—", acSaveNo

End Sub

Private Sub Form_Load()
    
    Me.Caption = cstSys

End Sub