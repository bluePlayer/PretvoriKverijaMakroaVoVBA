Option Compare Database

'------------------------------------------------------------
' godisnikM_10
'
'------------------------------------------------------------
Function godisnikM_10()
On Error GoTo godisnikM_10_Err

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "godisnikTQ_10_1", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_10_2", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_10_3", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_10_4", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_10_5", acViewNormal, acEdit
    DoCmd.OpenReport "godisnikReport10", acViewPreview, "", ""


godisnikM_10_Exit:
    Exit Function

godisnikM_10_Err:
    MsgBox Error$
    Resume godisnikM_10_Exit

End Function