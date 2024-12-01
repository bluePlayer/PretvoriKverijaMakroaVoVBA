Option Compare Database

'------------------------------------------------------------
' godisnikM_1
'
'------------------------------------------------------------
Function godisnikM_1()
On Error GoTo godisnikM_1_Err

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "godisnikTQ_1_1", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_1_2", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_1_3", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_1_4", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_1_5", acViewNormal, acEdit
    DoCmd.OpenReport "godisnikReport1", acViewPreview, "", ""


godisnikM_1_Exit:
    Exit Function

godisnikM_1_Err:
    MsgBox Error$
    Resume godisnikM_1_Exit

End Function