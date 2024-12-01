Option Compare Database

'------------------------------------------------------------
' godisnikM_41
'
'------------------------------------------------------------
Function godisnikM_41()
On Error GoTo godisnikM_41_Err

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "godisnikTQ_4_1", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_2", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_3", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_4", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_5", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_6", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_7", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_8", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_4_9", acViewNormal, acEdit
    DoCmd.OpenReport "godisnikReport4_p", acViewPreview, "", ""


godisnikM_41_Exit:
    Exit Function

godisnikM_41_Err:
    MsgBox Error$
    Resume godisnikM_41_Exit

End Function