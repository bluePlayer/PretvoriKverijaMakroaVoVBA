Option Compare Database

'------------------------------------------------------------
' godisnikM_7
'
'------------------------------------------------------------
Function godisnikM_7()
On Error GoTo godisnikM_7_Err

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "godisnikTQ_7_1", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_2", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_3", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_4", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_5", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_6", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_7", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_8", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_7_9", acViewNormal, acEdit
    DoCmd.OpenReport "godisnikReport7", acViewPreview, "", ""


godisnikM_7_Exit:
    Exit Function

godisnikM_7_Err:
    MsgBox Error$
    Resume godisnikM_7_Exit

End Function