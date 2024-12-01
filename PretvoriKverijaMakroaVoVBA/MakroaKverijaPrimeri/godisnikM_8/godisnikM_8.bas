Option Compare Database

'------------------------------------------------------------
' godisnikM_8
'
'------------------------------------------------------------
Function godisnikM_8()
On Error GoTo godisnikM_8_Err

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "godisnikTQ_8_1", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_2", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_3", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_4", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_5", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_6", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_7", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_8", acViewNormal, acEdit
    DoCmd.OpenQuery "godisnikTQ_8_9", acViewNormal, acEdit
    DoCmd.OpenReport "godisnikReport8", acViewPreview, "", ""


godisnikM_8_Exit:
    Exit Function

godisnikM_8_Err:
    MsgBox Error$
    Resume godisnikM_8_Exit

End Function