Public Function ExcelJetSQLStringSoPateka(ByVal imeProverka As String, ByVal patekaXLSFajl As String, ByVal imeXLSFajl As String) As String
    If Right(patekaXLSFajl, 1) = "\" Then
        ExcelJetSQLStringSoPateka = " INTO [" & imeProverka & "] IN ''[Excel 8.0;Database=" & patekaXLSFajl & imeXLSFajl & "]"
    Else
        ExcelJetSQLStringSoPateka = " INTO [" & imeProverka & "] IN ''[Excel 8.0;Database=" & patekaXLSFajl & "\" & imeXLSFajl & "]"
    End If
End Function

Public Function ExcelJetSQLString(ByVal imeProverka As String, ByVal imeXLSFajl As String) As String
    ExcelJetSQLString = " INTO [" & imeProverka & "] IN ''[Excel 8.0;Database=" & CurrentProject.Path & "\" & imeXLSFajl & "]"
End Function

Public Sub IzveziKverijaIMakroaVoModuli(ByVal pateka As String)
    Dim obj As AccessObject
    Dim moduleName, strTmp As String
    Dim makroFajl, oFolder
    Dim fs As Object
    Dim txt As Object
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Const ForReading = 1, TristateUseDefault = -2
    
    On Error GoTo Error_OpenOptionsDialog
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb

    For Each obj In CurrentProject.AllMacros
    
        moduleName = Konstanti.MAKRO_PRETSTAVKA_ZA_IZVEZUVANJE & obj.Name
        
        If Not DaliPostoiModulot(moduleName) Then
            DoCmd.SelectObject acMacro, obj.Name, True
            DoCmd.RunCommand acCmdConvertMacrosToVisualBasic
            
            DoCmd.Save acModule, moduleName
        End If
        
        CreateFolderIfNotExists (pateka & obj.Name)
        
        DoCmd.SelectObject acModule, moduleName, True
        Application.SaveAsText acModule, moduleName, pateka & obj.Name & "\" & obj.Name & ".bas"

        For Each qdf In db.QueryDefs
        
            Set makroFajl = fs.OpenTextFile(pateka & obj.Name & "\" & obj.Name & ".bas", ForReading, True, TristateUseDefault)
            
            Do While makroFajl.AtEndOfStream <> True
                strTmp = makroFajl.ReadLine
                
                If Len(strTmp) > 0 Then
                
                    If InStr(1, strTmp, "OpenQuery", vbTextCompare) > 0 Then
                        
                        
                        If InStr(1, strTmp, qdf.Name, vbTextCompare) > 0 Then
                            Set txt = fs.CreateTextFile(pateka & obj.Name & "\" & qdf.Name & ".sql", True)
                            txt.WriteLine qdf.sql
                            txt.Close
                        End If
                    End If
                End If
            Loop
            
        Next qdf
        
        makroFajl.Close
    
    Next obj
    
    ' Display a message box when the export is complete
    MsgBox "Export complete!", vbInformation, "Export Query SQL"
    
Exit_OpenOptionsDialog:
    Exit Sub

Error_OpenOptionsDialog:
    MsgBox Err & ": " & Err.Description
    Resume Exit_OpenOptionsDialog
End Sub

Public Sub ExportQuerySQL(ByVal strPath As String)
On Error GoTo Greshka

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    ' Open the current database
    Set db = CurrentDb
    
    ' TODO ovoj kod kao naogja vrska izmegju kveri i makro, samo shto ne raboti
    ' Najverojatno radi slednovo:
    ' Access does not search Visual Basic for Applications (VBA) code, macros, or data access pages for dependencies.
    ' https://learn.microsoft.com/en-us/office/vba/api/access.accessobject.isdependentupon
    
    'Dim vezhba As AccessObject, dbs As Object
    'Set dbs = Application.CurrentData
    ' Search for open AccessObject objects in AllQueries collection.
    'For Each vezhba In dbs.AllQueries
        'If vezhba.IsLoaded = True Then
            ' Print name of obj.
            'If vezhba.IsDependentUpon(acMacro, "godisnikM_10") = True Then
                'MsgBox vezhba.Name & " zavisi od makroto godisnikM_10"
            'End If
            
            'Debug.Print vezhba.Name
        'End If
    'Next vezhba
   
    ' Loop through all queries and export the SQL code to text files
    For Each qdf In db.QueryDefs
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Dim txt As Object
        Set txt = fs.CreateTextFile(strPath & qdf.Name & ".sql", True)
        txt.WriteLine qdf.sql
        txt.Close
    Next qdf
   
    ' Display a message box when the export is complete
    MsgBox "Export complete!", vbInformation, "Export Query SQL"
    
izlez:
    Exit Sub

Greshka:
    MsgBox Error$
    Resume izlez

End Sub

Public Function DaliPostoiModulot(ByVal moduleName As String) As Boolean
    Dim existingModule As AccessObject
    Dim postoi As Boolean
    
    postoi = False
    
    For Each existingModule In CurrentProject.AllModules
        If existingModule.Name = moduleName Then
            postoi = True
        End If
    Next existingModule
    
    DaliPostoiModulot = postoi
End Function

Public Function daliTabelataPostoi(ByVal imeNaTabela As String, ByRef db As DAO.Database) As Boolean
    Dim msg As String
    Dim postoi As Boolean
    
    postoi = False

    For Each tbl In db.TableDefs
        If tbl.Name = imeNaTabela Then
            postoi = True
            GoTo zavrshiv
        End If
    
    Next tbl
    
zavrshiv:
    daliTabelataPostoi = postoi
    
End Function

Public Sub PretvoriMakroaVoModuliISnimiNaDisk(ByVal pateka As String, Optional ByVal stvoriPapkiZaMakroa As Boolean = False)
    Dim obj As AccessObject
    Dim moduleName As String
    Dim fs As Object
    Dim txt As Object
    
    On Error GoTo Error_OpenOptionsDialog

    For Each obj In CurrentProject.AllMacros
    
        moduleName = Konstanti.MAKRO_PRETSTAVKA_ZA_IZVEZUVANJE & obj.Name
        
        ' TODO da se dodade kod za pravenje papka so imeto na modulot/makroto kade potoa bi se smestile kverijata
        ' koi se povikuvaat vo modulot/makroto
        ' so koristenje na CreateFolderIfNotExists()
        'If obj.IsDependentUpon(acQuery, "godisnikTQ_10_1") = True Then
        '    MsgBox obj.Name & " zavisi od kverito godisnikTQ_10_1"
        'End If
        
        If Not DaliPostoiModulot(moduleName) Then
            DoCmd.SelectObject acMacro, obj.Name, True
            DoCmd.RunCommand acCmdConvertMacrosToVisualBasic
            
            DoCmd.Save acModule, moduleName
        End If
        
        DoCmd.SelectObject acModule, moduleName, True
        Application.SaveAsText acModule, moduleName, pateka & obj.Name & ".bas"
    
    Next obj
    
Exit_OpenOptionsDialog:
    Exit Sub

Error_OpenOptionsDialog:
    MsgBox Err & ": " & Err.Description
    Resume Exit_OpenOptionsDialog
End Sub

Public Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Public Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub

Public Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Public Function CreateFolderIfNotExists(Directory As String)
    Dim Exists As Boolean
    Exists = FolderExists(Directory)
    If (Exists = False) Then
        MkDir Directory
    End If
End Function

