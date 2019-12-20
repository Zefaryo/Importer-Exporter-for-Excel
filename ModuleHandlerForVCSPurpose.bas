Attribute VB_Name = "ModuleHandlerForVCSPurpose"
'''PLEASE READ CAREFULLY

'''BEE WEAR Very very very dangerous sub-routine for a bad joke in the end of this module
'''Proceed with care or remove it

'''TUTORIAL:
'''Go to tools reference and activate "Microsoft VBA extensibility 5.3" and "Microsoft scripting runtime"
'''Go to options-> trust center->trust center settings->macro settings-> check "trust acces to the vba object model"
''' Change global consts to correct directory and folder name that is or will be created in this directory
''' This workbook must be the active one in Excel.
''' Use sub-routine exportmodules or importmodules to export/import to your liking
'''IMPORTANT: when reimporting, all modules will be deleted befor importing to import cleanly (without duplicates)
'''IMPORTANT: When exporting, files on your disk will be overriden



Option Explicit

Global Const Path As String = "C:\YourPath" 'Path to directory wher Export/Import folder will be created
Global Const FolderName As String = "FolderName" 'Target folder name used for Export/Import
Global Const CleanModulesOnImport As Boolean = True ' Wether or not to delete all the files in your VBA environment when you import files

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim sSourceWorkbook As String
    Dim sExportPath As String
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent


    ''' This sub creates a folder named after the name you put just above created in the directory you defined above
    ''' and export what needs to be in it. Trust me.
    ''' (In fact you should'nt, don't take me responsible for damages you caused by using it)
    ''' Note to self : create a "Are you sure dialogue", and make it less destructive/safer to use : human error proof
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Sub
    End If

    ''' This workbook must be the active one in Excel.
    sSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(sSourceWorkbook)
    
    sExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        sFileName = Replace(cmpComponent.Name, "_", "\")

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                sFileName = sFileName & ".cls"
            Case vbext_ct_MSForm
                sFileName = sFileName & ".frm"
            Case vbext_ct_StdModule
                sFileName = sFileName & ".bas"
            Case vbext_ct_Document
                ''' Note to self: This is a worksheet or workbook object. Don't try to export it. Never. Ever.
                bExport = False
        End Select
        
        If bExport And cmpComponent.Name <> "ModuleHandlerForVCSPurpose" Then
            ''' Export the component to a text file.
            cmpComponent.Export sExportPath & sFileName
            
        ''' remove it from the project if you want by uncommenting following line
        'wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export done"
End Sub

Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim sTargetWorkbook As String
    Dim sImportPath As String
    Dim sFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' This workbook must be the active one in Excel.
    sTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(sTargetWorkbook)

    '''Path where the code modules are/should be located.
    sImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(sImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       ' Exit Sub
    End If

    'Delete all modules/Userforms from the (active)Workbook
    If CleanModulesOnImport Then
        Call DeleteVBAModulesAndUserForms
    End If

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ImportFilesFromDirectory objFSO, objFSO.GetFolder(sImportPath), cmpComponents, ""
    MsgBox "Import done"
End Sub

Private Sub ImportFilesFromDirectory(objFSO As FileSystemObject, sFolder As Scripting.Folder, cmpComponents As VBIDE.VBComponents, sModulePrefix As String)
    Dim objFile As Scripting.File
    Dim objFolder As Scripting.Folder
    ' MsgBox "Importing files from " & sFolder.Path
    ''' Import all the code modules in the specified path to the ActiveWorkbook.
    For Each objFile In sFolder.Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            Dim LastIndex As Integer
            LastIndex = cmpComponents.Count + 1
            cmpComponents.Import objFile.Path
            If cmpComponents.Item(LastIndex).Name = objFSO.GetBaseName(objFile.Name) Then
                With cmpComponents.Item(LastIndex)
                    .Name = sModulePrefix & objFSO.GetBaseName(objFile.Name)
                End With
            End If
        End If
        
    Next objFile

    For Each objFolder in sFolder.SubFolders
        If Not InStr(1, objFolder.Name, ".") = 1 Then
            ImportFilesFromDirectory objFSO, objFolder, cmpComponents, sModulePrefix & objFolder.Name & "_"
        End If
    Next objFolder
End Sub

Function FolderWithVBAProjectFiles() As String
    
    Dim TargetPath As String
    Dim FSO As Object

    Set FSO = CreateObject("scripting.filesystemobject")
    
    TargetPath = Path
    
    If Right(TargetPath, 1) <> "\" Then
        TargetPath = Path & "\"
    End If
    
    If FSO.FolderExists(TargetPath & FolderName) = False Then
        On Error Resume Next
        MkDir TargetPath & FolderName
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(TargetPath & FolderName) = True Then
        FolderWithVBAProjectFiles = TargetPath & FolderName
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'We do nothing. And it's bette like this I assure you
                'Else if you want some scare uncomment next line ;)
            ElseIf VBComp.Name = "ModuleHandlerForVCSPurpose" Then
                ' Let's not suicide :)
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
