Attribute VB_Name = "Y_symbolicLink"
'Y_symbolicLink
'Copyright (c) 2015 mmYYmmdd
Option Explicit

Public Declare PtrSafe Function CreateSymbolicLinkW Lib "kernel32.dll" ( _
                        ByVal toName As String, _
                    ByVal fromName As String, _
                ByVal file_folder As Long) As Byte          'file:0 ,  folder:1

Public Declare PtrSafe Function CreateSymbolicLinkA Lib "kernel32.dll" ( _
                        ByVal toName As String, _
                    ByVal fromName As String, _
                ByVal file_folder As Long) As Byte          'file:0 ,  folder:1

Public Declare PtrSafe Function CreateHardLinkW Lib "kernel32.dll" ( _
                        ByVal toName As String, _
                    ByVal fromName As String, _
                ByVal attr As Long) As Byte

Public Declare PtrSafe Function CreateHardLinkA Lib "kernel32.dll" ( _
                        ByVal toName As String, _
                    ByVal fromName As String, _
                ByVal attr As Long) As Byte

Public Declare PtrSafe Function IsUserAnAdmin Lib "Shell32.dll" () As Byte

Public Declare PtrSafe Function GetLogicalDrives_imple Lib "kernel32.dll" Alias "GetLogicalDrives" () As Long

'VBAから
'CreateSymbolicLinkA("H:\Projects\LIB\mapM.dll", "H:\Projects\VC\ThreadTest\x64\Release\mapM.dll", 0)
'ワークシート上から
'CreateSymbolicLinkW("H:\Projects\LIB\mapM.dll", "H:\Projects\VC\ThreadTest\x64\Release\mapM.dll", 0)

'***********************************************************************************
'   toF = "H:\Bunsho\Info\Others\tmp\"
'   fromF = "H:\Bunsho\Info\Others\元住吉\"
'   ffiles = headN(selectCol(getFileFolderList(fromF), 0), 10)
'   printM zipWith(p_CreateSymbolicLink, mapF(p_plus(toF), ffiles), mapF(p_plus(fromF), ffiles))
'***********************************************************************************

' シンボリックリンクの作成
Function CreateSymbolicLink(ByRef toName As Variant, ByRef fromName As Variant) As Variant
    CreateSymbolicLink = CreateSymbolicLinkA(toName, fromName, 0)
End Function
    Function p_CreateSymbolicLink(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CreateSymbolicLink = make_funPointer(AddressOf CreateSymbolicLink, firstParam, secondParam)
    End Function

' ハードリンクの作成
Function CreateHardLink(ByRef toName As Variant, ByRef fromName As Variant) As Variant
    CreateHardLink = CreateHardLinkA(toName, fromName, 0)
End Function
    Function p_CreateHardLink(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CreateHardLink = make_funPointer(AddressOf CreateHardLink, firstParam, secondParam)
    End Function

' 接続されているドライブレターの取得
Public Function GetLogicalDrives() As String
    Dim fso As Object, dc As Object, d As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dc = fso.Drives
    For Each d In dc
        GetLogicalDrives = GetLogicalDrives & d.DriveLetter
    Next
End Function

' ファイル一覧
' フォルダが存在しないときはEmpty, フォルダが空の時はArray()
Function getFileFolderList(ByVal folderName As String, _
                  Optional ByVal files_only As Boolean = True) As Variant
    Dim fso As Object
    Dim fDer As Object
    Dim filesCollection As Object
    Dim z As Variant
    Dim ret As Variant
    Dim i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(folderName) Then Exit Function
    Set fDer = fso.GetFolder(folderName)
    If files_only Then
        Set filesCollection = fDer.files
    Else
        Set filesCollection = fDer.SubFolders
    End If
    ret = Array()
    getFileFolderList = ret
    If 0 < filesCollection.Count Then
        ReDim ret(0 To filesCollection.Count - 1, 0 To 1)
        i = LBound(ret)
        For Each z In filesCollection
            ret(i, 0) = z.name
            ret(i, 1) = z.DateLastModified 'DateCreated
            i = i + 1
        Next z
        getFileFolderList = catC(subM(ret, reverse(sortIndex(ret, Array(1)))), iota(1, rowSize(ret)))
    End If
    Set filesCollection = Nothing
    Set fDer = Nothing
    Set fso = Nothing
End Function

'ファイル削除（対象フォルダ名、ファイル名[配列]）
Function killFiles(ByRef folderName As Variant, _
                   ByVal fileNames As Variant, _
          Optional ByVal force As Boolean = False) As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(folderName) Then Exit Function
    If Not IsArray(fileNames) Then fileNames = VBA.Array(fileNames)
    Dim ret As Variant:     ret = repeat(0, sizeof(fileNames))
    Dim fileName As Variant
    Dim k As Long:    k = 0
    For Each fileName In fileNames
        Call killFile_imple(fso, fso.BuildPath(folderName, fileName), ret(k), force)
        k = k + 1
    Next fileName
    Set fso = Nothing
    swapVariant killFiles, ret
End Function
    
    Private Sub killFile_imple(ByRef fso As Object, _
                               ByRef fileName As String, _
                               ByRef i As Variant, _
                               ByVal force As Boolean)
        If fso.FileExists(fileName) Then
            fso.DeleteFile fileName, force
            If Not fso.FileExists(fileName) Then i = 1
        End If
    End Sub

'ファイルコピー（コピー元フォルダ名、ファイル名[配列]、コピー先フォルダ名）
Function copyFile(ByVal sourceFolder As String, _
                  ByVal fileNames As Variant, _
                  ByVal targetFolder As String, _
         Optional ByVal overwrite As Boolean = False) As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(sourceFolder) Then Exit Function
    If Not IsArray(fileNames) Then fileNames = VBA.Array(fileNames)
    Dim k As Long:          k = 0
    Dim ret As Variant:     ret = repeat(0, sizeof(fileNames))
    Dim fileName As Variant
    For Each fileName In fileNames
        Call copyFile_imple(fso, _
                            fso.BuildPath(sourceFolder, fileName), _
                            targetFolder, _
                            ret(k), _
                            overwrite)
        k = k + 1
    Next fileName
    Set fso = Nothing
    swapVariant copyFile, ret
End Function
    
    Private Sub copyFile_imple(ByRef fso As Object, _
                               ByRef sourceFile As String, _
                               ByVal targetFolder As String, _
                               ByRef counter As Variant, _
                               ByVal overwrite As Boolean)
        targetFolder = fso.GetParentFolderName(fso.BuildPath(targetFolder, "_")) & "\"
        If fso.FileExists(sourceFile) And fso.folderExists(targetFolder) Then
            If overwrite Or Not fso.FileExists(fso.BuildPath(targetFolder, fso.GetFileName(sourceFile))) Then
                fso.copyFile sourceFile, targetFolder
                counter = counter + 1
            End If
        End If
    End Sub

' フォルダ作成
Function createFolder(ByVal folderName As String) As Boolean
    Dim fso As Object, DriveName As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    DriveName = fso.GetDriveName(folderName)
    If Not fso.DriveExists(DriveName) Then
        createFolder = False
    ElseIf fso.GetDrive(DriveName).IsReady Then
        createFolder = createFolder_imple(fso, folderName)
    Else
        createFolder = False
    End If
    Set fso = Nothing
End Function

    Private Function createFolder_imple(ByVal fso As Object, _
                                    ByVal folderName As String) As Boolean
        If fso.folderExists(folderName) Then
            createFolder_imple = True
        Else
            If createFolder_imple(fso, fso.GetParentFolderName(folderName)) Then
                fso.createFolder folderName
                createFolder_imple = True
            Else    ' ありえないはず
                createFolder_imple = False
            End If
        End If
    End Function


