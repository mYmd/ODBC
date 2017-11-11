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

'VBA����
'CreateSymbolicLinkA("H:\Projects\LIB\mapM.dll", "H:\Projects\VC\ThreadTest\x64\Release\mapM.dll", 0)
'���[�N�V�[�g�ォ��
'CreateSymbolicLinkW("H:\Projects\LIB\mapM.dll", "H:\Projects\VC\ThreadTest\x64\Release\mapM.dll", 0)

'***********************************************************************************
'   toF = "H:\Bunsho\Info\Others\tmp\"
'   fromF = "H:\Bunsho\Info\Others\���Z�g\"
'   ffiles = headN(selectCol(getFileFolderList(fromF), 0), 10)
'   printM zipWith(p_CreateSymbolicLink, mapF(p_plus(toF), ffiles), mapF(p_plus(fromF), ffiles))
'***********************************************************************************

' �V���{���b�N�����N�̍쐬
Function CreateSymbolicLink(ByRef toName As Variant, ByRef fromName As Variant) As Variant
    CreateSymbolicLink = CreateSymbolicLinkA(toName, fromName, 0)
End Function
    Function p_CreateSymbolicLink(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CreateSymbolicLink = make_funPointer(AddressOf CreateSymbolicLink, firstParam, secondParam)
    End Function

' �n�[�h�����N�̍쐬
Function CreateHardLink(ByRef toName As Variant, ByRef fromName As Variant) As Variant
    CreateHardLink = CreateHardLinkA(toName, fromName, 0)
End Function
    Function p_CreateHardLink(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CreateHardLink = make_funPointer(AddressOf CreateHardLink, firstParam, secondParam)
    End Function

' �ڑ�����Ă���h���C�u���^�[�̎擾
Public Function GetLogicalDrives() As String
    Dim d As Long:  d = GetLogicalDrives_imple()
    Dim i As Long:  i = 0
    Do While d > 0
        If d Mod 2 = 1 Then
            GetLogicalDrives = GetLogicalDrives & VBA.Chr(VBA.Asc("A") + i)
        End If
        d = d \ 2
        i = i + 1
    Loop
End Function

' �t�@�C���ꗗ
' �t�H���_�����݂��Ȃ��Ƃ���Empty, �t�H���_����̎���Array()
Function getFileFolderList(ByVal folderName As String, _
                  Optional ByVal files_only As Boolean = True) As Variant
    Dim fso As Object
    Dim fDer As Object
    Dim filesCollection As Object
    Dim z As Variant
    Dim ret As Variant
    Dim i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderName) Then Exit Function
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
        getFileFolderList = catC(subM(ret, reverse(sortIndex(ret, Array(1)))), transpose(iota(1, rowSize(ret))))
    End If
    Set filesCollection = Nothing
    Set fDer = Nothing
    Set fso = Nothing
End Function

'�t�@�C���폜�i�Ώۃt�H���_���A�t�@�C����[�z��]�j
Function killFiles(ByRef folderName As Variant, _
                   ByVal fileNames As Variant, _
          Optional ByVal force As Boolean = False) As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderName) Then Exit Function
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

'�t�@�C���R�s�[�i�R�s�[���t�H���_���A�t�@�C����[�z��]�A�R�s�[��t�H���_���j
Function copyFile(ByVal sourceFolder As String, _
                  ByVal fileNames As Variant, _
                  ByVal targetFolder As String, _
         Optional ByVal overwrite As Boolean = False) As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sourceFolder) Then Exit Function
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
        If fso.FileExists(sourceFile) And fso.FolderExists(targetFolder) Then
            If overwrite Or Not fso.FileExists(fso.BuildPath(targetFolder, fso.GetFileName(sourceFile))) Then
                fso.copyFile sourceFile, targetFolder
                counter = counter + 1
            End If
        End If
    End Sub
