Attribute VB_Name = "MFileOpen"
'========================================================================================================
' IMPORTANT:
'
' This BAS code was written by Paul Mather,
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
'
' However, I removed all the uneeded parts (fonts, colours, save dialog, etc).
'========================================================================================================

Option Explicit

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_EXPLORER = &H80000
'Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
'Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public FileDialog As OPENFILENAME
Dim ParenthWnd As Long

Public Function ShowOpen(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile

    Dim ret As Long
    Dim Count As Integer
    Dim fileNameHolder As String
    Dim LastCharacter As Integer
    Dim NewCharacter As Integer
    Dim tempFiles(1 To 200) As String
    Dim hInst As Long
    Dim Thread As Long
    
    ParenthWnd = hWnd
    With FileDialog
        .nStructSize = Len(FileDialog)
        .hwndOwner = hWnd
        .sFileTitle = Space$(2048)
        .nTitleSize = Len(FileDialog.sFileTitle)
        .sFile = FileDialog.sFile & Space$(2047) & Chr$(0)
        .nFileSize = Len(FileDialog.sFile)
        .flags = OFS_FILE_OPEN_FLAGS
    End With
    
    ret = GetOpenFileName(FileDialog)

    If ret Then
        
        If Trim$(FileDialog.sFileTitle) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(Count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next
        Else
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    
    Else
        
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    
    End If
    
End Function
