Attribute VB_Name = "FileDlg"
'Implemented using API file open and save dialog
'Usage
'Text1 = FileDialog(Me, False, "Please select a file", "Text file (*.txt)|*.txt|All files (*.*)|*.*")
'Text2 = FileDialog(Me, True, "Save file to", "*.exe", FileName & "360safebox")
'Text3 = GetFolderName(me.hWnd, "Please select a directory")

Option Explicit
Public Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type
Public Type BROWSEINFO
    hwndOwner         As Long
    pidlRoot          As Long
    pszDisplayName    As Long
    lpszTitle         As Long
    ulFlags           As Long
    lpfnCallback      As Long
    lParam            As Long
    iImage            As Long
End Type

Public Const OFN_READONLY             As Long = &H1
Public Const OFN_OVERWRITEPROMPT      As Long = &H2
Public Const OFN_HIDEREADONLY         As Long = &H4
Public Const OFN_NOCHANGEDIR          As Long = &H8
Public Const OFN_SHOWHELP             As Long = &H10
Public Const OFN_ENABLEHOOK           As Long = &H20
Public Const OFN_ENABLETEMPLATE       As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_NOVALIDATE           As Long = &H100
Public Const OFN_ALLOWMULTISELECT     As Long = &H200
Public Const OFN_EXTENSIONDIFFERENT   As Long = &H400
Public Const OFN_PATHMUSTEXIST        As Long = &H800
Public Const OFN_FILEMUSTEXIST        As Long = &H1000
Public Const OFN_CREATEPROMPT         As Long = &H2000
Public Const OFN_SHAREAWARE           As Long = &H4000
Public Const OFN_NOREADONLYRETURN     As Long = &H8000
Public Const OFN_NOTESTFILECREATE     As Long = &H10000
Public Const OFN_NONETWORKBUTTON      As Long = &H20000
Public Const OFN_NOLONGNAMES          As Long = &H40000
Public Const OFN_EXPLORER             As Long = &H80000
Public Const OFN_NODEREFERENCELINKS   As Long = &H100000
Public Const OFN_LONGNAMES            As Long = &H200000

Public Const OFN_SHAREFALLTHROUGH     As Long = 2
Public Const OFN_SHARENOWARN          As Long = 1
Public Const OFN_SHAREWARN            As Long = 0


Public Const BrowseForFolders         As Long = &H1
Public Const BrowseForComputers       As Long = &H1000
Public Const BrowseForPrinters        As Long = &H2000
Public Const BrowseForEverything      As Long = &H4000

Public Const CSIDL_BITBUCKET          As Long = 10
Public Const CSIDL_CONTROLS           As Long = 3
Public Const CSIDL_DESKTOP            As Long = 0
Public Const CSIDL_DRIVES             As Long = 17
Public Const CSIDL_FONTS              As Long = 20
Public Const CSIDL_NETHOOD            As Long = 18
Public Const CSIDL_NETWORK            As Long = 19
Public Const CSIDL_PERSONAL           As Long = 5
Public Const CSIDL_PRINTERS           As Long = 4
Public Const CSIDL_PROGRAMS           As Long = 2
Public Const CSIDL_RECENT             As Long = 8
Public Const CSIDL_SENDTO             As Long = 9
Public Const CSIDL_STARTMENU          As Long = 11
'Download by http://www.NewXing.com
Public Const MAX_PATH                 As Long = 260
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Const BIF_RETURNONLYFSDIRS = &H1                                        'Browse for folders
Private Const BIF_NEWDIALOGSTYLE = &H40                                         'New style (with new folder button, resizable dialog)
Private Const BIF_NONEWFOLDERBUTTON = &H200                                     'New style, no new folder button (resizable only)

Public Function FileDialog(FormObject As Form, SaveDialog As Boolean, ByVal Title As String, ByVal Filter As String, Optional ByVal FileName As String, Optional ByVal Extention As String, Optional ByVal InitDir As String, Optional bModal As Boolean = True) As String
    Dim OFN   As OPENFILENAME
    Dim r     As Long
    
    If Len(FileName) > MAX_PATH Then Call MsgBox("Filename Length Overflow", vbExclamation, App.Title + " - FileDialog Function"): Exit Function
    FileName = FileName + String(MAX_PATH - Len(FileName), 0)
    
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = IIf(bModal, FormObject.hWnd, 0)
        .hInstance = App.hInstance
        .lpstrFilter = Replace(Filter, "|", vbNullChar)
        .lpstrFile = FileName
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space$(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
        .lpstrDefExt = Extention
    End With
    
    Dim L As Long
    L = GetTickCount
    
    If SaveDialog Then r = GetSaveFileName(OFN) Else r = GetOpenFileName(OFN)
    
    If GetTickCount - L < 20 Then
        OFN.lpstrFile = ""
        If SaveDialog Then r = GetSaveFileName(OFN) Else r = GetOpenFileName(OFN)
    End If
    
    If r = 1 Then FileDialog = Left$(OFN.lpstrFile, InStr(1, OFN.lpstrFile + vbNullChar, vbNullChar) - 1)
End Function
Public Function BrowseFolders(FormObject As Form, sMessage As String) As String
    Dim b As BROWSEINFO
    Dim r As Long
    Dim L As Long
    Dim f As String
    
    FormObject.Enabled = False
    With b
        .hwndOwner = FormObject.hWnd
        .lpszTitle = lstrcat(sMessage, "")
        .ulFlags = BrowseForFolders
    End With
    
    SHGetSpecialFolderLocation FormObject.hWnd, CSIDL_DRIVES, b.pidlRoot
    r = SHBrowseForFolder(b)
    
    If r <> 0 Then
        f = String(MAX_PATH, vbNullChar)
        SHGetPathFromIDList r, f
        CoTaskMemFree r
        L = InStr(1, f, vbNullChar) - 1
        If L < 0 Then L = 0
        f = Left(f, L)
        AddSlash f
    End If
    
    BrowseFolders = f
    FormObject.Enabled = True
    
End Function
Public Property Get WindowsDirectory() As String
    Static r As String
    If Len(r) = 0 Then
        Dim L As Long
        L = MAX_PATH
        r = String(L, 0)
        L = GetWindowsDirectory(r, L)
        If L > 0 Then
            r = Left$(r, L)
            AddSlash r
        Else
            r = ""
        End If
    End If
    WindowsDirectory = r
End Property
Public Property Get WindowsTempDirectory() As String
    Static m_WindowsTempDirectory As String
    If Len(m_WindowsTempDirectory) = 0 Then
        Dim Buffer As String
        Dim Length As Long
        Buffer = String(MAX_PATH, 0)
        Length = GetTempPath(MAX_PATH, Buffer)
        If Length > 0 Then
            m_WindowsTempDirectory = Left$(Buffer, Length)
            AddSlash m_WindowsTempDirectory
        End If
    End If
    WindowsTempDirectory = m_WindowsTempDirectory
End Property
Public Property Get WindowsSystemDirectory() As String
    Static m_WindowsSystemDirectory As String
    If Len(m_WindowsSystemDirectory) = 0 Then
        Dim Buffer As String
        Dim Length As Long
        Buffer = String(MAX_PATH, 0)
        Length = GetSystemDirectory(Buffer, MAX_PATH)
        If Length > 0 Then
            m_WindowsSystemDirectory = Left$(Buffer, Length)
            AddSlash m_WindowsSystemDirectory
        End If
    End If
    WindowsSystemDirectory = m_WindowsSystemDirectory
End Property
Public Property Get AppPath() As String
    Static m_AppPath As String 'Returns Program EXE File Name
    If Len(m_AppPath) = 0 Then
        Dim ret As Long
        Dim Length As Long
        Dim FilePath As String
        Dim FileHandle As Long
        FilePath = String(MAX_PATH, 0)
        FileHandle = GetModuleHandle(App.EXEName)
        ret = GetModuleFileName(FileHandle, FilePath, MAX_PATH)
        Length = InStr(1, FilePath, vbNullChar) - 1
        If Length > 0 Then m_AppPath = Left$(FilePath, Length)
    End If
    AppPath = m_AppPath
End Property
Public Property Get DefaultSettingsFile() As String
    Static m_DefaultSettingsFile As String
    If Len(m_DefaultSettingsFile) = 0 Then m_DefaultSettingsFile = FileTitleOnly(AppPath, True) & "Settings.Dat"
    DefaultSettingsFile = m_DefaultSettingsFile
End Property
Public Property Get DefaultLegendFile() As String
    Static m_DefaultLegendFile As String
    If Len(m_DefaultLegendFile) = 0 Then m_DefaultLegendFile = FileTitleOnly(AppPath, True) & "Legends.Txt"
    DefaultLegendFile = m_DefaultLegendFile
End Property
Public Function FileExists(FileName As String) As Boolean
    If Len(FileName) > 0 Then FileExists = (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function
Public Function DirectoryExists(ByVal Directory As String) As Boolean
    AddSlash Directory
    DirectoryExists = Len(Directory) > 0 And Len(Dir(Directory + "*.*", vbDirectory)) > 0
End Function
Public Function FileTitleOnly(FileName As String, Optional ReturnDirectory As Boolean) As String
    If ReturnDirectory Then
        FileTitleOnly = Left$(FileName, InStrRev(FileName, "\"))
    Else
        FileTitleOnly = Right$(FileName, Len(FileName) - InStrRev(FileName, "\"))
    End If
End Function
Public Function AddSlash(Directory As String) As String
    If InStrRev(Directory, "\") <> Len(Directory) Then
        Directory = Directory + "\"
    End If
    AddSlash = Directory
End Function
Public Sub RemoveSlash(Directory As String)
    If Len(Directory) > 3 And InStrRev(Directory, "\") = Len(Directory) Then Directory = Left$(Directory, Len(Directory) - 1)
End Sub
Public Sub RidFile(FileName As String)
    If FileExists(FileName) Then
        SetAttr FileName, vbNormal
        Kill FileName
    End If
End Sub
Public Function GetShortName(ByVal FileName As String) As String
    Dim Buffer As String
    Dim Length As Long
    Buffer = String(MAX_PATH, 0)
    Length = GetShortPathName(FileName, Buffer, MAX_PATH)
    If Length > 0 Then GetShortName = Left$(Buffer, Length)
End Function
Public Function CreateTempFile(Optional ByVal Prefix As String, Optional Directory As String) As String
    Dim Buffer As String
    Dim Length As Long
    Buffer = String(MAX_PATH, 0)
    If Len(Prefix) = 0 Then Prefix = Left$(App.Title + "TMP", 3)
    If Not DirectoryExists(Directory) Then Directory = WindowsTempDirectory
    If GetTempFileName(Directory, Prefix, 0&, Buffer) = 0 Then Exit Function
    Length = InStr(1, Buffer, vbNullChar) - 1
    If Length > 0 Then CreateTempFile = Left$(Buffer, Length)
End Function
Public Function CreatePath(ByVal Path As String) As Boolean
    On Error GoTo Fail
    Dim I As Integer
    Dim s As String
    AddSlash Path
    Do
        I = InStr(I + 1, Path, "\")
        If I = 0 Then Exit Do
        s = Left$(Path, I - 1)
        If Not DirectoryExists(s) Then MkDir s
    Loop Until I = Len(Path)
    
    If DirectoryExists(Path) Then
        CreatePath = True
        Exit Function
    End If
Fail:
    Call MsgBox(IIf(Err.Number = 0, "", "Error " + CStr(Err.Number) + ": " + Err.Description + vbCrLf) + "Could Not Create/Access Directory:" + vbCrLf + vbCrLf + Chr$(34) + Path + Chr$(34), vbExclamation, App.Title + " - CreatePath Function")
    
End Function


Public Function GetFolderName(hWnd As Long, Text As String) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim Path As String
    With bi
        .hwndOwner = hWnd
        .pidlRoot = 0&                                                          'Root directory, generally does not need to be changed
        .lpszTitle = lstrcat(Text, "")
        .ulFlags = BIF_RETURNONLYFSDIRS                                         'Adjust as needed
    End With
    pidl = SHBrowseForFolder(bi)
    Path = Space$(512)
    If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
        GetFolderName = Left(Path, InStr(Path, Chr(0)) - 1)
    End If
End Function

'Search for all files in the specified directory, does not support subfolder recursion
'Usage example
'  SearchFiles "C:\Program Files\WinRAR\", "*" 'Search for all files
'  SearchFiles "C:\Program Files\WinRAR\", "*.exe" 'Search for all exe files
'  SearchFiles "C:\Program Files\WinRAR\", "*in*.exe" 'Search for exe files that contain 'in' in the filename
Public Function SearchFiles(Path As String, FileType As String, Optional ForceSuffixMatch As Boolean = False) As String()
    Dim sPath As String, numFiles As Long, suffix As String, nLensuffix As Long
    Dim saFiles() As String
    
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    'Force suffix match
    If ForceSuffixMatch Then
        nLensuffix = Len(FileType) - InStrRev(FileType, ".")
        suffix = Right$(FileType, nLensuffix)
    End If
    
    On Error GoTo SFErr
    sPath = Dir(Path & FileType) 'Find the first file
    
    numFiles = 0
    Do While Len(sPath) 'Loop until no files are found
        If ForceSuffixMatch Then  'Force suffix match, not just a part in the middle
            If Right$(sPath, nLensuffix) = suffix Then
                ReDim Preserve saFiles(numFiles) As String
                saFiles(numFiles) = Path & sPath
                numFiles = numFiles + 1
            End If
        Else
            ReDim Preserve saFiles(numFiles) As String
            saFiles(numFiles) = Path & sPath
            numFiles = numFiles + 1
        End If
        sPath = Dir 'Find the next file
        'DoEvents 'Yield control
    Loop
    
    If numFiles Then
        SearchFiles = saFiles
    Else
        SearchFiles = Split("")
    End If
    Exit Function
SFErr:
    SearchFiles = Split("")
End Function
