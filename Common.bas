Attribute VB_Name = "Common"
Option Explicit

Public VbeInst As VBE

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long

Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TWIPSPERINCH = 1440
Private Type Size
    cx As Long
    cy As Long
End Type

' Declare API functions
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, ByVal lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpftLastWriteTime As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Public Const REG_SZ = 1
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const SYNCHRONIZE = &H100000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WOW64_64KEY = &H100

' These are used to get the system default font
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const DEFAULT_GUI_FONT = 17
Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Const WTOP = "top" ' Used for positioning in tkinter

Public g_DefaultFontName As String ' Cache the system default font name to avoid querying every time
Public g_Comps() As Object ' Current component list, the first item is the form instance

Public g_bUnicodePrefixU As Boolean ' Whether to add the prefix 'u' to Unicode strings
Public g_PythonExe As String ' Full path of python.exe for GUI precompilation
Public g_AppVerString As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const OFFICIAL_SITE As String = "https://github.com/cdhigh/Vb6Tkinter"
Public Const OFFICIAL_RELEASES As String = "https://github.com/cdhigh/Vb6Tkinter/releases"
Public Const OFFICIAL_UPDATE_INFO As String = "https://api.github.com/repos/cdhigh/Vb6Tkinter/releases"

' This function handles Unicode string prefixes.
' It checks if the string contains double-byte characters and adds the prefix if needed.
Public Function U(s As String) As String
    
    Dim nLen As Long
    s = Replace(s, vbCrLf, "\n")
    nLen = Len(s)
    
    If lstrlen(s) > nLen Then  ' Contains double-byte characters
        If g_bUnicodePrefixU Then
            U = IIf(isQuoted(s), "u" & s, "u'" & s & "'")
        Else
            U = IIf(isQuoted(s), s, "'" & s & "'")
        End If
    ElseIf nLen Then
        U = IIf(isQuoted(s), s, "'" & s & "'")
    Else
        U = "''"
    End If
    
End Function

' Determines if a string is already quoted with single or double quotes
Public Function isQuoted(s As String) As Boolean
    isQuoted = (Left$(s, 1) = "'" Or Left$(s, 1) = Chr$(34)) And (Right$(s, 1) = "'" Or Right$(s, 1) = Chr$(34))
End Function

' Removes quotes from a string if they exist
Public Function UnQuote(s As String) As String
    If isQuoted(s) Then
        UnQuote = Mid(s, 2, Len(s) - 2)
    Else
        UnQuote = s
    End If
End Function

' Adds quotes to a string with either single or double quotes
Public Function Quote(s As String) As String
    If isQuoted(s) Then
        Quote = s
    ElseIf InStr(1, s, "'") >= 1 Then ' If the string contains single quotes, use double quotes
        Quote = Chr$(34) & s & Chr$(34)
    Else
        Quote = "'" & s & "'"
    End If
End Function

' Quickly removes the first and last character of a string, assuming they are quotes
Public Function UnQuoteFast(s As String) As String
    UnQuoteFast = Mid(s, 2, Len(s) - 2)
End Function

' Quickly adds single quotes to a string
Public Function QuoteFast(s As String) As String
    QuoteFast = "'" & s & "'"
End Function

' Gets the full name of a file including its extension
Public Function FileFullName(ByVal sF As String) As String
    Dim ns As Long
    
    ns = InStrRev(sF, "\")
    If ns <= 0 Then
        FileFullName = sF
    Else
        FileFullName = Right$(sF, Len(sF) - ns)
    End If
End Function

' Gets the extension of a file
Public Function FileExt(sF As String) As String
    Dim sFName As String, ns As Long
    sFName = FileFullName(sF)
    ns = InStrRev(sFName, ".")
    If ns > 0 Then
        FileExt = Right$(sFName, Len(sFName) - ns)
    End If
End Function

' Gets the path name of a file, excluding the file name
Public Function PathName(sF As String) As String
    Dim ns As Long
    
    ns = InStrRev(sF, "\")
    If ns <= 0 Then
        PathName = ""
    Else
        PathName = Left$(sF, ns)
    End If
    
End Function

Function getDPI(bX As Boolean) As Integer                                       ' Get screen DPI
    Dim hdc As Long, RetVal As Long
    hdc = GetDC(0)
    If bX = True Then
        getDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Else
        getDPI = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    RetVal = ReleaseDC(0, hdc)
End Function
Function Twip2PixelX(x As Long) As Long                                         ' Convert horizontal Twips to Pixels
    Twip2PixelX = x / TWIPSPERINCH * getDPI(True)
End Function
Function Twip2PixelY(x As Long) As Long                                         ' Convert vertical Twips to Pixels
    Twip2PixelY = x / TWIPSPERINCH * getDPI(False)
End Function
Function Point2PixelX(x As Long) As Long                                        ' Convert horizontal Points to Pixels
    Point2PixelX = Twip2PixelX(x * 20)
End Function
Function Point2PixelY(x As Long) As Long                                        ' Convert vertical Points to Pixels
    Point2PixelY = Twip2PixelY(x * 20)
End Function
Function getScreenX() As Long                                                   ' Get screen width
    Dim hdc As Long, RetVal As Long
    hdc = GetDC(0)
    getScreenX = GetDeviceCaps(hdc, HORZRES)
    RetVal = ReleaseDC(0, hdc)
End Function
Function getScreenY() As Long                                                   ' Get screen height
    Dim hdc As Long, RetVal As Long
    hdc = GetDC(0)
    getScreenY = GetDeviceCaps(hdc, VERTRES)
    RetVal = ReleaseDC(0, hdc)
End Function

Public Function CharWidth() As Long                ' Get the width of the default font character (average)
    Dim hdc As Long, RetVal As Long
    Dim typSize     As Size
    Dim lngX     As Long
    Dim lngY     As Long
    
    hdc = GetDC(0)
    RetVal = GetTextExtentPoint32(hdc, "ABli", 4, typSize)
    CharWidth = typSize.cx / 4
    RetVal = ReleaseDC(0, hdc)
End Function

' Translate VB color to Python RGB color
' If the color is a system color, it cannot be translated
Public Function TranslateColor(ByVal dwColor As OLE_COLOR) As String
    Dim nColor As Long, hPalette As Long, clrHex As String
    If OleTranslateColor(dwColor, hPalette, nColor) = 0 Then
        clrHex = Replace(Format(Hex$(nColor), "@@@@@@"), " ", "0")
        TranslateColor = "'#" & Mid$(clrHex, 5, 2) & Mid$(clrHex, 3, 2) & Mid$(clrHex, 1, 2) & "'"
    End If
End Function

' Get all installed Python paths from the system registry
Public Function GetAllInstalledPython() As String()
MsgBox "Wait... GetAllInstalledPython"
    Dim nRe As Long, nHk As Long, nHk2 As Long, I As Long, nLen As Long
    Dim sVer As String, sAllPath As String, sBuff As String, sPythonExe As String
    Dim saVer() As String, nVerNum As Long
    
    nRe = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Python\PythonCore", 0, KEY_READ Or KEY_WOW64_64KEY, nHk)
    If nRe <> 0 Then
        GetAllInstalledPython = Split("")
        Exit Function
    End If
    
    I = 0
    nVerNum = 0
    nLen = 255
    sBuff = String$(255, 0)
    Do While (RegEnumKeyEx(nHk, I, sBuff, nLen, 0, vbNullString, ByVal 0&, ByVal 0&) = 0)
        If nLen > 1 Then
            sBuff = Left$(sBuff, InStr(1, sBuff, Chr(0)) - 1)
            
            ReDim Preserve saVer(nVerNum) As String
            saVer(nVerNum) = sBuff
            nVerNum = nVerNum + 1
        End If
        I = I + 1
        nLen = 255
        sBuff = String$(255, 0)
    Loop
    RegCloseKey nHk
    
    ' Query installation paths
    For I = 1 To nVerNum
        nRe = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Python\PythonCore\" & saVer(I - 1) & "\InstallPath", 0, KEY_READ Or KEY_WOW64_64KEY, nHk2)
        If nRe = 0 Then
            nLen = 255
            sBuff = String$(255, 0)
            nRe = RegQueryValueEx(nHk2, "", 0&, REG_SZ, sBuff, nLen)  ' Query default string value of the subkey
            If nRe = 0 And nLen > 1 Then
                sBuff = Left$(sBuff, InStr(1, sBuff, Chr(0)) - 1)
                
                sPythonExe = sBuff & IIf(Right$(sBuff, 1) = "\", "", "\") & "python.exe"
                sPythonExe = sPythonExe & "," & sBuff & IIf(Right$(sBuff, 1) = "\", "", "\") & "pythonw.exe"
                sAllPath = sAllPath & IIf(Len(sAllPath), ",", "") & sPythonExe
            End If
            RegCloseKey nHk2
        End If
    Next
    
    GetAllInstalledPython = Split(sAllPath, ",")
End Function

' Get the system default font name
Public Function GetDefaultFontName() As String
    Dim hFont As Long, lfont As LOGFONT
    
    If Len(g_DefaultFontName) Then
        GetDefaultFontName = g_DefaultFontName
    Else
        hFont = GetStockObject(DEFAULT_GUI_FONT)
        If hFont <> 0 Then
            GetObject hFont, Len(lfont), lfont
            DeleteObject hFont
            GetDefaultFontName = StrConv(lfont.lfFaceName, vbUnicode)
            If InStr(1, GetDefaultFontName, Chr(0)) > 0 Then
                GetDefaultFontName = Left$(GetDefaultFontName, InStr(1, GetDefaultFontName, Chr(0)) - 1)
            End If
            g_DefaultFontName = GetDefaultFontName  ' Cache it to avoid querying the API next time
        End If
    End If
End Function

' Get the list of controls in the current form, returning a string with "|" as the separator
Public Function GetAllComps() As String()
    Dim nCnt As Long, I As Long, sa() As String
    On Error Resume Next
    nCnt = UBound(g_Comps)
    On Error GoTo 0
    If nCnt <= 0 Then
        GetAllComps = Split("")
        Exit Function
    End If
    
    ReDim sa(nCnt) As String
    For I = 0 To nCnt
        sa(I) = g_Comps(I).Name & "|" & TypeName(g_Comps(I))
    Next
    GetAllComps = sa
End Function

' Sort controls in the form, simple bubble sort, needs to be done before adding controls
' The reason is to ensure that controls are added in order, so that larger controls are added first, followed by small controls
Public Sub SortWidgets(ByRef aCompsSorted() As Object, ByVal cnt As Long)
    Dim idx1 As Long, idx2 As Long
    Dim tmp4exchange As Object
    
    For idx1 = 0 To cnt - 2
        For idx2 = idx1 + 1 To cnt - 1
            If aCompsSorted(idx1).Compare(aCompsSorted(idx2)) > 0 Then ' Larger is in front
                Set tmp4exchange = aCompsSorted(idx1)
                Set aCompsSorted(idx1) = aCompsSorted(idx2)
                Set aCompsSorted(idx2) = tmp4exchange
            End If
        Next
    Next
    
End Sub

' Purify version string by removing non-numeric prefixes and suffixes, e.g., "v1.6.8 test" becomes "1.6.8"
Private Function purifyVerStr(txt As String) As String
    Dim maxCnt As Integer, idx As Integer, startIdx As Integer, endIdx As Integer
    Dim ch As String
    txt = Trim(txt)
    maxCnt = Len(txt)
    startIdx = 1
    endIdx = maxCnt
    ' Trim leading characters
    For idx = 1 To maxCnt
        ch = Mid(txt, idx, 1)
        If (ch >= "0") And (ch <= "9") Then
            startIdx = idx
            Exit For
        End If
    Next
    ' Trim trailing characters
    For idx = maxCnt To 1 Step -1
        ch = Mid(txt, idx, 1)
        If (ch >= "0") And (ch <= "9") Then
            endIdx = idx
            Exit For
        End If
    Next
    
    If startIdx <= endIdx Then
        purifyVerStr = Mid(txt, startIdx, endIdx - startIdx + 1)
    Else
        purifyVerStr = ""
    End If
End Function

' Compare two version numbers to determine if the new version is higher than the current version
' Version format is like 1.1.0
Public Function isVersionNewerThan(newVer As String, currVer As String) As Boolean
    Dim newArr As Variant, currArr As Variant, idx As Integer, maxCnt As Integer
    Dim vn As Integer, vc As Integer
    newVer = purifyVerStr(newVer)
    currVer = purifyVerStr(currVer)
    If Len(newVer) = 0 Or Len(currVer) = 0 Then
        isVersionNewerThan = False
        Exit Function
    End If
    
    newArr = Split(newVer, ".")
    currArr = Split(currVer, ".")
    maxCnt = UBound(newArr)
    If UBound(currArr) < maxCnt Then 'The smallest of the two arrays
        maxCnt = UBound(currArr)
    End If
    
    For idx = 0 To maxCnt
        vn = Val(newArr(idx))
        vc = Val(currArr(idx))
        If vn > vc Then
            isVersionNewerThan = True
            Exit Function
        ElseIf vn < vc Then
            isVersionNewerThan = False
            Exit Function
        End If
    Next
    
    ' If all previous values are the same, the longer one is considered larger
    If UBound(newArr) > UBound(currArr) Then
        isVersionNewerThan = True
    Else
        isVersionNewerThan = False
    End If
End Function
