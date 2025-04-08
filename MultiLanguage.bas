Attribute VB_Name = "MultiLanguage"
'Multi-language support module
'Language file: Vb6Tkinter.lng
' File format:
'    [Language Name]
'    ControlName=String
'    OtherStringName=String    'This is used for internal strings, such as help information, use \n in the string to represent a newline
'
'ChangeLanguage(LanguageName)   : Switch the display language of the controls, this function will also cache all strings of the corresponding language to memory at once
'L(Name, DefaultString)         : Get the specified string
'L_F(Name, DefaultString, OtherParameters) : Similar to Python's {0}{1} formatted string
'GetAllLanguageName()           : All language names in Vb6Tkinter.lng, string array

Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Const LanguageFile = "Vb6Tkinter.lng"
Private m_Lng As New Dictionary                                                 'Dictionary of {Name, String} for the corresponding language
Public Const DEF_LNG = "English(&E)"

'Get a string based on the current language setting
Public Function L(sKey As String, ByVal sDefault As String) As String
    sDefault = Replace(sDefault, "\n", vbCrLf)
    L = GetString("", sKey, sDefault)
End Function

'Get a string based on the current language setting
'Supports Python-like {0}{1} formatted strings, starting from {0}, does not support {} format (without numeric index)
Public Function L_F(sKey As String, ByVal sDefault As String, ParamArray v() As Variant) As String
    
    Dim s As String, I As Long
    
    s = L(sKey, sDefault)
    
    For I = 0 To UBound(v)
        s = Replace(s, "{" & I & "}", CStr(v(I)))
    Next
    
    L_F = s
    
End Function

Public Function GetAllLanguageName() As String()
    On Error Resume Next
    Dim s As String, ns As Long
    s = vbNullString
    If LngFileExist() Then
        s = Space(1000)
        ns = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, s, 1000, LngFile())
        GetAllLanguageName = Split(Trim(Replace(Left(s, ns), Chr(0) & Chr(0), "")), Chr(0))
    Else
        GetAllLanguageName = Split(DEF_LNG)                                     'If there is no language file, default to English
    End If
    s = ""
End Function

Public Function ChangeLanguage(Language As String) As Boolean
    Dim I As Long, Ctrl As Control, s As String, ns As Long, sa() As String
    
    'First cache the language strings of the corresponding language
    s = Space(10000)
    ns = GetPrivateProfileString(Language, vbNullString, vbNullString, s, 10000, LngFile())
    sa = Split(Trim(Replace(Left(s, ns), Chr(0) & Chr(0), "")), Chr(0))
    m_Lng.RemoveAll
    For I = 0 To UBound(sa)
        s = Space(256)
        ns = GetPrivateProfileString(Language, sa(I), "", s, 256, LngFile())
        s = Trim(Replace(Replace(Left(s, ns), Chr(0), ""), "\n", vbCrLf))
        If Len(s) Then m_Lng.Add sa(I), s
    Next
    
    'Switch the language of all controls
    For I = 0 To Forms.Count - 1
        For Each Ctrl In Forms(I).Controls
            ChangeControlLanguage Ctrl, Language
        Next
    Next I
    
    ChangeLanguage = ns > 0
    
End Function

Public Sub ChangeControlLanguage(ctl As Control, Language As String)
    Select Case TypeName(ctl)
    Case "Label", "CommandButton", "CheckBox", "OptionButton", "Frame", "Menu", "xpcmdbutton"
        ctl.Caption = GetString(Language, ctl.Name, ctl.Caption)
    End Select
End Sub

Private Function LngFile() As String
    LngFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & LanguageFile
End Function

Public Function LngFileExist() As Boolean
    LngFileExist = IIf(Dir(LngFile()) = LanguageFile, True, False)
End Function

Private Function GetString(Language As String, Key As String, sDefault As String) As String
    If m_Lng.Exists(Key) Then
        GetString = m_Lng.Item(Key)
    Else
        GetString = sDefault
    End If
End Function

