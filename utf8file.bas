Attribute VB_Name = "utf8file"
' UTF8 file read/write
Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

' UTF-8 code page constant
Private Const CP_UTF8 = 65001

' Returns the number of elements in a byte array
Private Function BytesLength(abBytes() As Byte) As Long
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function

' Converts a string to a UTF-8 byte array
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

' Converts a UTF-8 byte array to a string
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function

Public Function ReadFileIntoString(sFilePath As String) As String
    ' Reads file (if it exists) into a string.
    Dim strIn As String
    Dim hFile As Integer
    
    ' Check if file exists
    If Len(Dir(sFilePath)) = 0 Then
        Exit Function
    End If
    hFile = FreeFile
    Open sFilePath For Binary Access Read As #hFile
    strIn = Input(LOF(hFile), #hFile)
    Close #hFile
    ReadFileIntoString = strIn
    
End Function

Public Function WriteFileFromString(sFilePath As String, strIn As String) As Boolean
    ' Creates a file from a string. Clobbers any existing file.
    On Error GoTo OnError
    Dim hFile As Integer
    
    If Len(Dir(sFilePath)) > 0 Then
        Kill sFilePath
    End If
    hFile = FreeFile
    Open sFilePath For Binary Access Write As #hFile
    Put #hFile, , strIn
    Close #hFile
    WriteFileFromString = True
Done:
    Exit Function
OnError:
    Resume Done
    
End Function

Public Function ReadFileIntoBytes(sFilePath As String) As Byte()
    ' Reads file (if it exists) into an array of bytes.
    Dim abData() As Byte
    Dim hFile As Integer
    
    ' Set default return value that won't cause a run-time error
    ReadFileIntoBytes = StrConv("", vbFromUnicode)
    ' Check if file exists
    If Len(Dir(sFilePath)) = 0 Then
        Exit Function
    End If
    hFile = FreeFile
    Open sFilePath For Binary Access Read As #hFile
    abData = InputB(LOF(hFile), #hFile)
    Close #hFile
    ReadFileIntoBytes = abData
    
End Function

Public Function WriteFileFromBytes(sFilePath As String, abData() As Byte) As Boolean
    ' Creates a file from a string. Clobbers any existing file.
    On Error GoTo OnError
    Dim hFile As Integer
    
    If Len(Dir(sFilePath)) > 0 Then
        Kill sFilePath
    End If
    hFile = FreeFile
    Open sFilePath For Binary Access Write As #hFile
    Put #hFile, , abData
    Close #hFile
    WriteFileFromBytes = True
Done:
    Exit Function
OnError:
    Resume Done
    
End Function

' External interfaces
' Reads the binary data of a file into a byte array, returns the number of bytes read, 0 indicates failure
Public Function ReadFileBinaryContent(sFile As String, ByRef abContent() As Byte) As Long
    
    Dim fn As Long, nSize As Long
    
    On Error GoTo FileError
    
    ' Get binary data
    fn = FreeFile
    Open sFile For Binary As fn
    nSize = LOF(fn)
    ReDim abContent(nSize - 1) As Byte
    Get fn, , abContent
    Close fn
    ReadFileBinaryContent = nSize
    Exit Function
    
FileError:
    Close fn
    ReadFileBinaryContent = 0
    
End Function

' Writes UTF-8 file
Public Sub Utf8File_Write_VB(ByVal sFileName As String, ByVal vVar As String)
    Dim b() As Byte
    
    b = Utf8BytesFromString(vVar)
    WriteFileFromBytes sFileName, b
End Sub

' Below is the previous implementation, requires external dependencies
' Need to add reference to Microsoft Activex data objects 2.8 library
'Public Sub Utf8File_Write_VB(ByVal sFileName As String, ByVal vVar As String)
'    Dim adostream As New ADODB.Stream
'    Dim fn As Long, abContent() As Byte, nSize As Long
'    With adostream
'        .Type = adTypeText
'        .Mode = adModeReadWrite
'        .Charset = "utf-8"
'        .Open
'        .Position = 0
'        .WriteText vVar
'        .SaveToFile sFileName, adSaveCreateOverWrite
'        .Close
'    End With
'    Set adostream = Nothing
'
'    ' Remove BOM
'    On Error GoTo FileError
'
'    fn = FreeFile
'    Open sFileName For Binary As fn
'    nSize = LOF(fn)
'    ReDim abContent(nSize - 3) As Byte
'    Get fn, 4, abContent
'    Close fn
'    Open sFileName For Binary As fn
'    Put fn, , abContent
'    Close fn
'    Exit Sub
'
'FileError:
'    Close fn
'End Sub

' Need to add reference to Microsoft Activex data objects 2.8 library
'Public Function Utf8File_Read_VB(ByVal sFileName As String) As String
'    Dim adostream As New ADODB.Stream
'    With adostream
'        .Type = adTypeText
'        .Mode = adModeReadWrite
'        .Charset = "utf-8"
'        .Open
'        .LoadFromFile sFileName
'        Utf8File_Read_VB = .ReadText
'        .Close
'    End With
'    Set adostream = Nothing
'End Function

