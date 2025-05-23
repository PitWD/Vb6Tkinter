VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Usage
'dim ss as cStrBuilder
'set ss = new cStrBuilder
's.Append "test"
'msgbox s.toString()

' The secret to this class is that it uses the join
'function which is part of the VBA.Strings Class
Private mvarStringArray() As String
Private mvarArrayItems As Long

Public Sub Append(ByVal newStr As String)
    ReDim Preserve mvarStringArray(mvarArrayItems) As String
    mvarStringArray(mvarArrayItems) = newStr
    mvarArrayItems = mvarArrayItems + 1
End Sub

Public Property Get toString(Optional sDelimiter As String = "") As String
    If mvarArrayItems > 0 Then
        toString = Join(mvarStringArray, sDelimiter)
    Else
        toString = ""
    End If
End Property

Public Sub Reset()
    mvarArrayItems = 0
    Erase mvarStringArray
End Sub

Private Sub Class_Initialize()
    If mvarArrayItems > 0 Then Reset
End Sub

Private Sub Class_Terminate()
    Reset
End Sub

'Check if a string already exists. This checks the entire string that is appended each time, not sub-strings.
Public Function ExistString(s As String) As Boolean
    Dim i As Long, nLen As Long
    nLen = Len(s)
    For i = 0 To mvarArrayItems - 1
        If nLen = Len(mvarStringArray(i)) Then
            If s = mvarStringArray(i) Then
                ExistString = True
                Exit Function
            End If
        End If
    Next
End Function
