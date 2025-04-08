Attribute VB_Name = "Resizer"
Option Explicit
'Form controls resize with the form
Private FormOldWidth As Long
Private FormOldHeight As Long

'Call this function before calling ResizeForm
Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
        Obj.Tag = Obj.Left & "|" & Obj.Top & "|" & Obj.Width & "|" & Obj.Height
    Next Obj
    
End Sub

'Change the size of all elements in the form proportionally, call ResizeInit function before calling ResizeForm
Public Sub ResizeForm(FormName As Form)
    Dim pos() As String
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    
    If (FormOldWidth = 0) Or (FormOldHeight = 0) Then
        ResizeInit FormName
    End If
    
    ScaleX = FormName.ScaleWidth / FormOldWidth
    ScaleY = FormName.ScaleHeight / FormOldHeight
    
    On Error Resume Next
    For Each Obj In FormName
        ReDim pos(0) As String
        pos = Split(Obj.Tag, "|")
        If UBound(pos) >= 3 Then
            If TypeName(Obj) = "ComboBox" Then 'ComboBox height cannot change
                Obj.Move CSng(pos(0)) * ScaleX, CSng(pos(1)) * ScaleY, CSng(pos(2)) * ScaleX
            Else
                Obj.Move CSng(pos(0)) * ScaleX, CSng(pos(1)) * ScaleY, CSng(pos(2)) * ScaleX, CSng(pos(3)) * ScaleY
            End If
        End If
    Next
    
End Sub

'Get the design-time width of the control
Public Function GetOrignalWidth(ctl As Control) As Single
    
    Dim pos() As String, i As Long
    
    On Error Resume Next
    pos = Split(ctl.Tag, "|")
    If UBound(pos) >= 3 Then
        GetOrignalWidth = CSng(pos(2))
    Else
        GetOrignalWidth = 0
    End If
    
End Function

'Get the design-time height of the control
Public Function GetOrignalHeight(ctl As Control) As Single
    
    Dim pos() As String, i As Long
    
    On Error Resume Next
    pos = Split(ctl.Tag, "|")
    If UBound(pos) >= 3 Then
        GetOrignalHeight = CSng(pos(3))
    Else
        GetOrignalHeight = 0
    End If
    
End Function


