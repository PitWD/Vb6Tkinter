VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Menu class, this class is somewhat different from other control classes because it needs to handle nested menus and other special situations

Private m_dicTotal As Dictionary        'Store all properties, including default values
Private m_Base As clsBaseControl        'Base control class

Private m_Childs() As Object
Private m_numChilds As Long
Private m_IdxCurChild As Long
Private m_CanbeOutByMainForm As Boolean

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handler callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    Dim s() As String, I As Long, extra As String
    
    extra = IIf(Len(m_Base("tearoff")), ", tearoff=" & m_Base("tearoff"), "")
    
    sOut.Append "        self." & m_Base.Name & " = Menu(self." & m_Base.Parent & extra & ")"
    
    'Output each top-level menu one by one, the submenus under each top-level menu are responsible for generating the corresponding code
    For I = 0 To m_numChilds - 1
        m_Childs(I).toString sOut, sCmd, sI18n, rel, usettk, I
    Next
    
    sOut.Append "        self." & m_Base.Parent & ".configure(menu=self." & m_Base.Name & ")"
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(Optional o As Object, Optional parentWidth As Long, Optional parentHeight As Long, Optional dMethods As Dictionary)
    
    m_Base.SetVbWidgetInstance o
    
    m_Base("tearoff") = "0"
    
    'These are all the default values
    m_dicTotal("tearoff") = "0"
    m_dicTotal("label") = ""
    m_dicTotal("fg") = ""
    m_dicTotal("bg") = ""
    m_dicTotal("bd") = ""
    m_dicTotal("relief") = ""
    m_dicTotal("state") = ""
    m_dicTotal("underline") = "-1"
    m_dicTotal("variable") = m_Base.Name & "Var"
    m_dicTotal("font") = ""
    
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If sAttr = "tearoff" Then
        GetAttrValueList = 1
        sa = Split("1,0", ",")
    Else
        GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End If
End Function

'Check if this control has the specified attribute
Public Function hasAttribute(sAttr As String) As Boolean
    hasAttribute = m_Base.hasAttribute(sAttr)
End Function

'Get the current value of the specified attribute for this control, return an empty string if not set
Public Function GetAttrCurrentValue(sAttr As String) As String
    GetAttrCurrentValue = m_Base.GetAttrCurrentValue(sAttr)
End Function

Public Function Tips(sAttr As String) As String
    
    If sAttr = "tearoff" Then
        Tips = sAttr & vbCrLf & L("l_TipTearOff", "Determines menu can be torn off or not.")
    ElseIf sAttr = "postcommand" Then
        Tips = sAttr & vbCrLf & L("l_TipPostCmdMenu", "A procedure will be called every time someone brings up this menu.")
    Else
        Tips = m_Base.Tips(sAttr)
    End If
    
End Function

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.Name = "MainMenu"
    m_Base.ctlType = "Menu"
    m_Base.StyleName = ""
    Erase m_Childs
    m_numChilds = 0
    m_IdxCurChild = 0
    m_CanbeOutByMainForm = True
End Sub

'Return a collection, each item is a triplet "property name|value|default selection"
'This function is used to populate the property parameter list box on the main interface
Public Function Allitems() As Collection
    Dim re As Collection, k As Variant, ks As Collection
    
    Set re = New Collection
    
    'Standard parameters
    Set ks = m_dicTotal.Keys
    For Each k In ks
        If Len(m_Base(k)) Then
            re.Add k & "|" & m_Base(k) & "|1"
        Else
            re.Add k & "|" & m_dicTotal(k) & "|0"
        End If
    Next
    
    'User-added custom parameters (if any)
    Set ks = m_Base.Keys
    For Each k In ks
        If Not m_dicTotal.Exists(k) Then
            re.Add k & "|" & m_Base(k) & "|1"
        End If
    Next
    
    Set Allitems = re
    
End Function

Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
End Sub

Private Sub Class_Terminate()
    Set m_dicTotal = Nothing
    Set m_Base = Nothing
    Erase m_Childs
End Sub

Public Property Let Parent(s As String)
    m_Base.Parent = s
End Property

Public Property Get Parent() As String
    Parent = m_Base.Parent
End Property

Public Property Get Name() As String
    Name = m_Base.Name
End Property

Public Property Let Name(s As String)
    m_Base.Name = s
End Property

'Used to change its default corresponding widget type, note the legality of the attribute list after modifying the widget type
Public Function SetWidgetType(sType As String, sStyleName As String)
    'm_Base.ctlType = sType
    'm_Base.StyleName = sStyleName
End Function

'Determine whether the main processing function can call its toString() to generate code, default is True, setting to False indicates that it is handled by other objects
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = m_CanbeOutByMainForm
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    m_CanbeOutByMainForm = bEnable
End Property

'Object serialization function
Public Sub Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Sub

Public Sub Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Sub

Public Property Get Description() As String
    Description = L("l_DescMenu", "Menu widget.")
End Property

Public Sub AddChild(o As Object)
    ReDim Preserve m_Childs(m_numChilds) As Object
    Set m_Childs(m_numChilds) = o
    m_numChilds = m_numChilds + 1
End Sub

Public Function GetNextChild(Optional nIdxChild As Long = -1) As Object
    
    m_IdxCurChild = IIf(nIdxChild >= 0, nIdxChild, m_IdxCurChild)
    If m_IdxCurChild < m_numChilds Then
        Set GetNextChild = m_Childs(m_IdxCurChild)
        m_IdxCurChild = m_IdxCurChild + 1
    Else
        Set GetNextChild = Nothing
        m_IdxCurChild = 0
    End If
End Function

Public Property Get ChildCount() As Long
    ChildCount = m_numChilds
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property

'Function used to simulate compare sorting, actually to determine the dependency relationship between two objects
'Compare this object with another object, logical result is 'this object - another object'
'Return value meaning:
'<0: indicates that this object needs to output code before another object
'=0: indicates that there is no dependency relationship between the two, code order is not affected
'>0: another object needs to output code first.
'The overall logical result is similar to heavier items sinking to the bottom
Public Function Compare(ByRef Obj As Object) As Long
    If Parent = Obj.Name Then 'Parent control outputs code first
        Compare = 1
    ElseIf Obj.Parent = Name Then
        Compare = -1
    ElseIf Parent = WTOP And Obj.Parent <> WTOP Then 'Top-level controls output first
        Compare = -1
    ElseIf Parent <> WTOP And Obj.Parent = WTOP Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function


