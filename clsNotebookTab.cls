VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotebookTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Single page in the tab, used to set and maintain the corresponding properties and values, and generate the corresponding PYTHON TKinter code

Private m_Comps() As Object  'Save all controls in this tab page
Private m_CanbeOutByMainForm As Boolean
Private m_SubWidgets() As Object
Private m_numSubWidgets As Long
Private m_Name As String
Private m_Caption As String
Private m_Parent  As String

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handling callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    Dim I As Long
    
    'Sort by dependency relationship first
    SortWidgets m_SubWidgets, m_numSubWidgets
    
    For I = 0 To m_numSubWidgets - 1
        m_SubWidgets(I).toString sOut, sCmd, sI18n, rel, usettk
    Next
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, Optional dMethods As Dictionary)
    m_Name = o.Properties("Name")
    
    On Error Resume Next
    m_Caption = o.Properties("Caption")
    On Error GoTo 0
    
    Erase m_SubWidgets
    m_numSubWidgets = 0
End Sub

'Add a sub control
Public Sub AddSubWidget(oWidget As Object)
    ReDim Preserve m_SubWidgets(m_numSubWidgets) As Object
    Set m_SubWidgets(m_numSubWidgets) = oWidget
    m_numSubWidgets = m_numSubWidgets + 1
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 0
End Function

'Check if this control has the specified attribute
Public Function hasAttribute(sAttr As String) As Boolean
    hasAttribute = False
End Function

'Get the current value of the specified attribute for this control, return an empty string if not set
Public Function GetAttrCurrentValue(sAttr As String) As String
    GetAttrCurrentValue = ""
End Function

Public Function Tips(sAttr As String) As String
    Tips = ""
End Function

'Update the configuration selected by the user to the object, the parameter is many attribute/value pairs separated by "|"
Public Sub SetConfig(sAttrs As String)
    '''''
End Sub

'Modify or add a single configuration item, attribute/value separated by "|"
Public Sub SetSingleConfig(sAttr As String)
    ''''''
End Sub

'Return a collection, each item is a triplet "property name|value|default selection"
'This function is used to populate the property parameter list box on the main interface
Public Function Allitems() As Collection
    Set Allitems = New Collection
End Function

Private Sub Class_Initialize()
    m_numSubWidgets = 0
    m_CanbeOutByMainForm = True
    m_Parent = WTOP
End Sub

Public Property Let Parent(s As String)
    m_Parent = s
End Property

Public Property Get Parent() As String
    Parent = m_Parent
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

'Used to change its default corresponding widget type, note the legality of the attribute list after modifying the widget type
Public Function SetWidgetType(sType As String, sStyleName As String)
    ''''
End Function

'Determine whether the main processing function can call its toString() to generate code, default is True, setting to False indicates that it is handled by other objects
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = m_CanbeOutByMainForm
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    m_CanbeOutByMainForm = bEnable
End Property

'Object serialization function
Public Function Serializer(vSer As clsSerialization)
    'vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    'vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescNotebookTab", "Tab of widget Notebook. TTK needed.")
End Property

Public Property Let ScaleMode(nV As Long)
    'm_Base.ScaleMode = nV
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