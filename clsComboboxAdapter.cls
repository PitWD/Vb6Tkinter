VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboboxAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Combobox adapter class
'Since there is no combobox in TK, the VB combobox is mapped to Tk's OptionMenu
'If TTK is enabled, it is mapped to TTK's Combobox
'Once there is a combobox on the VB form, both OptionMenu and Combobox are created
'So an intermediate layer is used as an adapter to select one of them

Private m_usettk As Boolean
Private m_OptionMenu As clsOptionMenu
Private m_Combobox As clsCombobox
Private m_CanbeOutByMainForm As Boolean

Public Property Let TTK(usettk As Boolean)
    m_usettk = usettk
End Property

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handler callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    If usettk Then
        m_Combobox.toString sOut, sCmd, sI18n, rel, usettk
    Else
        m_OptionMenu.toString sOut, sCmd, sI18n, rel, usettk
    End If
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    m_Combobox.InitConfig o, parentWidth, parentHeight, dMethods
    m_OptionMenu.InitConfig o, parentWidth, parentHeight, dMethods
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If m_usettk Then
        GetAttrValueList = m_Combobox.GetAttrValueList(sAttr, sa)
    Else
        GetAttrValueList = m_OptionMenu.GetAttrValueList(sAttr, sa)
    End If
End Function

'Check if this control has the specified attribute
Public Function hasAttribute(sAttr As String) As Boolean
    If m_usettk Then
        hasAttribute = m_Combobox.hasAttribute(sAttr)
    Else
        hasAttribute = m_OptionMenu.hasAttribute(sAttr)
    End If
End Function

'Get the current value of the specified attribute for this control, return an empty string if not set
Public Function GetAttrCurrentValue(sAttr As String) As String
    If m_usettk Then
        GetAttrCurrentValue = m_Combobox.GetAttrCurrentValue(sAttr)
    Else
        GetAttrCurrentValue = m_OptionMenu.GetAttrCurrentValue(sAttr)
    End If
End Function

Public Function Tips(sAttr As String) As String
    If m_usettk Then
        Tips = m_Combobox.Tips(sAttr)
    Else
        Tips = m_OptionMenu.Tips(sAttr)
    End If
End Function

Private Sub Class_Initialize()
    m_usettk = True
    Set m_OptionMenu = New clsOptionMenu
    Set m_Combobox = New clsCombobox
    m_CanbeOutByMainForm = True
End Sub

'Return a collection, each item is a triplet "property name|value|default selection"
'This function is used to populate the property parameter list box on the main interface
Public Function Allitems() As Collection
    If m_usettk Then
        Set Allitems = m_Combobox.Allitems()
    Else
        Set Allitems = m_OptionMenu.Allitems()
    End If
End Function

'Update the configuration selected by the user to the object, the parameter is many attribute/value pairs separated by "|"
Public Sub SetConfig(sAttrs As String)
    If m_usettk Then
        m_Combobox.SetConfig (sAttrs)
    Else
        m_OptionMenu.SetConfig (sAttrs)
    End If
End Sub

'Modify or add a single configuration item, attribute/value separated by "|"
Public Sub SetSingleConfig(sAttr As String)
    If m_usettk Then
        m_Combobox.SetSingleConfig (sAttr)
    Else
        m_OptionMenu.SetSingleConfig (sAttr)
    End If
End Sub

Public Property Let Parent(s As String)
    m_Combobox.Parent = s
    m_OptionMenu.Parent = s
End Property

Public Property Get Parent() As String
    If m_usettk Then
        Parent = m_Combobox.Parent
    Else
        Parent = m_OptionMenu.Parent
    End If
End Property

Public Property Get Name() As String
    If m_usettk Then
        Name = m_Combobox.Name
    Else
        Name = m_OptionMenu.Name
    End If
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
Public Function Serializer(vSer As clsSerialization)
    If m_usettk Then
        m_Combobox.Serializer vSer
    Else
        m_OptionMenu.Serializer vSer
    End If
End Function

Public Function Deserializer(vSer As clsSerialization)
    If m_usettk Then
        m_Combobox.Deserializer vSer
    Else
        m_OptionMenu.Deserializer vSer
    End If
End Function

Public Property Get Description() As String
    If m_usettk Then
        Description = m_Combobox.Description
    Else
        Description = m_OptionMenu.Description
    End If
End Property


Public Property Let ScaleMode(nV As Long)
    m_Combobox.ScaleMode = nV
    m_OptionMenu.ScaleMode = nV
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


