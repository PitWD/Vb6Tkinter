VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCombobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Combobox class, used to set and maintain the corresponding properties and values, and generate the corresponding PYTHON TKinter code

Private m_dicTotal As Dictionary        'Store all properties, including default values
Private m_Base As clsBaseControl        'Base control class
Private m_List As String                'Predefined list items
Private m_CanbeOutByMainForm As Boolean

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handler callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    'First create the dropdown list
    If Len(m_Base("values")) Then
        sOut.Append "        self." & m_Base.Name & "List = " & m_List
        If Len(m_Base("textvariable")) Then m_Base("text") = GetFirstItemText()
    End If
    
    m_Base.toString sOut, sCmd, sI18n, rel, usettk
    
    'By default, the combobox displays the first item after startup
    If Len(m_Base("values")) Then
        If Len(m_Base("textvariable")) Then
            m_Base.Remove ("text")
        Else
            sOut.Append "        self." & m_Base.Name & ".set(self." & m_Base.Name & "List[0])"
        End If
    End If
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    
    Dim I As Long, oList As Object, sList As String
    Dim s As String, sa() As String, idx As Long
    
    m_Base.SetVbWidgetInstance o
    
    m_Base.Name = o.Properties("Name")
    
    'Save dropdown option values
    Set oList = o.Properties("List")
    sList = "["
    I = IIf(o.Properties("Style") = 2, 1, 0)  ' When Style=2 (Dropdown List), the first IndexedValue is the control name, so start from the second item
    Do While True
        If oList.IndexedValue(I) = "" Then Exit Do
        sList = sList & U(oList.IndexedValue(I)) & ", "
        I = I + 1
    Loop
    m_List = sList & IIf(sList = "[", "'',]", "]")
    
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("displayrows") = ""
    m_dicTotal("textvariable") = m_Base.Name & "Var"
    m_dicTotal("tooltip") = o.Properties("ToolTipText")
    m_dicTotal("fg") = TranslateColor(o.Properties("ForeColor"))
    m_dicTotal("bg") = TranslateColor(o.Properties("BackColor"))
    m_dicTotal("bd") = "1"
    m_dicTotal("relief") = "SUNKEN"
    m_dicTotal("justify") = "'left'"
    m_dicTotal("state") = IIf(o.Properties("Enabled"), IIf(o.Properties("Locked") Or o.Properties("Style") = 2, "'readonly'", "'normal'"), "'disabled'")
    m_dicTotal("values") = m_Base.Name & "List"
    m_dicTotal("font") = "(" & U(CStr(o.Properties("Font")("Name"))) & "," & Round(o.Properties("Font")("Size")) & _
                        IIf(o.Properties("Font")("Bold"), ",'bold'", "") & IIf(o.Properties("Font")("Italic"), ",'italic'", "") & _
                        IIf(o.Properties("Font")("Underline"), ",'underline'", "") & IIf(o.Properties("Font")("Strikethrough"), ",'overstrike')", ")")
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
    m_dicTotal("cursor") = m_Base.GetCursorName(o.Properties("MousePointer"))
    m_dicTotal("exportselection") = "0"
    m_dicTotal("postcommand") = ""
    m_dicTotal("bindcommand") = m_Base.GetBindCommandStr(dMethods)
    
    'These are the default selections
    m_Base("x") = m_dicTotal("x")
    m_Base("y") = m_dicTotal("y")
    m_Base("width") = m_dicTotal("width")
    m_Base("height") = m_dicTotal("height")
    m_Base("relx") = m_dicTotal("relx")
    m_Base("rely") = m_dicTotal("rely")
    m_Base("relwidth") = m_dicTotal("relwidth")
    m_Base("relheight") = m_dicTotal("relheight")
    m_Base("values") = m_dicTotal("values")
    m_Base("textvariable") = m_dicTotal("textvariable")
    m_Base("exportselection") = m_dicTotal("exportselection")
    If m_dicTotal("tooltip") <> "" Then m_Base("tooltip") = m_dicTotal("tooltip")
    If o.Properties("ForeColor") <> vbWindowText Then m_Base("fg") = m_dicTotal("fg")
    If o.Properties("BackColor") <> vbWindowBackground Then m_Base("bg") = m_dicTotal("bg")
    If m_dicTotal("state") <> "'normal'" Then m_Base("state") = m_dicTotal("state")
    If m_dicTotal("font") <> "(" & U(GetDefaultFontName()) & ",9)" Then m_Base("font") = m_dicTotal("font")
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    If m_dicTotal("cursor") <> "" Then m_Base("cursor") = m_dicTotal("cursor")
    If m_Base.IsExistCommand(dMethods, "Change") Then m_Base("textvariable") = m_dicTotal("textvariable")
    If m_dicTotal("bindcommand") <> "" Then m_Base("bindcommand") = m_dicTotal("bindcommand")
    
    
    'Use the control's tag property to save some additional default properties,
    'Start character: p@, followed by property name, each property is separated by @, property value is optional, if there is a value, use k=v format
    s = Trim(o.Properties("Tag"))
    If Len(s) Then
        sa = Split(s, "@")
        If sa(0) = "p" Or sa(0) = "P" Then
            For I = 1 To UBound(sa)
                s = sa(I)
                idx = InStr(2, s, "=")
                If idx > 0 Then 'Has value
                    m_dicTotal(Left$(s, idx - 1)) = Mid$(s, idx + 1)
                    m_Base(Left$(s, idx - 1)) = m_dicTotal(Left$(s, idx - 1))
                ElseIf Len(s) > 0 And Len(m_dicTotal(s)) > 0 Then 'Only property, check if property is valid
                    m_Base(s) = m_dicTotal(s)
                End If
            Next
        End If
    End If
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    
    GetAttrValueList = 1
    Select Case sAttr
        Case "state"
            sa = Split("'normal','readonly','disabled'", ",")
        Case "exportselection"
            sa = Split("False,True", ",")
        Case "bindcommand"
            GetAttrValueList = 2
            sa = Split("<<ComboboxSelected>>", ",")
        Case Else
            GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End Select
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
    
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "displayrows":
        Tips = Tips & L("l_TipDisplayRows", "How many rows be displayed.")
    Case "state":
        Tips = Tips & L("l_TipStateHasRO", "State of the widget. They are 'normal', 'readonly', 'disabled'")
    Case "values":
        Tips = Tips & L("l_TipValues", "List of text of widget. This attribute set to a name of variable (type list).")
    Case "postcommand":
        Tips = Tips & L("l_TipPostCmd", "A procedure that is called immediately before displaying the values")
    Case "exportselection":
        Tips = Tips & L("l_TipExportSelection", "If set to True, the widget selection is linked to the Window Manager selection")
    Case Else:
        Tips = m_Base.Tips(sAttr)
    End Select
    
End Function

'Get the text of the first item in the dropdown list
Private Function GetFirstItemText() As String
    If Len(m_List) Then
        GetFirstItemText = Mid(m_List, 3, InStr(1, m_List, ",") - 4)
    Else
        GetFirstItemText = ""
    End If
End Function

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Combobox"
    m_Base.StyleName = "TCombobox"
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

'Update the configuration selected by the user to the object, the parameter is many attribute/value pairs separated by "|"
Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

'Modify or add a single configuration item, attribute/value separated by "|"
Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
End Sub

Private Sub Class_Terminate()
    Set m_dicTotal = Nothing
    Set m_Base = Nothing
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

'Used to change its default corresponding widget type, note the legality of the attribute list after modifying the widget type
Public Function SetWidgetType(sType As String, sStyleName As String)
    m_Base.ctlType = sType
    m_Base.StyleName = sStyleName
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
    vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescCombobox", "Combobox widget. TTK needed.")
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



