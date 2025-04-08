VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Notebook class, used to set and maintain the corresponding properties and values, and generate the corresponding PYTHON TKinter code

Private m_dicTotal As Dictionary       'Store all properties, including default values
Private m_Base As clsBaseControl       'Base control class
Private m_TabTitles() As String        'Tab titles, starting from index 1, index 0 is not used
Private m_Tabs() As Object             'Tab objects, starting from index 1, index 0 is not used
Private m_nTabs As Long                'Number of tabs
Private m_CanbeOutByMainForm As Boolean

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handling callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    m_Base.toString sOut, sCmd, sI18n, rel, usettk
    CreateTabWidgets sOut, sCmd, sI18n, rel, usettk
End Sub

Private Sub CreateTabWidgets(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    Dim I As Long, nTabs As Long, s As String
    Dim bTabExisted As Boolean, sCaption As String
    
    nTabs = IIf(m_nTabs > UBound(m_Tabs), m_nTabs, UBound(m_Tabs))
    
    For I = 1 To nTabs  'TAB index starts from 1
        If I > UBound(m_Tabs) Then
            bTabExisted = False
        ElseIf m_Tabs(I) Is Nothing Then
            bTabExisted = False
        Else
            bTabExisted = True
        End If
        
        'Prefer to use the tab title set in TabStrip, if empty, use the title of the Frame control.
        If bTabExisted Then
            If I > m_nTabs Then
                sCaption = m_Tabs(I).Caption
            ElseIf Len(m_TabTitles(I)) = 0 Then
                sCaption = m_TTabs(I).Caption
            ElseIf I <= m_nTabs Then
                sCaption = m_TabTitles(I)
            End If
        ElseIf I <= m_nTabs Then
            sCaption = m_TabTitles(I)
        End If
        If Len(sCaption) = 0 Then sCaption = "Tab" & I
        
        CreateSingleTabWidget sOut, sCmd, sI18n, rel, usettk, I, bTabExisted, sCaption
    Next
End Sub

'Create controls inside a specific tab
Private Sub CreateSingleTabWidget(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, _
                       ByVal usettk As Boolean, ByVal nTabIdx As Long, ByVal bUseTabObj As Boolean, ByVal sCaption As String)
    Dim sTabName As String
    
    If bUseTabObj Then
        sTabName = "self." & m_Tabs(nTabIdx).Name
        sOut.Append vbCrLf & Space(8) & sTabName & " = Frame(self." & m_Base.Name & ")"
        m_Tabs(nTabIdx).toString sOut, sCmd, sI18n, rel, usettk
    Else
        sTabName = "self." & m_Base.Name & "__Tab" & nTabIdx
        sOut.Append vbCrLf & Space(8) & sTabName & " = Frame(self." & m_Base.Name & ")"
        sOut.Append Space(8) & sTabName & "Lbl = Label(" & sTabName & ", text='Please add widgets in code.')"
        sOut.Append Space(8) & sTabName & "Lbl.place(relx=0.1,rely=0.5)"
    End If
    
    sOut.Append "        self." & m_Base.Name & ".add(" & sTabName & ", text=" & U(sCaption) & ")"
    'Internationalization translation
    sI18n.Append "        self." & m_Base.Name & ".tab(" & (nTabIdx - 1) & ", text=_(" & U(sCaption) & "))"
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    Dim I As Long, oiTabs As ITabs, s As String, sa() As String, idx As Long
    
    m_Base.SetVbWidgetInstance o
    
    m_Base.Name = o.Properties("Name")
    
    Set oiTabs = o.Properties("Tabs").object
    m_nTabs = oiTabs.Count
    
    ReDim m_TabTitles(m_nTabs) As String
    For I = 1 To m_nTabs
        m_TabTitles(I) = oiTabs.Item(I).Caption  'TAB index starts from 1
    Next
    
    'These are all the default values
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("relief") = "FLAT"
    m_dicTotal("tooltip") = o.Properties("ToolTipText")
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
    m_dicTotal("cursor") = m_Base.GetCursorName(o.Properties("MousePointer"))
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
    If m_dicTotal("tooltip") <> "" Then m_Base("tooltip") = m_dicTotal("tooltip")
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    If m_dicTotal("cursor") <> "" Then m_Base("cursor") = m_dicTotal("cursor")
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

'Add a tab
Public Sub AddTab(oTab As Object, nTabIdx As Long)
    
    If nTabIdx > UBound(m_Tabs) Then
        ReDim Preserve m_Tabs(nTabIdx) As Object
    End If
    
    Set m_Tabs(nTabIdx) = oTab
    
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
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
    Tips = m_Base.Tips(sAttr)
End Function

'Update the configuration selected by the user to the object, the parameter is many attribute/value pairs separated by "|"
Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

'Modify or add a single configuration item, attribute/value separated by "|"
Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
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

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Notebook"
    m_Base.StyleName = "TNotebook"
    ReDim m_TabTitles(0) As String
    ReDim m_Tabs(0) As Object
    m_CanbeOutByMainForm = True
End Sub

Private Sub Class_Terminate()
    Set m_dicTotal = Nothing
    Set m_Base = Nothing
    Erase m_Tabs
    Erase m_TabTitles
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
    Description = L("l_DescNotebook", "Notebook widget. TTK needed.")
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



