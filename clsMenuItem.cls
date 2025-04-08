VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Menu item class, representing each menu item

Private m_dicTotal As Dictionary        'Store all properties, including default values
Private m_Base As clsBaseControl        'Base control class
Private m_Visible As Boolean
Private m_IsSeparator As Boolean
Private m_IsCheckBox As Boolean

Private m_Childs() As Object
Private m_numChilds As Long
Private m_IdxCurChild As Long

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handler callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
'index: The index of this menu item under its parent menu, starting from 0, used for i18n setting label
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean, index As Long)
    Dim I As Long, extra As String, var As String, sTmp As String, s() As String, sr2 As String, sr3 As String
    
    extra = IIf(Len(m_Base("tearoff")), ", tearoff=" & m_Base("tearoff"), "")
    
    If m_numChilds > 0 Then
        sOut.Append vbCrLf & "        self." & m_Base.Name & " = Menu(self." & m_Base.Parent & extra & ")"
        For I = 0 To m_numChilds - 1                                            'Recursively generate code
            m_Childs(I).toString sOut, sCmd, sI18n, rel, usettk, I
        Next
        
        If m_Visible Then                                                       'If Visible=0, it generally means this menu is set as a right-click popup menu
            m_Base("variable") = ""
            m_Base("command") = ""
            extra = GetExtraParams()
            sOut.Append "        self." & m_Base.Parent & ".add_cascade(menu=" & "self." & m_Base.Name & IIf(Len(extra), ", " & extra, "") & ")"
        End If
    ElseIf m_IsSeparator Then                                                   'Menu separator
        sOut.Append "        self." & m_Base.Parent & ".add_separator()"
    ElseIf m_IsCheckBox Then                                                    ' Menu add checkbox
        If m_Base("variable") <> "" Then
            sOut.Append "        self." & m_Base("variable") & " = StringVar()"
            sOut.Append "        self." & m_Base("variable") & ".set(1)"
        End If
        
        If m_Base("command") = "" Then m_Base("command") = m_Base.Name & "_Cmd"
        extra = GetExtraParams()
        sOut.Append "        self." & m_Base.Parent & ".add_checkbutton(" & extra & ")"
        If Len(m_Base("command")) Then
            sCmd.Append m_Base.CreateFuncDefOOP(m_Base("command"), "event=None")
        End If
    Else
        var = m_Base("variable")
        m_Base("variable") = ""
        If m_Base("command") = "" Then m_Base("command") = m_Base.Name & "_Cmd"
        extra = GetExtraParams()
        m_Base("variable") = var
        sOut.Append "        self." & m_Base.Parent & ".add_command(" & extra & ")"
        If Len(m_Base("command")) Then
            sCmd.Append m_Base.CreateFuncDefOOP(m_Base("command"), "event=None")
        End If
    End If
    
    'Internationalization translation
    If (Not m_IsSeparator) And Len(m_Base("label")) > 0 Then
        sI18n.Append "        self." & m_Base.Parent & ".entryconfigure(" & index & ", label=_(" & U(m_Base("label")) & "))"
    End If
    
    'If there are other event handlers that need to be bound using the bind statement
    If m_numChilds = 0 And Not m_IsSeparator And Len(m_Base("bindcommand")) Then
        sTmp = UnQuote(m_Base("bindcommand")) 'Automatically remove brackets, if any
        
        s = Split(sTmp, ",")
        For I = 0 To UBound(s)
            s(I) = Trim(s(I))
            If Left(s(I), 1) = "<" And Right(s(I), 1) = ">" Then
                sOut.Append "        self." & WTOP & ".bind_all('" & s(I) & "', " & "self." & m_Base("command") & ")"
                
                sr2 = Mid$(s(I), Len(s(I)) - 1, 1)
                sr3 = Mid$(s(I), Len(s(I)) - 2, 1)
                
                'Python is case-sensitive, and so are the corresponding shortcuts. If the shortcut key contains a letter key, bind the corresponding uppercase/lowercase as well
                If sr3 = "-" Then
                    If sr2 >= "a" And sr2 <= "z" Then
                        s(I) = Left$(s(I), Len(s(I)) - 2) & UCase$(sr2) & ">"
                        sOut.Append "        self." & WTOP & ".bind_all('" & s(I) & "', " & "self." & m_Base("command") & ")"
                    ElseIf sr2 >= "A" And sr2 <= "Z" Then
                        s(I) = Left$(s(I), Len(s(I)) - 2) & LCase$(sr2) & ">"
                        sOut.Append "        self." & WTOP & ".bind_all('" & s(I) & "', " & "self." & m_Base("command") & ")"
                    End If
                End If
            End If
        Next
    End If
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(o As Object, Optional parentWidth As Long, Optional parentHeight As Long, Optional dMethods As Dictionary)
    
    Dim s As String
    
    m_Base.SetVbWidgetInstance o
    
    m_Base.Name = o.Properties("Name")
    
    m_Visible = o.Properties("Visible")
    m_IsSeparator = (o.Properties("Caption") = "-")
    m_IsCheckBox = (o.Properties("Checked") = True)
    
    'These are all the default values
    m_dicTotal("label") = Replace(o.Properties("Caption"), "&", "")
    m_dicTotal("fg") = ""
    m_dicTotal("bg") = ""
    m_dicTotal("bd") = ""
    m_dicTotal("tearoff") = "0"
    m_dicTotal("relief") = "RAISED"
    m_dicTotal("state") = IIf(o.Properties("Enabled"), "'normal'", "'disabled'")
    m_dicTotal("underline") = IIf(InStr(1, o.Properties("Caption"), "&"), InStr(1, o.Properties("Caption"), "&") - 1, "-1")
    m_dicTotal("variable") = o.Properties("Name") & "Var"
    m_dicTotal("font") = ""
    m_dicTotal("accelerator") = TransShortcut(o.Properties("Shortcut"))
    m_dicTotal("command") = o.Properties("Name") & "_Cmd"
    m_dicTotal("postcommand") = ""
    m_dicTotal("bindcommand") = ""
    
    m_Base("tearoff") = m_dicTotal("tearoff")
    m_Base("variable") = m_dicTotal("variable")
    m_Base("command") = m_dicTotal("command")
    m_Base("label") = m_dicTotal("label")
    If m_dicTotal("state") <> "'normal'" Then m_Base("state") = m_dicTotal("state")
    If m_dicTotal("underline") <> "-1" Then m_Base("underline") = m_dicTotal("underline")
    
    If m_dicTotal("accelerator") <> "" Then                                     'Shortcut key binding
        s = m_dicTotal("accelerator")
        m_Base("accelerator") = s
        s = "'<" & Replace(s, "Ctrl", "Control") & ">'"
        m_Base("bindcommand") = s
    End If
    
End Sub

'Convert VB shortcut enumeration value to a readable string
Private Function TransShortcut(nsc As Long) As String
    Select Case nsc
    Case 0
        TransShortcut = ""
    Case vbextMenuShortcutCtrlA To vbextMenuShortcutCtrlZ                       '1 - 26
        TransShortcut = "Ctrl-" & Chr(nsc + 64)
    Case vbextMenuShortcutF1 To vbextMenuShortcutF12                            ' 27 - 38
        TransShortcut = "F" & CStr(nsc - 26)
    Case vbextMenuShortcutCtrlF1 To vbextMenuShortcutCtrlF12                    ' 39 - 50
        TransShortcut = "Ctrl-F" & CStr(nsc - 38)
    Case vbextMenuShortcutShiftF1 To vbextMenuShortcutShiftF12                  ' 51 - 62
        TransShortcut = "Shift-F" & CStr(nsc - 50)
    Case vbextMenuShortcutCtrlShiftF1 To vbextMenuShortcutCtrlShiftF12          '  63 - 74
        TransShortcut = "Ctrl-Shift-F" & CStr(nsc - 62)
    Case vbextMenuShortcutCtrlIns                                               '= 75
        TransShortcut = "Ctrl-Insert"
    Case vbextMenuShortcutShiftIns                                              '= 76
        TransShortcut = "Shift-Insert"
    Case vbextMenuShortcutDel                                                   '= 77
        TransShortcut = "Delete"
    Case vbextMenuShortcutShiftDel                                              ' = 78
        TransShortcut = "Shift-Delete"
    Case vbextMenuShortcutAltBksp                                               ' = 79
        TransShortcut = "Alt-BackSpace"
    Case Else
        TransShortcut = ""
    End Select
    
End Function

'Generate a list of other parameters chosen by the user in addition to the mandatory parameters
Public Function GetExtraParams() As String
    Dim cfg As Variant, k As Variant, ks As Variant, sValue As String
    
    Set ks = m_Base.Keys
    
    For Each k In ks
        If isExtra(k) And Len(m_Base(k)) Then
            'Attributes that need to be enclosed in quotes, if the user forgets, add them here automatically
            If k = "label" Then
                sValue = U(m_Base(k))
            ElseIf InStr(1, " fg, bg, state, accelerator, ", " " & k & ",") Then
                sValue = Quote(m_Base(k))
            Else
                sValue = m_Base(k)
            End If
            GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & k & "=" & sValue
        End If
    Next
    
    If Len(m_Base("command")) Then
        GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & "command=self." & m_Base("command")
    End If
    If Len(m_Base("variable")) Then
        GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & "variable=self." & m_Base("variable")
    End If
    If Len(m_Base("font")) Then
        GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & "font=self." & m_Base.Name & "Font"
    End If
End Function

Private Function isExtra(ByVal sK As String) As Boolean
    isExtra = (InStr(1, " tearoff, variable, command, postcommand, bindcommand, font, ", Space(1) & sK & ",") <= 0)
End Function

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If sAttr = "tearoff" Then
        GetAttrValueList = 1
        sa = Split("1,0", ",")
    ElseIf sAttr = "bindcommand" Then
        GetAttrValueList = 2
        sa = Split("<<MenuSelect>>", ",")
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
    ElseIf sAttr = "accelerator" Then
        Tips = sAttr & vbCrLf & L("l_TipAcceleratorMenu", "To display a keystroke combination on the right side of a menu choice.")
    Else
        Tips = m_Base.Tips(sAttr)
    End If
    
End Function

Private Sub Class_Initialize()
    
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Menu"
    m_Base.StyleName = ""
    m_Base.Parent = "MainMenu"
    Erase m_Childs
    m_numChilds = 0
    m_IdxCurChild = 0
    
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
    EnableOutByMainForm = False
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    'm_CanbeOutByMainForm = bEnable
End Property

'Object serialization function
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescMenuItem", "Menu item in Menu widget.")
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

