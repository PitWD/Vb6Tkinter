VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBaseControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements common code for all control classes. Since VB does not support true inheritance, other control classes will embed this class instead of inheriting it.

Private m_dic As Dictionary                                                     'Current properties/values to generate code
Private m_Type As String                                                        'Corresponding PYTHON control type
Private m_Name As String                                                        'Control name
Private m_Parent As String
Private m_Value As String                                                       ' Control value (if any)
Private m_StyleName As String                                                   'Style base class
Private m_ScaleMode As Long
Private m_vbWidgetInstance As Object                                            'Reference to the VB control

'Generates PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handler callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean, Optional sOtherParams As String = "")
    Dim s() As String, I As Long, extra As String
    Dim sTmp As String, sStyle As String, sCmdName As String, sUnderlineCmd As String
    
    'Some controls only exist in tk, even if ttk is enabled, styles cannot be applied
    usettk = IIf(Len(m_StyleName), usettk, False)
    
    'If variable binding is needed, create the corresponding variable first
    If Len(Dic("textvariable")) Then
        sOut.Append "        self." & Dic("textvariable") & " = StringVar(value=" & U(Dic("text")) & ")"
    End If
    If Len(Dic("variable")) Then
        If m_Type = "Radiobutton" Then  'The variable for a group of radio buttons is the same, do not create it repeatedly
            sTmp = "        self." & Dic("variable") & " = StringVar()"
            If Not sOut.ExistString(sTmp) Then
                sOut.Append sTmp
            End If
        ElseIf m_Type = "Checkbutton" Or m_Type = "Progressbar" Then 'The variable for check buttons/progress bars is better to use IntVar
            sTmp = IIf(IsNumeric(m_Value), m_Value, "0")
            sOut.Append "        self." & Dic("variable") & " = IntVar(value=" & sTmp & ")"
        Else
            sOut.Append "        self." & Dic("variable") & " = StringVar(value=" & U(m_Value) & ")"
        End If
    End If
    If Len(Dic("listvariable")) Then
        sOut.Append "        self." & Dic("listvariable") & " = StringVar(value=" & U(m_Value) & ")"
    End If
    If Not usettk And Len(Dic("font")) Then
        If Left$(Dic("font"), 1) = "(" Then
            sOut.Append "        self." & m_Name & "Font = Font(font=" & Dic("font") & ")"
        Else
            sOut.Append "        self." & m_Name & "Font = Font(font=(" & Dic("font") & "))"
        End If
    End If
    
    If Len(Dic("columns")) Then
        sOut.Append "        self." & Dic("columns") & " = []    " & L("l_cmtTodoCols", "#TODO Add list of titles here, first column fixed for tree view.")
    End If
    If Len(Dic("displaycolumns")) And InStr(1, Dic("displaycolumns"), "#all") <= 0 Then
        sOut.Append "        self." & Dic("displaycolumns") & " = []    " & L("l_cmtTodoDisCols", "#TODO Add list of titles will be displayed, first column fixed for tree view.")
    End If
    
    'Convert displayrows to height attribute for combo box
    If Len(Dic("displayrows")) Then
        sOtherParams = sOtherParams & IIf(Len(sOtherParams), ", ", "") & "height=" & Dic("displayrows")
    End If
    
    If usettk Then                                                              'Create STYLE object
        sStyle = GetStyleParams(usettk)
        If Len(sStyle) Then
            sOut.Append "        self.style.configure('T" & m_Name & "." & m_StyleName & "', " & sStyle & ")"
            If m_Type = "LabelFrame" Then ' LabelFrame's font and foreground color need to be set to Label
                sOut.Append "        self.style.configure('T" & m_Name & "." & m_StyleName & ".Label" & "', " & sStyle & ")"
            End If
            sOtherParams = sOtherParams & IIf(Len(sOtherParams), ", ", "") & "style='T" & m_Name & "." & m_StyleName & "'"
        End If
    End If
    
    extra = GetExtraParams(usettk)
    extra = extra & IIf(Len(extra) > 0 And Len(sOtherParams) > 0, ", ", "") & sOtherParams
    'Create this control instance
    sOut.Append "        self." & m_Name & " = " & m_Type & "(self." & m_Parent & IIf(Len(extra), ", ", "") & extra & ")"
    
    'If needed, add a few simple "wrapper methods" to the instance for accessing textvariable, such as setText/text methods for Entry
    If Len(Dic("textvariable")) Then
        sOut.Append "        self." & m_Name & ".setText = lambda x: self." & Dic("textvariable") & ".set(x)"
        sOut.Append "        self." & m_Name & ".text = lambda : self." & Dic("textvariable") & ".get()"
        'Text inside Entry/Text/Combobox generally does not need translation
        If m_Type <> "Entry" And m_Type <> "Text" And m_Type <> "Combobox" Then
            sI18n.Append "        self." & m_Name & ".setText(_(" & U(Dic("text")) & "))"
        End If
    End If
    
    'If needed, add a few simple "wrapper methods" to the instance for accessing variable, such as setValue/value methods for CheckButton
    If Len(Dic("variable")) Then
        If m_Type = "Radiobutton" Then 'Radio button
            If Len(Dic("value")) Then
                sOut.Append "        self." & m_Name & ".setValue = lambda x: self." & Dic("variable") & ".set(" & Dic("value") & " if x else '')"
                sOut.Append "        self." & m_Name & ".value = lambda : 1 if self." & Dic("variable") & ".get() == " & Dic("value") & " else 0"
                If m_vbWidgetInstance.Properties("Value") And Len(Dic("value")) Then 'Select one by default
                    sOut.Append "        self." & m_Name & ".setValue(1)"
                End If
            Else
                sOut.Append "        self." & m_Name & ".setValue = lambda x: self." & Dic("variable") & ".set(x)"
                sOut.Append "        self." & m_Name & ".value = lambda : self." & Dic("variable") & ".get()"
            End If
        ElseIf m_Type = "Checkbutton" Or m_Type = "Progressbar" Then  'Check button and progress bar
            sOut.Append "        self." & m_Name & ".setValue = lambda x: self." & Dic("variable") & ".set(x)"
            sOut.Append "        self." & m_Name & ".value = lambda : self." & Dic("variable") & ".get()"
        End If
    End If
    
    'If tooltip text is set, create Tooltip instance
    If Len(Dic("tooltip")) Then
        sOut.Append "        self." & m_Name & "Tooltip = Tooltip(self." & m_Name & ", " & Quote(Dic("tooltip")) & ")"
        sI18n.Append "        self." & m_Name & "Tooltip.text = _(" & Quote(Dic("tooltip")) & ")"
    End If
    
    'Place control
    sOut.Append "        self." & m_Name & ".place(" & GetPositionParams(rel) & ")"
    
    'Create event handler function skeletons (if any), if the event contains a dot, it indicates calling a system function and does not need a function body
    If Len(Dic("command")) > 0 And InStr(1, Dic("command"), ".") <= 0 And InStr(1, Dic("command"), "lambda ") <= 0 Then
        sCmd.Append CreateFuncDefOOP(Dic("command"), "event=None")
    End If
    If Len(Dic("postcommand")) Then
        sCmd.Append CreateFuncDefOOP(Dic("postcommand"), "event=None")
    End If
    
    'Handle underline shortcuts
    If Len(Dic("underline")) > 0 And Dic("underline") <> "-1" And IsNumeric(Dic("underline")) Then
        If m_Type = "Button" Or m_Type = "Checkbutton" Or m_Type = "Radiobutton" Then
            sUnderlineCmd = "lambda e: self." & m_Name & ".focus_set() or self." & m_Name & ".invoke()"
        ElseIf Len(Dic("command")) > 0 Then
            sUnderlineCmd = "self." & Dic("command")
        Else
            sUnderlineCmd = "lambda e: self." & m_Name & ".focus_set()"
        End If
        
        If Len(sUnderlineCmd) Then
            sTmp = Dic("text")
            If Len(sTmp) = 0 Then sTmp = Dic("label")
            If Len(sTmp) And CLng(Dic("underline")) < Len(sTmp) Then
                sOut.Append "        self." & WTOP & ".bind_all('<Alt-" & Mid(sTmp, CLng(Dic("underline")) + 1, 1) & ">', " & sUnderlineCmd & ")"
                If Mid(sTmp, CLng(Dic("underline")) + 1, 1) >= "a" And Mid(sTmp, CLng(Dic("underline")) + 1, 1) <= "z" Then
                    sOut.Append "        self." & WTOP & ".bind_all('<Alt-" & UCase(Mid(sTmp, CLng(Dic("underline")) + 1, 1)) & ">', " & sUnderlineCmd & ")"
                ElseIf Mid(sTmp, CLng(Dic("underline")) + 1, 1) >= "A" And Mid(sTmp, CLng(Dic("underline")) + 1, 1) <= "Z" Then
                    sOut.Append "        self." & WTOP & ".bind_all('<Alt-" & LCase(Mid(sTmp, CLng(Dic("underline")) + 1, 1)) & ">', " & sUnderlineCmd & ")"
                End If
            End If
        End If
    End If
    
    If Len(Dic("bindcommand")) Then                                             'Other event handlers that need to use bind statements
        sTmp = Dic("bindcommand")
        sTmp = Replace(sTmp, "'", "")                                           'Automatically remove brackets and spaces if any
        sTmp = Replace(sTmp, Chr(34), "")
        sTmp = Replace(sTmp, " ", "")
        
        s = Split(sTmp, ",")
        For I = 0 To UBound(s)
            s(I) = Trim(s(I))
            If s(I) = "<Change>" Then 'Specially handle this custom event, use control variable monitoring to simulate
                If (m_Type = "Combobox" Or m_Type = "Entry" Or m_Type = "Label") And (Len(Dic("textvariable")) > 0) Then
                    sCmdName = m_Name & "_Change"
                    sOut.Append "        self." & Dic("textvariable") & ".trace('w', self." & sCmdName & ")"
                    sCmd.Append CreateFuncDefOOP(sCmdName, "*args")
                End If
            ElseIf Left(s(I), 1) = "<" And Right(s(I), 1) = ">" Then
                
                sCmdName = m_Name & "_" & Replace(Replace(Replace(s(I), "<", ""), ">", ""), "-", "_")
                
                sOut.Append "        self." & m_Name & ".bind('" & s(I) & "', self." & sCmdName & ")"
                sCmd.Append CreateFuncDefOOP(sCmdName, "event")
                
                'Python is case-sensitive, the corresponding shortcuts are the same. If the set shortcut contains letter keys, bind the corresponding uppercase/lowercase together
                If Right(s(I), 3) >= "-a>" And Right(s(I), 3) <= "-z>" Then
                    
                    s(I) = Left(s(I), Len(s(I)) - 2) & UCase(Mid(s(I), Len(s(I)) - 1, 1)) & ">" 'Change to uppercase
                    sOut.Append "        self." & m_Name & ".bind('" & s(I) & "', self." & sCmdName & ")"
                    
                ElseIf Right(s(I), 3) >= "-A>" And Right(s(I), 3) <= "-Z>" Then
                    
                    s(I) = Left(s(I), Len(s(I)) - 2) & LCase(Mid(s(I), Len(s(I)) - 1, 1)) & ">" 'Change to lowercase
                    sOut.Append "        self." & m_Name & ".bind('" & s(I) & "', self." & sCmdName & ")"
                    
                End If
            End If
        Next
    End If
    
End Sub

'Generates the corresponding control position information based on rel (whether to use relative coordinates)
Public Function GetPositionParams(rel As Boolean) As String
    If rel Then
        GetPositionParams = "relx=" & Commas2Points(Dic("relx")) & ", rely=" & Commas2Points(Dic("rely")) & _
            ", relwidth=" & Commas2Points(Dic("relwidth"))
        If m_Type <> "Combobox" Then 'Combobox does not need to set the height attribute
            GetPositionParams = GetPositionParams & ", relheight=" & Commas2Points(Dic("relheight"))
        End If
    Else
        GetPositionParams = "x=" & Commas2Points(Dic("x")) & ", y=" & Commas2Points(Dic("y")) & _
            ", width=" & Commas2Points(Dic("width"))
        If m_Type <> "Combobox" Then 'Combobox does not need to set the height attribute
            GetPositionParams = GetPositionParams & ", height=" & Commas2Points(Dic("height"))
        End If
    End If
End Function

'Generates a list of optional parameters selected by the user, in addition to the required parameters
Public Function GetExtraParams(usettk As Boolean) As String
    
    Dim cfg As Variant, k As Variant, ks As Variant, sValue As String, s As New cStrBuilder
    
    Set ks = m_dic.Keys
    
    For Each k In ks
        If isExtra(k, usettk) And Len(Dic(k)) Then
            'Attributes that need to be enclosed in quotes. If the user forgets, add them here automatically
            If k = "text" Or k = "label" Then
                If m_Type = "Entry" And Len(Dic("textvariable")) > 0 And usettk Then 'In ttk mode, Entry uses textvariable to display characters instead of using the text attribute
                    'Ignore the text attribute
                    sValue = ""
                Else
                    sValue = U(Dic(k))
                End If
            ElseIf InStr(1, ",fg,bg,anchor,justify,show,state,activestyle,labelanchor,mode,cursor,highlightbackground,highlightcolor,selectbackground,selectforeground,", _
                        "," & k & ",") Then
                sValue = Quote(Dic(k))
            Else
                sValue = Dic(k)
            End If
            
            If Len(sValue) Then
                s.Append k & "=" & sValue
            End If
        End If
    Next
    
    If Len(Dic("columns")) Then s.Append "columns=self." & Dic("columns")
    If Len(Dic("displaycolumns")) Then
        If InStr(1, Dic("displaycolumns"), "#all") <= 0 Then
            s.Append "displaycolumns=self." & Dic("displaycolumns")
        Else
            s.Append "displaycolumns='#all'"
        End If
    End If
    
    If Len(Dic("textvariable")) Then s.Append "textvariable=self." & Dic("textvariable")
    If Len(Dic("variable")) Then s.Append "variable=self." & Dic("variable")
    If Len(Dic("listvariable")) Then s.Append "listvariable=self." & Dic("listvariable")
    If Len(Dic("values")) Then s.Append "values=self." & Dic("values")
    
    If Len(Dic("command")) Then
        If InStr(1, Dic("command"), "lambda ") > 0 Then 'Anonymous function, no need to add self.
            s.Append "command=" & Dic("command")
        Else
            s.Append "command=self." & Dic("command")
        End If
    End If

If Len(Dic("postcommand")) Then s.Append "postcommand=self." & Dic("postcommand")
If Len(Dic("xscrollcommand")) Then s.Append "xscrollcommand=self." & Dic("xscrollcommand")
If Len(Dic("yscrollcommand")) Then s.Append "yscrollcommand=self." & Dic("yscrollcommand")
If Len(Dic("font")) Then
    If usettk Then 'In TTK mode, most control fonts are written in styles, but the fonts for Entry/Combobox need to be written in the creation function
        If m_Type = "Entry" Or m_Type = "Combobox" Then
            s.Append "font=" & IIf(Left$(Dic("font"), 1) = "(", Dic("font"), "(" & Dic("font") & ")") 'Automatically add parentheses
        End If
    Else
        s.Append "font=self." & m_Name & "Font"
    End If
End If

GetExtraParams = s.toString(", ")
End Function

'Determines whether an attribute is an extra parameter, more attributes are considered based on TTK
Private Function isExtra(ByVal sK As String, usettk As Boolean) As Boolean
    'Attributes in these two string lists are not extra parameters, the rest are extra parameters
    Const NOT_EXTRA_STRING As String = ",x,y,relx,rely,width,height,relwidth,relheight,command,bindcommand,xscrollcommand," & _
        "yscrollcommand,postcommand,font,textvariable,variable,listvariable,values,displayrows,columns,displaycolumns,tooltip,"
    Const NOT_EXTRA_STRING_TTK As String = NOT_EXTRA_STRING & ",fg,bg,bd,relief,activerelief,overrelief,anchor,jump,indicatoron,resolution,digits," & _
        "sliderlength,sliderrelief,showvalue,tickinterval,"
    
    If usettk Then
        isExtra = (InStr(1, NOT_EXTRA_STRING_TTK, "," & sK & ",") <= 0)
    Else
        isExtra = (InStr(1, NOT_EXTRA_STRING, "," & sK & ",") <= 0)
    End If
End Function

'If TTK extension is used, use this function to get TTK related parameters and create a valid string parameter list
Private Function GetStyleParams(usettk As Boolean) As String
    Dim s As New cStrBuilder
    
    If Len(Dic("relief")) Then s.Append "relief=" & Dic("relief")
    If Len(Dic("activerelief")) Then s.Append "activerelief=" & Dic("activerelief")
    If Len(Dic("overrelief")) Then s.Append "overrelief=" & Dic("overrelief")
    If Len(Dic("anchor")) Then s.Append "anchor=" & Quote(Dic("anchor"))
    If Len(Dic("fg")) Then s.Append "foreground=" & Quote(Dic("fg"))
    If Len(Dic("bg")) Then s.Append "background=" & Quote(Dic("bg"))
    If Len(Dic("bd")) Then s.Append "borderwidth=" & Dic("bd")
    If Len(Dic("jump")) Then s.Append "jump=" & Dic("jump")
    If Len(Dic("indicatoron")) Then s.Append "indicatoron=" & Dic("indicatoron")
    If Len(Dic("font")) Then
        'For ttk.Entry, font must be written in the constructor to be effective, for ttk.LabelFrame, the font must be set to the .Label property to be effective
        If m_Type <> "Entry" And m_Type <> "Combobox" Then
            If Left$(Dic("font"), 1) = "(" Then
                s.Append "font=" & Dic("font")
            Else  'If the user forgot to add parentheses, add them here
                s.Append "font=(" & Dic("font") & ")"
            End If
        End If
    End If
    If Len(Dic("showvalue")) Then s.Append "showvalue=" & Dic("showvalue")
    If Len(Dic("tickinterval")) Then s.Append "tickinterval=" & Dic("tickinterval")
    If Len(Dic("sliderrelief")) Then s.Append "sliderrelief=" & Dic("sliderrelief")
    If Len(Dic("sliderlength")) Then s.Append "sliderlength=" & Dic("sliderlength")
    If Len(Dic("digits")) Then s.Append "digits=" & Dic("digits")
    
    GetStyleParams = s.toString(", ")
    
End Function

'Automatically fills the bindcommand field based on function declarations in the code module
'dMethods: Dictionary with control names as keys and control event handler function names as comma-separated strings
Public Function GetBindCommandStr(dMethods As Dictionary) As String
    Dim s As String, sOut As cStrBuilder
    If Not dMethods.Exists(m_Name) Then Exit Function
    
    Set sOut = New cStrBuilder
    s = dMethods.Item(m_Name)
    
    'For these controls, the Click event is better set by the command attribute, no need to bind
    If m_Type = "Button" Or m_Type = "Checkbutton" Or m_Type = "Radiobutton" Then
        If InStr(1, s, "," & m_Name & "_MouseDown,") > 0 Then sOut.Append "<Button-1>"
    ElseIf m_Type = "Listbox" Then
        If InStr(1, s, "," & m_Name & "_Click,") > 0 Then sOut.Append "<<ListboxSelect>>"
    ElseIf m_Type = "Combobox" Then
        If InStr(1, s, "," & m_Name & "_Change,") > 0 Then sOut.Append "<<ComboboxSelected>>"
    ElseIf m_Type = "Text" Then
        If InStr(1, s, "," & m_Name & "_Change,") > 0 Then sOut.Append "<<Modified>>"
    ElseIf m_Type <> "Notebook" Then
        If InStr(1, s, "," & m_Name & "_Click,") > 0 Or InStr(1, s, "," & m_Name & "_MouseDown,") > 0 Then sOut.Append "<Button-1>"
    End If
    
    If InStr(1, s, "," & m_Name & "_DblClick,") > 0 Then sOut.Append "<Double-Button-1>"
    If InStr(1, s, "," & m_Name & "_Resize,") > 0 Then sOut.Append "<Configure>"
    If InStr(1, s, "," & m_Name & "_GotFocus,") > 0 Then sOut.Append "<FocusIn>"
    If InStr(1, s, "," & m_Name & "_LostFocus,") > 0 Then sOut.Append "<FocusOut>"
    If InStr(1, s, "," & m_Name & "_KeyPress,") > 0 Or InStr(1, s, "," & m_Name & "_KeyDown,") > 0 Then sOut.Append "<KeyPress>"
    If InStr(1, s, "," & m_Name & "_KeyUp,") > 0 Then sOut.Append "<KeyRelease>"
    If InStr(1, s, "," & m_Name & "_MouseUp,") > 0 Then sOut.Append "<ButtonRelease-1>"
    If InStr(1, s, "," & m_Name & "_Enter,") > 0 Or InStr(1, s, "," & m_Name & "_MouseMove,") > 0 Then sOut.Append "<Motion>"
    If InStr(1, s, "," & m_Name & "_Leave,") > 0 Then sOut.Append "<Leave>"
    
    If m_Type <> "Combobox" And m_Type <> "Text" And InStr(1, s, "," & m_Name & "_Change,") > 0 Then
        sOut.Append "<Change>"  'This event is handled internally, not a standard tk event
    End If
    
    If m_Type = "Treeview" Then
        If InStr(1, s, "," & m_Name & "_NodeClick,") > 0 Then sOut.Append "<<TreeviewSelect>>"
        If InStr(1, s, "," & m_Name & "_Collapse,") > 0 Then sOut.Append "<<TreeviewClose>>"
        If InStr(1, s, "," & m_Name & "_Expand,") > 0 Then sOut.Append "<<TreeviewOpen>>"
    ElseIf m_Type = "Notebook" Then
        If (InStr(1, s, "," & m_Name & "_BeforeClick,") > 0) Or (InStr(1, s, "," & m_Name & "_Click,") > 0) Then sOut.Append "<<NotebookTabChanged>>"
    End If
    
    GetBindCommandStr = sOut.toString(",")
End Function

Public Function IsExistCommand(dMethods As Dictionary, cmdTxt As String) As Boolean
    If dMethods.Exists(m_Name) Then IsExistCommand = (InStr(1, dMethods.Item(m_Name), "," & m_Name & "_" & cmdTxt & ",") > 0)
End Function

'Set/Get the value of the dictionary, Dic() is the default property of clsBaseControl 'Attribute Dic.VB_UserMemId = 0'
Public Property Get Dic(ByVal sKey As String) As String
Attribute Dic.VB_UserMemId = 0
    If m_dic.Exists(sKey) Then Dic = m_dic(sKey)
End Property

Public Property Let Dic(ByVal sKey As String, ByVal sValue As String)
    m_dic(sKey) = sValue
End Property

Public Sub Remove(ByVal sKey As String)
    m_dic.Remove (sKey)
End Sub

'Check if this control has the specified attribute
Public Function hasAttribute(sAttr As String) As Boolean
    If m_dic.Exists(sAttr) Then
        hasAttribute = True
    Else
        hasAttribute = False
    End If
End Function

'Get the current value of the specified attribute for this control, return an empty string if not set
Public Function GetAttrCurrentValue(sAttr As String) As String
    If m_dic.Exists(sAttr) Then
        GetAttrCurrentValue = m_dic(sAttr) & "" 'Concatenating with an empty string to automatically convert other types to string
    Else
        GetAttrCurrentValue = ""
    End If
End Function

'Update the configuration selected by the user to the object, the parameter is many attribute/value pairs separated by "|"
Public Sub SetConfig(sAttrs As String)
    Dim sa() As String, I As Long
    
    sa = Split(sAttrs, "|")
    Debug.Assert (UBound(sa) Mod 1 = 0)
    
    m_dic.RemoveAll
    
    For I = 0 To UBound(sa) - 1 Step 2
        m_dic(sa(I)) = sa(I + 1)
    Next
    
End Sub

'Modify or add a single configuration item, attribute/value separated by "|"
Public Sub SetSingleConfig(sAttr As String)
    Dim sa() As String
    sa = Split(sAttr, "|")
    Debug.Assert (UBound(sa) = 1)
    m_dic(sa(0)) = sa(1)
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 1
    Select Case sAttr
        Case "anchor"
            sa = Split("'w','n','s','e','nw','ne','sw','se','center'", ",")
        Case "relief", "overrelief"
            sa = Split("FLAT,GROOVE,RAISED,RIDGE,SOLID,SUNKEN", ",")
        Case "takefocus"
            sa = Split("1,0", ",")
        Case "state"
            sa = Split("'normal','disabled'", ",")
        Case "justify"
            sa = Split("'left','right','center'", ",")
        Case "orient"
            sa = Split("'horizontal','vertical'", ",")
        Case "cursor"
            sa = Split("'arrow','bottom_left_corner','bottom_right_corner','center_ptr','circle','clock','cross'," & _
                "'crosshair','dot','double_arrow','exchange','fleur','hand1','hand2','icon','left_ptr','plus'," & _
                "'question_arrow','sb_h_double_arrow','sb_v_double_arrow','sizing','tcross','watch','xterm','X_cursor'", ",")
            GetAttrValueList = 2
        Case Else
            GetAttrValueList = 0
    End Select
            
End Function

'Return attribute online help
Public Function Tips(sAttr As String) As String
    
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "text", "label":
        Tips = Tips & L("l_TipText", "Text displayed on the widget.")
    Case "x", "y":
        Tips = Tips & L("l_TipXY", "Position of widget.")
    Case "width":
        Tips = Tips & L("l_TipWidth", "Width of widget.")
    Case "height":
        Tips = Tips & L("l_TipHeight", "Height of widget.")
    Case "relx", "rely":
        Tips = Tips & L("l_TipRelXY", "Relative position of widget. value between 0 and 1.")
    Case "relwidth":
        Tips = Tips & L("l_TipRelWidth", "Relative width of widget. value between 0 and 1.")
    Case "relheight":
        Tips = Tips & L("l_TipRelHeight", "Relative height of widget. value between 0 and 1.")
    Case "fg":
        Tips = Tips & L("l_TipFg", "Normal foreground (text) color. format is #RRGGBB, for example : #FF0000.")
    Case "bg":
        Tips = Tips & L("l_TipBg", "Normal background color. format is #RRGGBB, for example : #FF0000.")
    Case "bd":
        Tips = Tips & L("l_TipBd", "Width of the border around the outside of widget.")
    Case "anchor":
        Tips = Tips & L("l_TipAnchor", "Controls where the text is positioned.\nThey are: 'w'|'n'|'s'|'e'|'nw'|'ne'|'sw'|'se'|'center'.")
    Case "relief":
        Tips = Tips & L("l_TipRelief", "Refers to certain simulated 3-D effects around the outside of widget.\nThey are: FLAT, GROOVE, RAISED, RIDGE, SOLID, SUNKEN.")
    Case "overrelief":
        Tips = Tips & L("l_TipOverRelief", "The relief style to be used while the mouse is on the widget.\nThey are: FLAT, GROOVE, RAISED, RIDGE, SOLID, SUNKEN.")
    Case "takefocus":
        Tips = Tips & L("l_TipTakeFocus", "Normally, keyboard focus does visit widget.\nSet to zero to prevent focus from visiting the widget.")
    Case "state":
        Tips = Tips & L("l_TipState", "State of the widget. They are 'normal', 'disabled'")
    Case "underline":
        Tips = Tips & L("l_TipUnderline", "If nonnegative, the corresponding text character will be underlined, index of first char is 0.")
    Case "justify":
        Tips = Tips & L("l_TipJustify", "How the (Multiline) text is justified: 'left','right','center'")
    Case "padding":
        Tips = Tips & L("l_TipPadding", "Specifies the amount of extra space to add around the outside of the widget.")
    Case "orient":
        Tips = Tips & L("l_TipOrient", "One of 'horizontal' or 'vertical'. Specifies the orientation of the widget.")
    Case "cursor":
        Tips = Tips & L("l_TipCursor", "Cursor of widget. Choose a cursor embedded in tk, or add a prefix '@' in a filename to use a custom cursor, for example '@custom.cur'.")
    Case "variable":
        Tips = Tips & L("l_TipVariable", "The control variable that tracks the current state/value of the widget.")
    Case "textvariable":
        Tips = Tips & L("l_TipTextVariable", "The control variable that tracks the current text displayed of the widget.")
    Case "command":
        Tips = Tips & L("l_TipCommand", "A procedure to be called when the widget is activated or clicked. Can be set to 'top.destroy' and other methods of tk or lambda function too.")
    Case "bindcommand":
        Tips = Tips & L("l_TipBindCommand", "Used to attach events binding to a widget. For example: <Control-C>, <F8>, <Alt-A>.")
    Case "font":
        Tips = Tips & L("l_TipFont", "Font to be used for the widget. It is a tuple, the first two elements are name and size, the following are some of 'bold', 'italic', 'underline', 'overstrike'.")
    Case "xscrollcommand":
        Tips = Tips & L("l_TipXScrlCmd", "If widget is scrollable, this attribute should be the .set() method of the horizontal scrollbar.")
    Case "yscrollcommand":
        Tips = Tips & L("l_TipYScrlCmd", "If widget is scrollable, this attribute should be the .set() method of the vertical scrollbar.")
    Case "xscrollincrement":
        Tips = Tips & L("l_TipXScrlIncre", "Default is 0, if set to a positive value, the widget can be positioned only on multiples of that distance.")
    Case "yscrollincrement":
        Tips = Tips & L("l_TipYScrlIncre", "Default is 0, if set to a positive value, the widget can be positioned only on multiples of that distance.")
    Case "scrollregion"
        Tips = Tips & L("l_TipScrlregion", "A tuple (w, n, e, s) that defines over how large an area the canvas can be scrolled.")
    Case "confine"
        Tips = Tips & L("l_TipConfine", "If true (the default), the canvas cannot be scrolled outside of the scrollregion.")
    Case "highlightbackground"
        Tips = Tips & L("l_TipHlbg", "Color of the focus highlight when the widget does not have focus.")
    Case "highlightcolor"
        Tips = Tips & L("l_TipHlColor", "Color shown in the focus highlight.")
    Case "highlightthickness"
        Tips = Tips & L("l_TipHlThickness", "Thickness of the focus highlight. The default value is 1.")
    Case "selectbackground"
        Tips = Tips & L("l_TipSltbg", "The background color to use when displaying selected items.")
    Case "selectborderwidth"
        Tips = Tips & L("l_TipSltbd", "The width of the border to use around selected items.")
    Case "selectforeground"
        Tips = Tips & L("l_TipSltfg", "The foreground color to use when displaying selected items.")
    Case "tooltip"
        Tips = Tips & L("l_TipTooltip", "The tooltip text of the widget.")
    Case Else:
        Tips = Tips & L("l_TipUnknown", "Unknown Attribute")
    End Select
End Function

'Set the corresponding tkinter control class name
Public Property Let ctlType(sType As String)
    m_Type = sType
End Property

'Get the corresponding tkinter control class name
Public Property Get ctlType() As String
    ctlType = m_Type
End Property

'Set the parent window of the control, default is top
Public Property Let Parent(s As String)
    m_Parent = s
End Property

'Return the parent window of the control, default is top
Public Property Get Parent() As String
    Parent = m_Parent
End Property

'The name of the control class represented by the class instance
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(s As String)
    m_Name = s
End Property

Public Property Let Value(s As String)
    m_Value = s
End Property

Public Property Let StyleName(s As String)
    m_StyleName = s
End Property

Public Property Let ScaleMode(nV As Long)
    m_ScaleMode = nV
End Property
Public Property Get ScaleMode() As Long
    ScaleMode = m_ScaleMode
End Property

' Object serialization function
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_dic
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_dic
End Function

Public Function Keys() As Collection
    Set Keys = New Collection
    Dim k As Variant
    For Each k In m_dic.Keys
        Keys.Add k
    Next
End Function

Private Sub Class_Initialize()
    m_Type = ""
    m_Name = ""
    m_Value = ""
    m_StyleName = ""
    m_Parent = WTOP
    m_ScaleMode = vbTwips
    Set m_dic = New Dictionary
End Sub

'Construct a function skeleton; FuncName: function name, Params: parameters, sbody: function body
Public Function CreateFuncDef(funcName As String, Optional sparams As String = "", Optional sbody As String = "") As String
    Dim txtMod As String, widgetName As String
    If Len(funcName) = 0 Then Exit Function
    CreateFuncDef = "def " & funcName & "(" & sparams & "):" & vbCrLf
    If Len(sbody) Then
        CreateFuncDef = CreateFuncDef & sbody & vbCrLf
    Else
    #If DebugVer Then
        CreateFuncDef = CreateFuncDef & Space(4) & "print('" & funcName & "')" & vbCrLf
    #Else
        CreateFuncDef = CreateFuncDef & Space(4) & "#TODO, Please finish the function here!" & vbCrLf
        txtMod = "_Modified"
        If m_Type = "Text" And Right(funcName, Len(txtMod)) = txtMod Then
            widgetName = Left(funcName, Len(funcName) - Len(txtMod))
            CreateFuncDef = CreateFuncDef & Space(4) & widgetName & ".edit_modified(False)" & vbCrLf
        Else
            CreateFuncDef = CreateFuncDef & Space(4) & "pass" & vbCrLf
        End If
    #End If
    End If
End Function

'Construct a function skeleton (object-oriented code); FuncName: function name, Params: parameters, sbody: function body
Public Function CreateFuncDefOOP(funcName As String, Optional sparams As String = "", Optional sbody As String = "") As String
    Dim txtMod As String, widgetName As String
    If Len(funcName) = 0 Then Exit Function
    CreateFuncDefOOP = "    def " & funcName & "(self" & IIf(Len(sparams), ", ", "") & sparams & "):" & vbCrLf
    If Len(sbody) Then
        CreateFuncDefOOP = CreateFuncDefOOP & sbody & vbCrLf
    Else
    #If DebugVer Then
        CreateFuncDefOOP = CreateFuncDefOOP & "        print('" & funcName & "')" & vbCrLf
    #Else
        CreateFuncDefOOP = CreateFuncDefOOP & "        #TODO, Please finish the function here!" & vbCrLf
        txtMod = "_Modified"
        If m_Type = "Text" And Right(funcName, Len(txtMod)) = txtMod Then
            widgetName = Left(funcName, Len(funcName) - Len(txtMod))
            CreateFuncDefOOP = CreateFuncDefOOP & "        self." & widgetName & ".edit_modified(False)" & vbCrLf
        Else
            CreateFuncDefOOP = CreateFuncDefOOP & "        pass" & vbCrLf
        End If
    #End If
    End If
End Function

'Convert various units to pixels
Public Function toPixelX(nX As Long) As Long
    If m_ScaleMode = vbTwips Then
        toPixelX = Twip2PixelX(nX)
    ElseIf m_ScaleMode = vbPoints Then
        toPixelX = Point2PixelX(nX)
    Else
        toPixelX = nX
    End If
End Function

Public Function toPixelY(nY As Long) As Long
    If m_ScaleMode = vbTwips Then
        toPixelY = Twip2PixelY(nY)
    ElseIf m_ScaleMode = vbPoints Then
        toPixelY = Point2PixelY(nY)
    Else
        toPixelY = nY
    End If
End Function

'Translate VB mouse pointer values to tkinter mouse pointer names
Public Function GetCursorName(nCursor As Long) As String
    Select Case nCursor
        Case vbArrow: GetCursorName = "'arrow'"
        Case vbCrosshair: GetCursorName = "'cross'"
        Case vbIbeam: GetCursorName = "'xterm'"
        Case vbSizePointer: GetCursorName = "'fleur'"
        Case vbSizeNESW: GetCursorName = "'bottom_left_corner'"
        Case vbSizeNS: GetCursorName = "'sb_v_double_arrow'"
        Case vbSizeNWSE: GetCursorName = "'bottom_right_corner'"
        Case vbSizeWE: GetCursorName = "'sb_h_double_arrow'"
        Case vbUpArrow: GetCursorName = "'center_ptr'"
        Case vbHourglass: GetCursorName = "'clock'"
        Case vbNoDrop: GetCursorName = "'X_cursor'"
        Case vbArrowHourglass: GetCursorName = "'watch'"
        Case vbArrowQuestion: GetCursorName = "'question_arrow'"
        Case vbSizeAll: GetCursorName = "'sizing'"
        Case Else: GetCursorName = ""
    End Select
End Function

'For latin, decimal 123,45 changed to 123.45
Public Function Commas2Points(ByVal InS As String) As String
    Commas2Points = Replace(InS, ",", ".")
End Function

'Set the reference to the VB control
Public Sub SetVbWidgetInstance(ByRef o As Object)
    Set m_vbWidgetInstance = o
End Sub
