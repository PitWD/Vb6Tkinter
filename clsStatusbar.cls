VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatusbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Statusbar class, since both TK and TTK libraries do not have a statusbar control, we implement a statusbar control here

'A simple statusbar python control class definition, as VB does not allow too many continuation characters, it is divided into two parts and then connected in the program
'Usage of this control:
'  statusbar = Statusbar(top)  #Create a statusbar with only one pane
'  statusbar = Statusbar(top, slots=[30,10,]) #Create three panes, the first two with widths of 30, 10 pixels respectively, and the third occupying the remaining space
'  statusbar.pack(side=BOTTOM, fill=X)   #Place at the bottom of the form
'  statusbar.set(0,'Demo mode')  #Set the content of the first pane
Const DefineStatusbar1  As String = "class Statusbar(Frame):" & vbCrLf & _
"    """ & """" & """A Simple Statusbar" & vbCrLf & _
"    Usage:self.status = Statusbar(self.top, panelwidths=(15,5,))" & vbCrLf & _
"          self.status.pack(side=BOTTOM, fill=X)" & vbCrLf & _
"          self.status.set(0,'Demo mode')" & vbCrLf & _
"          self.status.text('Demo mode')" & vbCrLf & _
"    """ & """" & """" & vbCrLf & _
"    def __init__(self, master, **kw):" & vbCrLf & _
"        """ & """" & """Options:" & vbCrLf & _
"        panelwidths - a tuple of width of panels, actual number of panels is len(panelwidths)+1." & vbCrLf & _
"        """ & """" & """" & vbCrLf

Const DefineStatusbar2 As String = vbCrLf & _
"        panelwidths = kw['panelwidths'] if 'panelwidths' in kw else []" & vbCrLf & _
"        self.lbls = []" & vbCrLf & _
"        for pnlwidth in panelwidths:" & vbCrLf & _
"            lbl = Label(self, width=pnlwidth, anchor=W, relief=SUNKEN)" & vbCrLf & _
"            self.lbls.append(lbl)" & vbCrLf & _
"            lbl.pack(side=LEFT, fill=Y)" & vbCrLf & _
"        lbl = Label(self, anchor=W, relief=SUNKEN)" & vbCrLf & _
"        self.lbls.append(lbl)" & vbCrLf & _
"        lbl.pack(fill=BOTH, expand=1)" & vbCrLf & vbCrLf

Const DefineStatusbar3 As String = "    def set(self, panel, format, *args):" & vbCrLf & _
"        if panel >= len(self.lbls): raise IndexError" & vbCrLf & _
"        self.lbls[panel]['text'] = format % args" & vbCrLf & _
"        self.lbls[panel].update_idletasks()" & vbCrLf & vbCrLf & _
"    text = lambda self,format,*args : self.set(0,format,*args)" & vbCrLf & vbCrLf & _
"    def panelwidth(self, panel, width=None):" & vbCrLf & _
"        if panel >= len(self.lbls): raise IndexError" & vbCrLf & _
"        if width is None:" & vbCrLf & _
"            panelwidth = self.lbls[panel]['width']" & vbCrLf & _
"        else:" & vbCrLf & _
"            self.lbls[panel]['width'] = width" & vbCrLf & vbCrLf & _
"    def clear(self):" & vbCrLf & _
"        for panel in self.lbls:" & vbCrLf & _
"            panel.config(text='')" & vbCrLf & _
"            panel.update_idletasks()" & vbCrLf & vbCrLf

Private m_Base As clsBaseControl
Private m_CanbeOutByMainForm As Boolean

'Output the definition of the statusbar control
'OutOnlyV3: True-Output only Python3 code, False-Output code compatible with Python2/3
Public Property Get WidgetCode(OutOnlyV3 As Boolean) As String
    Dim sInit As String
    If OutOnlyV3 Then
        sInit = "        super().__init__(master)"
    Else
        sInit = "        Frame.__init__(self, master)"
    End If
    WidgetCode = DefineStatusbar1 & sInit & DefineStatusbar2 & DefineStatusbar3
End Property

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handling callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    
    sOut.Append "        self." & m_Base.Name & " = Statusbar(self." & m_Base.Parent & _
        IIf(Len(m_Base("panelwidths")), ", panelwidths=" & m_Base("panelwidths"), "") & ")"
    
    If Len(m_Base("tooltip")) Then
        sOut.Append "        self." & m_Base.Name & "Tooltip = Tooltip(self." & m_Base.Name & ", " & Quote(m_Base("tooltip")) & ")"
    End If
    
    If Len(m_Base("side")) Then
        sOut.Append "        self." & m_Base.Name & ".pack(side=" & m_Base("side") & ", fill=X)"
    ElseIf rel Then
        sOut.Append "        self." & m_Base.Name & ".place(relx=" & m_Base("relx") & ", rely=" & _
            m_Base("rely") & ", relwidth=" & m_Base("relwidth") & ", relheight=" & m_Base("relheight") & ")"
    Else
        sOut.Append "        self." & m_Base.Name & ".place(x=" & m_Base("x") & ", y=" & _
            m_Base("y") & ", width=" & m_Base("width") & ", height=" & m_Base("height") & ")"
    End If
End Sub

'Call this function immediately after creating the object to initialize the parameters
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    
    Dim I As Long, nCnt As Long, s As String, Panels As IPanels
    Dim sa() As String, idx As Long
    
    m_Base.SetVbWidgetInstance o
    
    m_Base.Name = o.Properties("Name")
    
    m_Base("x") = m_Base.toPixelX(o.Properties("Left"))
    m_Base("y") = m_Base.toPixelY(o.Properties("Top"))
    m_Base("width") = m_Base.toPixelX(o.Properties("Width"))
    m_Base("height") = m_Base.toPixelY(o.Properties("Height"))
    m_Base("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_Base("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_Base("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_Base("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    
    If o.Properties("Align") = vbAlignTop Then
        m_Base("side") = "TOP"
    ElseIf o.Properties("Align") = vbAlignBottom Then
        m_Base("side") = "BOTTOM"
    End If
    
    If o.Properties("ToolTipText") <> "" Then
        m_Base("tooltip") = o.Properties("ToolTipText")
    End If
    
    m_Base("panelwidths") = ""
    Set Panels = o.Properties("Panels").object
    nCnt = Panels.Count
    If o.Properties("Style") = sbrNormal And nCnt > 0 Then
        For I = 1 To nCnt
            s = s & Round(m_Base.toPixelX(Panels(I).MinWidth) / CharWidth()) & ","
        Next
        
        If Len(s) Then
            m_Base("panelwidths") = "(" & s & ")"
        End If
    End If
    
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
                    m_Base(Left$(s, idx - 1)) = Mid$(s, idx + 1)
                End If
            Next
        End If
    End If
End Sub

'Set the list of possible values for the attribute
'Return value: 0 - no optional values, 1 - a strictly limited list of optional values, 2 - in addition to the provided list of optional values, other values can be manually entered
'Output: sa() array of optional values
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If sAttr = "side" Then
        GetAttrValueList = 2
        sa = Split("TOP,BOTTOM", ",")
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
    
    If sAttr = "side" Then
        Tips = sAttr & vbCrLf & L("l_TipSide", "Position of statusbar. they are TOP,BOTTOM, or blank.")
    ElseIf sAttr = "panelwidths" Then
        Tips = sAttr & vbCrLf & L("l_TipPanelWidths", "A tuple or list of number of chars of panels.")
    Else
        Tips = m_Base.Tips(sAttr)
    End If
    
End Function

'Update the configuration selected by the user to the object, the parameter is many attribute/value pairs separated by "|"
Public Sub SetConfig(sAttrs As String)

    Dim sa() As String, I As Long
    
    sa = Split(sAttrs, "|")
    Debug.Assert (UBound(sa) Mod 1 = 0)
    
    'm_Base.RemoveAll
    
    For I = 0 To UBound(sa) - 1 Step 2
        SetSingleConfig sa(I) & "|" & sa(I + 1)
    Next
    
End Sub

'Modify or add a single configuration item, attribute/value separated by "|"
Public Sub SetSingleConfig(sAttr As String)
    
    Dim sa() As String, s As String
    
    sa = Split(sAttr, "|")
    Debug.Assert (UBound(sa) = 1)
    
    If sa(0) = "panelwidths" Then
        s = Replace(Replace(sa(1), "'", ""), " ", "")
        s = Replace(Replace(s, "[", "("), "]", ")")
        If Left$(s, 1) <> "(" And Right$(s, 1) <> ")" Then
            If InStr(1, s, ",") Then 'Use commas
                m_Base("panelwidths") = "(" & s & ")"
            ElseIf InStr(1, s, " ") Then 'Use spaces
                s = Replace(Replace(s, "  ", " "), " ", ",")
                m_Base("panelwidths") = "(" & s & ")"
            Else
                m_Base("panelwidths") = s
            End If
        Else
            m_Base("panelwidths") = s
        End If
    Else
        m_Base(sa(0)) = sa(1)
    End If
    
End Sub

'Return a collection, each item is a triplet "property name|value|default selection"
'This function is used to populate the property parameter list box on the main interface
Public Function Allitems() As Collection
    Dim re As Collection, k As Variant, ks As Collection
    
    Set re = New Collection
    
    Set ks = m_Base.Keys
    For Each k In ks
        re.Add k & "|" & m_Base(k) & "|1"
    Next
    
    Set Allitems = re
    
End Function

Private Sub Class_Initialize()
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Statusbar"
    m_Base.StyleName = ""
    m_CanbeOutByMainForm = True
End Sub

Private Sub Class_Terminate()
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
    Description = L("l_DescStatusbar", "Statusbar widget. it has method set(), text(), panelwidth(), clear()")
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



