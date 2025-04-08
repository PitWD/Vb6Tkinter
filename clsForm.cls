VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Abstract of form class, storing VB form properties
Private m_Base As clsBaseControl
Private m_Name As String
Private m_Title As String
Private m_Geometry As String
Private m_WindowState As String
Private m_left As Long
Private m_Top As Long
Private m_Topmost As Boolean
Private m_Alpha As String
Private m_Resizable As Boolean
Private m_IconFile As String
Private m_IconEmbedding As Boolean
Private m_Protocol As String
Private m_Bindcommand As String
Private m_FrmPath As String          'Store the directory where the form is located
Private m_ScaleMode As Long
Private m_IconFilesInPath As String 'All icon files in the form's directory, separated by commas
Private m_StartupPosition As String

Const defaultIconName As String = "icon.gif" 'If this file name exists, it is automatically used as the form icon

'Output PYTHON code,
'sOut: Output parameter, interface code
'sCmd: Output parameter, event handler callback code
'sI18n: Output parameter, control text translation code
'rel: Whether to use relative coordinates,
'usettk: Whether to use TTK theme extension
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
     'If the form needs to be centered on the screen, calculate the relative position, this statement should be placed at the front to avoid a small box then moving to the center
    If m_StartupPosition = "Center" Then
        CenterWindow sOut
    End If
    
    sOut.Append "        self.master.title(" & U(m_Title) & ")"
    sI18n.Append "        self.master.title(_(" & U(m_Title) & "))"
    If m_StartupPosition = "Manual" Then  'Manually specify the initial position of the window
        sOut.Append "        self.master.geometry('" & m_Geometry & CoordinateOnStart() & "')"
    ElseIf m_StartupPosition <> "Center" Then
        sOut.Append "        self.master.geometry('" & m_Geometry & "')"
    End If
    If Not m_Resizable Then sOut.Append "        self.master.resizable(0, 0)"
    If Len(m_WindowState) > 0 And m_WindowState <> "'normal'" Then
        sOut.Append "        self.master.wm_state(" & m_WindowState & ")"
    End If
    If m_Topmost Then
        'Adding wait function because of http://code.activestate.com/lists/python-list/385778/, some users reported that the form occasionally does not display,
        'the wait function can avoid this problem
        sOut.Append "        self.master.wait_visibility(self.master)"
        sOut.Append "        self.master.wm_attributes('-topmost', 1)"
    End If
    If Len(m_Alpha) Then
        sOut.Append "        self.master.wm_attributes('-alpha', " & m_Alpha & ")"
    End If
    
    ProcessIcon sOut  'Process form icon
    ProcessBindCommand sOut, sCmd ' Process bindcommand
    ProcessProtocol sOut, sCmd 'Window message interception, if any
End Sub

'Process form icon
Private Sub ProcessIcon(sOut As cStrBuilder)
    
    Dim s As String, sBase64 As String
    
    If Len(m_IconFile) = 0 Then Exit Sub
    
    s = UCase(FileExt(m_IconFile))
    
    'Process form icon
    If s = "GIF" And m_IconEmbedding Then
        sBase64 = CreateIconBase64(Space(12)) 'Embed GIF image in source code using BASE64
        If Len(sBase64) Then
            sOut.Append "        self.icondata = " & String(3, Chr(34)) & vbCrLf & sBase64 & String(3, Chr(34))
            sOut.Append "        self.iconimg = PhotoImage(data=self.icondata)"
            sOut.Append "        self.master.iconphoto(True, self.iconimg)"
        End If
    ElseIf Len(s) = 0 Or s = "ICO" Then 'If there is no extension, assume it is ico format
        sOut.Append "        self.master.iconbitmap(default=r'" & FileFullName(m_IconFile) & "')"
    Else
        sOut.Append "        self.iconimg = PhotoImage(file=r'" & FileFullName(m_IconFile) & "')"
        sOut.Append "        self.master.iconphoto(True, self.iconimg)"
    End If
End Sub

' Process bindcommand
Private Sub ProcessBindCommand(sOut As cStrBuilder, sCmd As cStrBuilder)
    
    Dim I As Long, s() As String, sCmdName As String
    
    If Len(m_Bindcommand) = 0 Then Exit Sub
    
    s = Split(m_Bindcommand, ",")
    
    For I = 0 To UBound(s)
        s(I) = Trim(s(I))
        If Left$(s(I), 1) = "<" And Right$(s(I), 1) = ">" Then
            
            sCmdName = m_Name & "_" & Replace(Replace(Replace(s(I), "<", ""), ">", ""), "-", "_")
            
            sOut.Append "        self.master.bind('" & s(I) & "', self." & sCmdName & ")"
            sCmd.Append m_Base.CreateFuncDefOOP(sCmdName, "event")
            
            'Python is case sensitive, and so are the corresponding shortcut keys, if the shortcut key contains a letter key, bind the corresponding uppercase/lowercase as well
            If Right$(s(I), 3) >= "-a>" And Right$(s(I), 3) <= "-z>" Then
                s(I) = Left(s(I), Len(s(I)) - 2) & UCase(Mid(s(I), Len(s(I)) - 1, 1)) & ">" 'Change to uppercase
                sOut.Append "        self.master.bind('" & s(I) & "', self." & sCmdName & ")"
            ElseIf Right$(s(I), 3) >= "-A>" And Right$(s(I), 3) <= "-Z>" Then
                s(I) = Left(s(I), Len(s(I)) - 2) & LCase(Mid(s(I), Len(s(I)) - 1, 1)) & ">" 'Change to lowercase
                sOut.Append "        self.master.bind('" & s(I) & "', self." & sCmdName & ")"
            End If
        End If
    Next
    
End Sub

'Window message interception, if any
Private Sub ProcessProtocol(sOut As cStrBuilder, sCmd As cStrBuilder)
    
    Dim sa() As String, I As Long, s As String
    
    If Len(m_Protocol) = 0 Then Exit Sub
    
    sa = Split(m_Protocol, ",")
    
    For I = 0 To UBound(sa)
        sOut.Append "        self.master.protocol('" & sa(I) & "', self.EV_" & sa(I) & ")"
        If sa(I) = "WM_DELETE_WINDOW" Then
            s = "        if askyesno(message='Are you sure to quit?'):" & vbCrLf & Space(12) & "self.master.destroy()"
            sCmd.Append m_Base.CreateFuncDefOOP("EV_" & sa(I), "event=None", s)
        Else
            sCmd.Append m_Base.CreateFuncDefOOP("EV_" & sa(I), "event=None")
        End If
    Next
End Sub

'Place the main form in the center of the screen
'In fact, these statements can be replaced with one statement: self.master.eval('tk::PlaceWindow . center')
Private Sub CenterWindow(sOut As cStrBuilder)
    Dim W As String, H As String
    
    W = Left$(m_Geometry, InStr(1, m_Geometry, "x") - 1)
    H = Mid$(m_Geometry, InStr(1, m_Geometry, "x") + 1)
    sOut.Append "        x = int((self.master.winfo_screenwidth() - " & W & ") / 2)"
    sOut.Append "        y = int((self.master.winfo_screenheight() - " & H & ") / 2)"
    sOut.Append "        self.master.geometry('" & W & "x" & H & "+{}+{}'.format(" & "x, y))"
End Sub

Private Function CreateIconBase64(sPrexSpace As String) As String

    Dim sIconFile As String, abContent() As Byte
    
    CreateIconBase64 = ""
    If Len(m_IconFile) <= 0 Then Exit Function
    
    sIconFile = IIf(InStr(1, m_IconFile, "\") Or InStr(1, m_IconFile, "/"), "", m_FrmPath) & m_IconFile
    
    On Error GoTo DirErr
    If Dir(sIconFile) = "" Then
        MsgBox L_F("l_msgFileNotExist", "File '{0}' not exist!", sIconFile), vbInformation
        Exit Function
    End If
    
    If ReadFileBinaryContent(sIconFile, abContent) = 0 Then
        MsgBox L_F("l_msgReadFileError", "Error in Reading File {0}.", sIconFile), vbInformation
        Exit Function
    End If
    
    Base64Encode abContent, CreateIconBase64, sPrexSpace
    Exit Function
DirErr:
    MsgBox L_F("l_msgFileNotExist", "File '{0}' not exist!", sIconFile), vbInformation
End Function


'Call this function immediately after creating the object to initialize the parameters, the form size is in pixels
Public Sub InitConfig(o As Object, nWidth As Long, nHeight As Long, dMethods As Dictionary)
    
    Dim s As String, cAttr As String, cValue As String, sa() As String, I As Long, idx As Long
    
    m_Base.SetVbWidgetInstance o
    
    'Get the directory where the form is located
    On Error Resume Next
    s = o.FileNames(1)
    If Err.Number = 0 And Len(s) Then
        m_FrmPath = PathName(s)
    Else
        m_FrmPath = ""
    End If
    Err.Clear
    On Error GoTo 0
    
    AutoDetectIconFile
    
    m_Name = o.Properties("Name")
    m_Title = o.Properties("Caption")
    m_Geometry = nWidth & "x" & nHeight
    m_WindowState = IIf(o.Properties("WindowState") = vbMaximized, "'zoomed'", IIf(o.Properties("WindowState") = vbMinimized, "'iconic'", "'normal'"))
    m_Resizable = IIf(o.Properties("MaxButton"), True, False)
    m_left = m_Base.toPixelX(o.Properties("Left"))
    m_Top = m_Base.toPixelY(o.Properties("Top"))
    m_IconEmbedding = True
    m_Protocol = ""
    If InStr(1, m_IconFilesInPath, ",") <= 0 Then
        m_IconFile = m_IconFilesInPath
    Else
        m_IconFile = ""
        
        'If there is an image file named icon.gif, automatically select this file
        sa = Split(m_IconFilesInPath, ",")
        For I = 0 To UBound(sa)
            s = sa(I)
            If (Len(s) > Len(defaultIconName)) And (Right(LCase(s), Len(defaultIconName)) = defaultIconName) Then
                m_IconFile = s
                Exit For
            End If
        Next
    End If
    m_Bindcommand = ""
    m_Topmost = False
    m_Alpha = ""
    If o.Properties("StartUpPosition") = 1 Or o.Properties("StartUpPosition") = 2 Then
        m_StartupPosition = "Center"
    ElseIf o.Properties("StartUpPosition") = 0 Then
        m_StartupPosition = "Manual"
    Else
        m_StartupPosition = "Default"
    End If
    
    
    'Use the control's tag property to save some additional default properties,
    'Start character: p@, followed by (property=value) pairs, each property is separated by @, property values are mandatory for forms
    s = Trim(o.Properties("Tag"))
    If Len(s) Then
        sa = Split(s, "@")
        If UCase(sa(0)) = "P" Then
            For I = 1 To UBound(sa)
                s = sa(I)
                idx = InStr(2, s, "=")
                If idx > 0 Then 'Has value
                    cAttr = Left$(s, idx - 1)
                    cValue = Mid$(s, idx + 1)
                    Select Case cAttr
                    Case "topmost"
                        m_Topmost = IIf((cValue = "0" Or UCase(cValue) = "FALSE"), False, True)
                    Case "alpha"
                        m_Alpha = cValue
                    Case "icon"
                        m_IconFile = cValue
                    Case "iconembedding"
                        m_IconEmbedding = IIf((cValue = "0" Or UCase(cValue) = "FALSE"), False, True)
                    Case "bindcommand"
                        m_Bindcommand = cValue
                    Case "protocol"
                        m_Protocol = cValue
                    End Select
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
    If sAttr = "resizable" Or sAttr = "iconembedding" Or sAttr = "topmost" Then
        sa = Split("1,0", ",")
    ElseIf sAttr = "windowstate" Then
        sa = Split("'normal','iconic','zoomed'", ",")
    ElseIf sAttr = "icon" And Len(m_IconFilesInPath) > 0 Then
        sa = Split(m_IconFilesInPath, ",")
        GetAttrValueList = 2
    ElseIf sAttr = "protocol" Then
        sa = Split("WM_DELETE_WINDOW", ",")
        GetAttrValueList = 2
    ElseIf sAttr = "alpha" Then
        sa = Split("0.0,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0", ",")
        GetAttrValueList = 2
    ElseIf sAttr = "position" Then
        sa = Split("Default,Center,Manual", ",")
    Else
        GetAttrValueList = 0
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
    Tips = sAttr & vbCrLf
    Select Case sAttr:
        Case "title"
            Tips = Tips & L("l_TipFrmTitle", "Title of Form.")
        Case "geometry"
            Tips = Tips & L("l_TipFrmGeometry", "Geometry of Form(width x height), unit is pixel.")
        Case "left"
            Tips = Tips & L("l_TipFrmLeft", "Coordinate of Form on startup, unit is pixel.")
        Case "top"
            Tips = Tips & L("l_TipFrmTop", "Coordinate of Form on startup, unit is pixel.")
        Case "resizable"
            Tips = Tips & L("l_TipFrmResizable", "Form is resizable or not, They are 1/0.")
        Case "windowstate"
            Tips = Tips & L("l_TipFrmState", "window state of Form on startup. 'iconic':minimized, 'zoomed':maximized, 'normal': normal.")
        Case "topmost"
            Tips = Tips & L("l_TipFrmTopmost", "Specifies whether this is a topmost window (displays above all other windows).")
        Case "alpha"
            Tips = Tips & L("l_TipFrmAlpha", "Specifies the alpha transparency level of the toplevel. It accepts a value from 0.0 (fully transparent) to 1.0 (opaque).")
        Case "icon"
            Tips = Tips & L("l_TipFrmIcon", "Icon file of Form, supports ico,gif,ppm,pgm.")
        Case "iconembedding"
            Tips = Tips & L("l_TipFrmIconEmbedding", "Is icon embedding in source code or not, supports gif only.")
        Case "bindcommand"
            Tips = Tips & L("l_TipBindCommand", "Used to attach events binding to a widget. for example:<Control-C>,<F8>,<Alt-A>.")
        Case "protocol"
            Tips = Tips & L("l_TipFrmProtocol", "Process window events, for example 'WM_DELETE_WINDOW', separate events using comma like 'WM_DELETE_WINDOW,WM_SAVE'.")
        Case "position"
            Tips = Tips & L("l_TipFrmPosition", "The Startup position of Window.")
        Case Else:
            Tips = Tips & L("l_TipUnknown", "Unknown Attribute")
    End Select
End Function

'Store all icon files in the form directory into m_IconFilesInPath for later selection
Private Function AutoDetectIconFile() As String
    Dim sOut As New cStrBuilder, s As String
    m_IconFilesInPath = ""
    If Len(m_FrmPath) Then
        s = Join(SearchFiles(m_FrmPath, "*.ico", True), ",")
        If Len(s) Then sOut.Append s
        s = Join(SearchFiles(m_FrmPath, "*.gif", True), ",")
        If Len(s) Then sOut.Append s
        s = Join(SearchFiles(m_FrmPath, "*.ppm", True), ",")
        If Len(s) Then sOut.Append s
        s = Join(SearchFiles(m_FrmPath, "*.pgm", True), ",")
        If Len(s) Then sOut.Append s
        m_IconFilesInPath = sOut.toString(",")
    End If
End Function

'Calculate the coordinates when starting
Private Function CoordinateOnStart() As String
    CoordinateOnStart = IIf(m_left >= 0, "+" & CStr(m_left), CStr(m_left))
    CoordinateOnStart = CoordinateOnStart & IIf(m_Top >= 0, "+" & CStr(m_Top), CStr(m_Top))
End Function

'Return a collection, each item is a triplet "property name|value|default selection"
'This function is used to populate the property parameter list box on the main interface
Public Function Allitems() As Collection
    Dim re As Collection
    
    Set re = New Collection
    re.Add "geometry|" & m_Geometry & "|1"
    re.Add "title|" & m_Title & "|1"
    re.Add "left|" & m_left & "|1"
    re.Add "top|" & m_Top & "|1"
    re.Add "resizable|" & IIf(m_Resizable, 1, 0) & "|1"
    re.Add "windowstate|" & m_WindowState & "|1"
    re.Add IIf(m_Topmost, "topmost|1|1", "topmost||0")
    re.Add IIf(m_Alpha <> "", "alpha|" & m_Alpha & "|1", "alpha||0")
    re.Add "position|" & m_StartupPosition & "|1"
    re.Add "icon|" & m_IconFile & "|1"
    re.Add "iconembedding|" & IIf(m_IconEmbedding, 1, 0) & "|1"
    re.Add IIf(m_Bindcommand <> "", "bindcommand|" & m_Bindcommand & "|1", "bindcommand||0")
    re.Add IIf(m_Protocol <> "", "protocol|" & m_Protocol & "|1", "protocol||0")
    
    Set Allitems = re
End Function

Public Sub SetConfig(sAttrs As String)
    Dim sa() As String, I As Long
    
    m_Title = ""
    m_Resizable = True
    m_WindowState = ""
    m_Topmost = False
    m_Alpha = ""
    m_IconFile = ""
    m_IconEmbedding = False
    m_Bindcommand = ""
    m_Protocol = ""
    m_StartupPosition = ""
    m_left = 0
    m_Top = 0
    
    sa = Split(sAttrs, "|")
    
    For I = 0 To UBound(sa) - 1 Step 2
        SetSingleConfig sa(I) & "|" & sa(I + 1)
    Next
    
End Sub

Public Sub SetSingleConfig(sAttr As String)
    Dim sa() As String
    
    sa = Split(sAttr, "|")
    'Debug.Assert (UBound(sa) Mod 1 = 0)
    
    If UBound(sa) >= 1 Then
        Select Case sa(0)
            Case "title"
                m_Title = UnQuote(Trim(sa(1)))
            Case "geometry"
                m_Geometry = UnQuote(Trim(sa(1)))
            Case "left"
                m_left = C2Lng(Trim(sa(1)))
            Case "top"
                m_Top = C2Lng(Trim(sa(1)))
            Case "resizable"
                m_Resizable = IIf(Trim(sa(1)) = "1", True, False)
            Case "windowstate"
                m_WindowState = Trim(sa(1))
            Case "topmost"
                m_Topmost = IIf(Trim(sa(1)) = "1", True, False)
            Case "alpha"
                m_Alpha = Trim(sa(1))
            Case "icon"
                m_IconFile = UnQuote(Trim(sa(1)))
            Case "iconembedding"
                m_IconEmbedding = IIf(Trim(sa(1)) = "1", True, False)
            Case "bindcommand"
                m_Bindcommand = Trim(Replace(Replace(Replace(sa(1), Chr(34), ""), "'", ""), " ", ""))
            Case "protocol"
                m_Protocol = Trim(Replace(Replace(Replace(sa(1), Chr(34), ""), "'", ""), " ", ""))
            Case "position"
                m_StartupPosition = Trim(sa(1))
        End Select
    End If
End Sub

Private Function C2Lng(sV As String) As Long
    On Error Resume Next
    C2Lng = CLng(sV)
End Function

Public Property Get Parent() As String
    Parent = ""
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(s As String)
    m_Name = s
End Property

'Used to change its default corresponding widget type, note the legality of the attribute list after modifying the widget type
Public Function SetWidgetType(sType As String, sStyleName As String)
    'm_Base.ctlType = sType
    'm_Base.StyleName = sStyleName
End Function

'Determine whether the main processing function can call its toString() to generate code, default is True, setting to False indicates that it is handled by other objects
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = True
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    'm_CanbeOutByMainForm = bEnable
End Property

Public Property Get Description() As String
    Description = L("l_DescForm", "Main Form.")
End Property

Public Property Let ScaleMode(nV As Long)
    m_ScaleMode = nV
End Property

Private Sub Class_Initialize()
    Set m_Base = New clsBaseControl
End Sub

Private Sub Class_Terminate()
    Set m_Base = Nothing
End Sub

'Object serialization function
Public Function Serializer(vSer As clsSerialization)
    'vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    'vSer.Deserializer m_Base
End Function

