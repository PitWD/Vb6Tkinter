VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Vb6Tkinter https://github.com/cdhigh"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   542
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   865
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cmbEditCombo 
      CausesValidation=   0   'False
      Height          =   345
      ItemData        =   "MainForm.frx":058A
      Left            =   7560
      List            =   "MainForm.frx":058C
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbEditList 
      Height          =   345
      ItemData        =   "MainForm.frx":058E
      Left            =   6360
      List            =   "MainForm.frx":0590
      Style           =   2  'Dropdown-Liste
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Vb6Tkinter.xpcmdbutton CmdRefsFormsList 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Refresh Forms(&R)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar stabar 
      Align           =   2  'Unten ausrichten
      Height          =   377
      Left            =   0
      TabIndex        =   10
      Top             =   7748
      Width           =   12974
      _ExtentX        =   22886
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2687
            MinWidth        =   2687
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbFrms 
      Height          =   345
      ItemData        =   "MainForm.frx":0592
      Left            =   120
      List            =   "MainForm.frx":0594
      Style           =   2  'Dropdown-Liste
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox TxtTips 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   7
      Top             =   4920
      Width           =   2415
   End
   Begin VB.ListBox LstComps 
      Height          =   3210
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin Vb6Tkinter.GridOcx LstCfg 
      Height          =   6855
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12091
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   9
      Top             =   840
      Width           =   4095
   End
   Begin Vb6Tkinter.xpcmdbutton CmdGenCode 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Generate Code(&G)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Vb6Tkinter.xpcmdbutton CmdCopyToClipboard 
      Height          =   495
      Left            =   5340
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Copy to Clipboard(&C)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Vb6Tkinter.xpcmdbutton CmdSaveToFile 
      Height          =   495
      Left            =   7950
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Save to File(&F)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Vb6Tkinter.xpcmdbutton CmdQuit 
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Quit(&Q)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File(&F)"
      Begin VB.Menu mnuRefreshForms 
         Caption         =   "Refresh Forms(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenCode 
         Caption         =   "Generate Code(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSeparator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveToFile 
         Caption         =   "Save Code to File(&F)"
         Begin VB.Menu mnuSaveAll 
            Caption         =   "Save All Code(&A)"
         End
         Begin VB.Menu mnuSaveUiOnly 
            Caption         =   "Save Class UI Only(&G)"
         End
      End
      Begin VB.Menu mnuCopyToClipboard 
         Caption         =   "Copy Code To Clipboard(&C)"
         Begin VB.Menu mnuCopyToClipAll 
            Caption         =   "Copy All Code(&A)"
         End
         Begin VB.Menu mnuCopyToClipUiOnly 
            Caption         =   "Copy Class UI Only(&G)"
         End
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddProperty 
         Caption         =   "Add One Property(&P)"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit(&Q)"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options(&O)"
      Begin VB.Menu mnuV2andV3Code 
         Caption         =   "Compatible Code for Python 2.x/3.x(&C)"
      End
      Begin VB.Menu mnuUseTtk 
         Caption         =   "Use TTK Themed Library(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRelPos 
         Caption         =   "Use Relative Position(&R)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuI18n 
         Caption         =   "Support i18n(&I)"
      End
      Begin VB.Menu mnuUnicodePrefixU 
         Caption         =   "Add A Prefix 'u' to Unicode String(&U)"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPythonExe 
         Caption         =   "Set diretory of python.exe(&E)..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools(&T)"
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview(&P)"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEncodeAFile 
         Caption         =   "Encode File to Base64(&B)"
      End
      Begin VB.Menu mnuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpdate 
         Caption         =   "Check Update(&U)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About(&A)"
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "Language(&L)"
      Begin VB.Menu mnuLng 
         Caption         =   "English(&E)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mConnect As Connect

' 2012.11.23: To facilitate the binding of the scrollbar and the list box, it was modified to a global variable g_Comps, placed in Common.bas
' Private m_Comps() As Object             ' Matches the number of rows in LstComps, corresponding to instances generated by each component
Private m_MainMenu As clsMenu              ' Menu object
Private m_PrevCompIdx As Long
Private m_curFrm As Object
Private m_prevsf As String
Private m_nLngNum As Long                  ' Language type
Private m_HasCommonDialog As Boolean
Private m_saTmpFile() As String
Private m_TxtCodeExpanded As Boolean
Private m_TxtTipsExpanded As Boolean
Private m_BriefCaption As String

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Private Sub Form_Load()
    
    Dim s As String
    ReDim g_Comps(0) As Object
    
    g_AppVerString = App.Major & "." & App.Minor & IIf(App.Revision > 0, "." & App.Revision, "")
    g_DefaultFontName = ""
    m_HasCommonDialog = False
    m_TxtCodeExpanded = False
    m_TxtTipsExpanded = False
    
    ReDim m_saTmpFile(0) As String
    
    'Multi-Language Project
    InitMultiLanguage
    
    LstCfg.Redraw = False
    LstCfg.Editable = True
    LstCfg.EditType = EnterKey Or MouseDblClick Or F2Key
    LstCfg.CheckBoxes = True
    LstCfg.AddColumn "Property", 2260, lgAlignCenterCenter
    LstCfg.AddColumn "Value", 3450, lgAlignCenterCenter
    LstCfg.ColAlignment(0) = lgAlignLeftCenter
    LstCfg.ColAlignment(1) = lgAlignLeftCenter
    LstCfg.SelectBackColor = &HFCC597 'vbHighlight
    LstCfg.Redraw = True
    
    m_BriefCaption = "Vb6Tkinter v" & g_AppVerString
    #If DebugVer Then
        m_BriefCaption = m_BriefCaption & " [Debug Mode] "
    #End If
    Me.Caption = m_BriefCaption
    
    mnuV2andV3Code.Checked = GetSetting(App.Title, "Settings", "V2andV3Code", "0") = "1"
    mnuUseTtk.Checked = GetSetting(App.Title, "Settings", "UseTtk", "1") = "1"
    mnuRelPos.Checked = GetSetting(App.Title, "Settings", "RelPos", "1") = "1"
    mnuI18n.Checked = GetSetting(App.Title, "Settings", "i18n", "1") = "1"
    mnuUnicodePrefixU.Checked = GetSetting(App.Title, "Settings", "UnicodePrefix", "0") = "1"
    g_bUnicodePrefixU = mnuUnicodePrefixU.Checked
    
    g_PythonExe = GetSetting(App.Title, "Settings", "PythonExe", "")
    
    Set cmbEditList.Font = LstCfg.Font
    Set cmbEditCombo.Font = LstCfg.Font
    
    ResizeInit Me
    
    CmdRefsFormsList_Click
    
End Sub

' Multi-language support initialization
Private Sub InitMultiLanguage()
    
    Dim I As Long, s As String, sa() As String
    
    If Not LngFileExist() Then
        m_nLngNum = 0
        mnuLng(0).Checked = True
        Exit Sub
    End If
    
    sa = GetAllLanguageName()
    mnuLng(0).Caption = sa(0)
    m_nLngNum = 1
    For I = 1 To UBound(sa)
        Load mnuLng(I)
        mnuLng(I).Caption = sa(I)
        m_nLngNum = m_nLngNum + 1
    Next
    
    ' Switch language, prioritize the language saved in the registry, followed by the operating system's language
    s = GetSetting(App.Title, "Settings", "Language", "")
    I = m_nLngNum
    If Len(s) Then                                                              ' Select the previously saved language type, if it exists
        For I = 0 To m_nLngNum - 1
            If mnuLng(I).Caption = s Then
                ChangeLanguage (mnuLng(I).Caption)
                mnuLng(I).Checked = True
                Exit For
            End If
        Next
    End If
    
    ' Attempt to determine the operating system's language
    If I > m_nLngNum - 1 Then
        
        I = GetSystemDefaultLCID()
        If I = &H804 Or I = &H4 Or I = &H1004 Then
            s = "Simplified Chinese"
        ElseIf I = &H404 Or I = &HC04 Then
            s = "Traditional Chinese"
        ElseIf I Mod 16 = 9 Then
            s = "English"
        Else                                                                    ' For other languages, default to English. Users can select the appropriate language after the software starts.
            s = "English"
        End If
        
        For I = 0 To m_nLngNum - 1
            If InStr(1, mnuLng(I).Caption, s) > 0 Then
                ChangeLanguage (mnuLng(I).Caption)
                mnuLng(I).Checked = True
                Exit For
            End If
        Next
        
        ' Unable to automatically determine the language, default to the first one
        If I > m_nLngNum - 1 Then
            ChangeLanguage (mnuLng(0).Caption)
            mnuLng(0).Checked = True
        End If
    End If
    
End Sub

Private Sub CmdQuit_Click()
    mConnect.Hide
End Sub

Private Sub cmbFrms_Click()
    
    Dim frm As Object
    
    ' Find the corresponding form reference
    Set m_curFrm = Nothing
    If Len(cmbFrms.Text) Then
        For Each frm In VbeInst.ActiveVBProject.VBComponents
            If frm.Type = vbext_ct_VBForm And frm.Name = cmbFrms.Text Then
                Set m_curFrm = frm
                Exit For
            End If
        Next
    End If
    
    m_PrevCompIdx = -1
    
    ' Add controls to the list
    If Not ResetLstComps(m_curFrm) Then
        LstComps.Clear
        LstCfg.Clear
        TxtTips.Text = L("l_TipHasNoControl", "Has no control on Form, please add at least one control on it.")
        m_PrevCompIdx = -1
    Else
        LstComps.ListIndex = 0
        LstComps_Click
    End If
    
    If LstComps.ListCount > 0 Then
        CmdGenCode.Enabled = True
        CmdCopyToClipboard.Enabled = True
        CmdSaveToFile.Enabled = True
        mnuSaveToFile.Enabled = True
        mnuCopyToClipboard.Enabled = True
        mnuAddProperty.Enabled = True
        mnuGenCode.Enabled = True
    Else
        CmdGenCode.Enabled = False
        CmdCopyToClipboard.Enabled = False
        CmdSaveToFile.Enabled = False
        mnuSaveToFile.Enabled = False
        mnuCopyToClipboard.Enabled = False
        mnuAddProperty.Enabled = False
        mnuGenCode.Enabled = False
    End If
    
End Sub

Private Sub CmdCopyToClipboard_Click()
    Me.PopupMenu mnuCopyToClipboard
End Sub

' Update the lists, create corresponding control class instances, return False if initialization fails, True if successful
Private Function ResetLstComps(frm As Object) As Long
    
    Dim Obj As Object, ObjClsModule As Object, I As Long, s As String, j As Long, idx As Long
    Dim nScaleMode As Long, nScaleWidth As Long, nScaleHeight As Long
    Dim CodeMember As Member, CodeMembers As Members, dMethods As New Dictionary
    Dim ctlsIgnored As String
    
    ResetLstComps = False
    If frm Is Nothing Then Exit Function
    
    LstComps.Clear
    ' Erase g_Comps
    Set m_MainMenu = Nothing
    
    ' Create a form instance as the first item in the list
    ReDim g_Comps(0) As Object
    Set g_Comps(0) = New clsForm
    
    ' Since ScaleX/ScaleY are unique methods of the form class, convert the form size to pixel units here
    nScaleWidth = Round(ScaleX(frm.Properties("ScaleWidth"), frm.Properties("ScaleMode"), vbPixels))
    nScaleHeight = Round(ScaleY(frm.Properties("ScaleHeight"), frm.Properties("ScaleMode"), vbPixels))
    g_Comps(0).InitConfig frm, nScaleWidth, nScaleHeight, dMethods
    g_Comps(0).Name = WTOP
    LstComps.AddItem g_Comps(0).Name & " (Form)"
    I = 1
    
    m_HasCommonDialog = False
    
    ' Get all procedure functions in the form's code module, save them as a dictionary object, and pass them to various modules for automatically generating corresponding bind commands
    If Not frm.CodeModule Is Nothing Then
        Set CodeMembers = frm.CodeModule.Members
        If Not CodeMembers Is Nothing Then
            For Each CodeMember In CodeMembers
                If CodeMember.Type = vbext_mt_Method Then
                    idx = InStrRev(CodeMember.Name, "_")
                    If idx > 1 Then
                        s = Left$(CodeMember.Name, idx - 1)
                        If dMethods.Exists(s) Then
                            dMethods.Item(s) = dMethods.Item(s) & "," & CodeMember.Name & "," ' Use commas as delimiters for easy searching
                        Else
                            dMethods.Item(s) = "," & CodeMember.Name & ","
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    ' Add controls to the list
    For Each Obj In frm.Designer.VBControls
        
        CreateObj Obj, ObjClsModule                                             ' Generate corresponding class module instance
        
        If Not ObjClsModule Is Nothing Then
            
            ' For automatic unit conversion, set this value before InitConfig
            ObjClsModule.ScaleMode = frm.Properties("ScaleMode")
            
            ' If the form has a menu control, create a main menu object. The main menu control will manage all menu items.
            If Obj.ClassName = "Menu" And m_MainMenu Is Nothing Then
                ReDim Preserve g_Comps(I) As Object
                Set m_MainMenu = New clsMenu
                Set g_Comps(I) = m_MainMenu
                LstComps.AddItem m_MainMenu.Name & " (MainMenu)"
                m_MainMenu.InitConfig
                I = I + 1
            End If
            
            ' Add control to the control list
            ReDim Preserve g_Comps(I) As Object
            Set g_Comps(I) = ObjClsModule
            LstComps.AddItem Obj.Properties("Name") & " (" & Obj.ClassName & ")"
            
            ' Initialize the class module object corresponding to each control
            If Obj.Container Is frm.Designer Then
                g_Comps(I).Parent = IIf(Obj.ClassName = "Menu", "MainMenu", WTOP)
                g_Comps(I).InitConfig Obj, frm.Properties("ScaleWidth"), frm.Properties("ScaleHeight"), dMethods
            ElseIf Obj.Container.ClassName = "Menu" Then  ' Submenu
                g_Comps(I).Parent = Obj.Container.Properties("Name")
                g_Comps(I).InitConfig Obj, 0, 0, dMethods
            Else
                On Error Resume Next
                nScaleMode = Obj.Container.Properties("ScaleMode")
                nScaleWidth = Obj.Container.Properties("ScaleWidth")
                nScaleHeight = Obj.Container.Properties("ScaleHeight")
                If Err.Number Then         ' Frame and some other containers do not support the ScaleWidth property, so use Width instead
                    nScaleMode = vbTwips
                    nScaleWidth = Me.ScaleX(Obj.Container.Properties("Width"), frm.Properties("ScaleMode"), vbTwips)
                    nScaleHeight = Me.ScaleY(Obj.Container.Properties("Height"), frm.Properties("ScaleMode"), vbTwips)
                End If
                Err.Clear
                On Error GoTo 0
                g_Comps(I).ScaleMode = nScaleMode
                g_Comps(I).Parent = Obj.Container.Properties("Name")
                g_Comps(I).InitConfig Obj, nScaleWidth, nScaleHeight, dMethods
            End If
            
            I = I + 1
            ResetLstComps = True
        ElseIf Obj.ClassName = "CommonDialog" Then
            m_HasCommonDialog = True
        ElseIf Len(ctlsIgnored) = 0 Or InStr(1, ctlsIgnored, Obj.ClassName & ",") <= 0 Then
            If MsgBox(L_F("l_msgCtlNotSupport", "The addin does not support '{0}' control (Name:{1}).\n\nIt will not be processed.\n\n'Ok' to continue.\n'Cancel' for ignoring controls of the same type.", _
                Obj.ClassName, Obj.Properties("Name")), vbInformation + vbOKCancel, App.Title) = vbCancel Then
                ctlsIgnored = ctlsIgnored & Obj.ClassName & ","
            End If
        End If
    Next  ' frm.Designer.VBControls
    
    ' Generate the hierarchical tree structure of the menu as a basis for generating code
    CreateMenuHiberarchy
    
    ' Organize the parent-child relationships of Notebook and its sub-widget controls
    ArrangeNotebookAndSubWidgets
    
    ' Attempt to automatically bind scrollbars to corresponding controls
    TryAssignScrollbar2Widgets
    
    ' Correctly set the TTK property of ComboboxAdapter
    For I = 1 To UBound(g_Comps)
        If TypeName(g_Comps(I)) = "clsComboboxAdapter" Then g_Comps(I).TTK = mnuUseTtk.Checked
    Next
    
End Function

' Generate a control character instance object: Input ctlobj: control object, clsobj: corresponding string object
Private Function CreateObj(ByRef ctlobj As Object, ByRef clsobj As Object) As Object
    Dim o As Object, sName As String, idx As Long
    
    Select Case ctlobj.ClassName:
        Case "Label"
            Set clsobj = New clsLabel
        Case "CommandButton"
            Set clsobj = New clsButton
        Case "TextBox"
            If ctlobj.Properties("MultiLine") Then Set clsobj = New clsText Else Set clsobj = New clsEntry
        Case "CheckBox"
            Set clsobj = New clsCheckbutton
        Case "OptionButton"
            Set clsobj = New clsRadiobutton
        Case "ComboBox"
            Set clsobj = New clsComboboxAdapter
            clsobj.TTK = mnuUseTtk.Checked
        Case "ListBox"
            Set clsobj = New clsListbox
        Case "HScrollBar", "VScrollBar"
            Set clsobj = New clsScrollbar
        Case "Slider"
            Set clsobj = New clsScale
        Case "Frame"
            ' Determine if it is a page of a TabStrip control
            idx = InStr(2, ctlobj.Properties("Name"), "__Tab") ' Start from 2 to ensure there is at least one character before __Tab
            If idx > 1 And Not m_curFrm Is Nothing Then
                ' Loop to check if there is a suitable TabStrip control
                Set clsobj = Nothing
                sName = Left$(ctlobj.Properties("Name"), idx - 1)
                For Each o In m_curFrm.Designer.VBControls
                    If o.ClassName = "TabStrip" And o.Properties("Name") = sName Then
                        Set clsobj = New clsNotebookTab  ' Use the Tab class instead of Frame
                        Exit For
                    End If
                Next
                If clsobj Is Nothing Then Set clsobj = New clsLabelFrame ' No corresponding TabStrip object
            Else
                Set clsobj = New clsLabelFrame
            End If
        Case "PictureBox"
            ' Determine if it is a page of a TabStrip control
            idx = InStr(2, ctlobj.Properties("Name"), "__Tab") ' Start from 2 to ensure there is at least one character before __Tab
            If idx > 1 And Not m_curFrm Is Nothing Then
                ' Loop to check if there is a suitable TabStrip control
                Set clsobj = Nothing
                sName = Left$(ctlobj.Properties("Name"), idx - 1)
                For Each o In m_curFrm.Designer.VBControls
                    If o.ClassName = "TabStrip" And o.Properties("Name") = sName Then
                        Set clsobj = New clsNotebookTab  ' Use the Tab class instead of PictureBox
                        Exit For
                    End If
                Next
                If clsobj Is Nothing Then Set clsobj = New clsCanvas ' No corresponding TabStrip object
            Else
                Set clsobj = New clsCanvas
            End If
        Case "Menu"
            Set clsobj = New clsMenuItem
        Case "ProgressBar"
            Set clsobj = New clsProgressBar                                         ' Requires TTK to be enabled
            mnuUseTtk.Checked = True
        Case "TreeView"
            Set clsobj = New clsTreeview                                            ' Requires TTK to be enabled
            mnuUseTtk.Checked = True
        Case "TabStrip"
            Set clsobj = New clsNotebook                                            ' Requires TTK to be enabled
            mnuUseTtk.Checked = True
        Case "Line"
            Set clsobj = New clsSeparator
            mnuUseTtk.Checked = True
        Case "StatusBar"
            Set clsobj = New clsStatusbar
        Case Else:
            Set clsobj = Nothing
    End Select
    
    Set CreateObj = clsobj
    
End Function

' Generate the hierarchical tree structure of the menu as a basis for generating code
' Subclasses store the name of the parent class, and parent classes store references to all subclasses
Private Sub CreateMenuHiberarchy()

    Dim I As Long, j As Long
    If Not m_MainMenu Is Nothing Then
        For I = 0 To UBound(g_Comps)
            If TypeName(g_Comps(I)) = "clsMenu" Then
                ' Add all top-level menus as child controls of clsMenu
                For j = 0 To UBound(g_Comps)
                    If TypeName(g_Comps(j)) = "clsMenuItem" And g_Comps(j).Parent = "MainMenu" Then
                        g_Comps(I).AddChild g_Comps(j)
                    End If
                Next
            ElseIf TypeName(g_Comps(I)) = "clsMenuItem" Then
                ' Submenus may also have submenus
                For j = 0 To UBound(g_Comps)
                    If TypeName(g_Comps(j)) = "clsMenuItem" And g_Comps(j).Parent = g_Comps(I).Name Then
                        g_Comps(I).AddChild g_Comps(j)
                    End If
                Next
            End If
        Next
    End If
    
End Sub

' Organize the parent-child relationships of Notebook controls and their internal controls
Private Sub ArrangeNotebookAndSubWidgets()

    Dim I As Long, j As Long, k As Long, L As Long, idx As Long, ctlNum As Long
    Dim sTabName As String, sNbName As String, sTmp As String
    
    If UBound(g_Comps) <= 0 Then  ' 0 is fixed as the top-level form
        Exit Sub
    End If
    
    ctlNum = UBound(g_Comps)
    For I = 1 To ctlNum
        If TypeName(g_Comps(I)) = "clsNotebookTab" Then
            sTabName = g_Comps(I).Name
            idx = InStr(2, sTabName, "__Tab")
            If idx > 1 Then
                sNbName = Left$(sTabName, idx - 1) ' Notebook control name
                For j = 1 To ctlNum
                    If TypeName(g_Comps(j)) = "clsNotebook" And g_Comps(j).Name = sNbName Then
                        ' Get the TAB number
                        sTmp = Right$(sTabName, 1)
                        If sTmp >= "1" And sTmp <= "9" Then ' Supports up to 9 tabs
                            g_Comps(j).AddTab g_Comps(I), CLng(sTmp)  ' Add to Notebook object
                            g_Comps(I).EnableOutByMainForm = False
                            
                            ' All controls within this tab page are managed by clsNotebookTab and will no longer output code from the main window
                            For k = 1 To ctlNum
                                If g_Comps(k).Parent = sTabName Then
                                    g_Comps(k).EnableOutByMainForm = False
                                    g_Comps(I).AddSubWidget g_Comps(k)
                                    
                                    ' In case there are other container controls within the tab page
                                    If TypeName(g_Comps(k)) = "clsCanvas" Or TypeName(g_Comps(k)) = "clsLabelFrame" Then
                                        For L = 1 To ctlNum
                                            If g_Comps(L).Parent = g_Comps(k).Name Then
                                                g_Comps(L).EnableOutByMainForm = False
                                                g_Comps(I).AddSubWidget g_Comps(L)
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                            
                        End If
                    End If
                Next
            End If
        End If
    Next

End Sub

'Perform some analysis, try to automatically bind the scrollbar to the appropriate control, not necessarily successful, and may misjudge, just do your best
Private Sub TryAssignScrollbar2Widgets()
    
    Dim I As Long, ctlNum As Long, Obj As Object, o As Object, oName As String
    Dim vX1 As Long, vY1 As Long, vX2 As Long, vY2 As Long
    Dim oX1 As Long, oY1 As Long, oX2 As Long, oY2 As Long
    Dim thresholdX1 As Long, thresholdY1 As Long
    Dim thresholdX2 As Long, thresholdY2 As Long
    Dim isWidgetScrl As Boolean, Assigned As Boolean
    
    If UBound(g_Comps) <= 0 Or m_curFrm Is Nothing Then  ' 0 is fixed as the top-level form
        Exit Sub
    End If
    
    'Use 20 pixels as the threshold for finding controls in both horizontal and vertical directions
    thresholdX1 = Round(ScaleX(20, vbPixels, m_curFrm.Properties("ScaleMode")))
    thresholdY1 = Round(ScaleY(20, vbPixels, m_curFrm.Properties("ScaleMode")))
    thresholdX2 = Round(ScaleX(5, vbPixels, m_curFrm.Properties("ScaleMode"))) '5 pixels are allowed to overlap between the control and the scrollbar
    thresholdY2 = Round(ScaleY(5, vbPixels, m_curFrm.Properties("ScaleMode")))
    
    ctlNum = UBound(g_Comps)
    For Each Obj In m_curFrm.Designer.VBControls
        If Obj.ClassName = "HScrollBar" Then
            'Horizontal scrollbar, determine if there is a control above it that needs to set the scrollbar
            'vX1, vY1, vX2, vY2 form a rectangle, if the lower left and lower right corners of other controls fall within this rectangle, it is considered that the scrollbar corresponds
            vX1 = Obj.Properties("Left") - thresholdX1
            If vX1 < 0 Then vX1 = 0
            vY1 = Obj.Properties("Top") - thresholdY1
            If vY1 < 0 Then vY1 = 0
            vX2 = Obj.Properties("Left") + Obj.Properties("Width") + thresholdX1
            vY2 = Obj.Properties("Top") + thresholdY2
            Assigned = False
            For Each o In m_curFrm.Designer.VBControls
                If (o.Container Is Obj.Container) And _
                    InStr(1, "PictureBox,ListBox,TreeView,TextBox,", o.ClassName & ",") > 0 Then  'Only these controls may need scrollbars
                    isWidgetScrl = True
                    If o.ClassName = "TextBox" Then 'Only multiline textboxes support scrolling
                        If Not o.Properties("MultiLine") Then
                            isWidgetScrl = False
                        End If
                    End If
                    If isWidgetScrl Then
                        oX1 = o.Properties("Left")
                        oY1 = o.Properties("Top") + o.Properties("Height")
                        oX2 = oX1 + o.Properties("Width")
                        oY2 = oY1
                        
                        'The first line is the lower left corner judgment, the second line is the lower right corner judgment
                        If (oX1 >= vX1 And oX1 <= vX2 And oY1 >= vY1 And oY1 <= vY2) _
                            And (oX2 >= vX1 And oX2 <= vX2 And oY2 >= vY1 And oY2 <= vY2) Then
                            'Set the xscrollcommand property of the control
                            oName = o.Properties("Name")
                            For I = 1 To ctlNum
                                If g_Comps(I).Name = oName Then
                                    g_Comps(I).SetSingleConfig ("xscrollcommand|" & Obj.Properties("Name") & ".set")
                                    Assigned = True
                                    Debug.Print oName & " assigned to " & Obj.Properties("name")
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                If Assigned Then
                    Exit For
                End If
            Next
        ElseIf Obj.ClassName = "VScrollBar" Then
            'Vertical scrollbar, determine if there is a control to its left that needs to set the scrollbar
            'vX1, vY1, vX2, vY2 form a rectangle, if the upper right and lower right corners of other controls fall within this rectangle, it is considered that the scrollbar corresponds
            vX1 = Obj.Properties("Left") - thresholdX1
            If vX1 < 0 Then vX1 = 0
            vY1 = Obj.Properties("Top") - thresholdY1
            If vY1 < 0 Then vY1 = 0
            vX2 = Obj.Properties("Left") + thresholdX2
            vY2 = Obj.Properties("Top") + Obj.Properties("Height") + thresholdY1
            Assigned = False
            For Each o In m_curFrm.Designer.VBControls
                If (o.Container Is Obj.Container) And _
                    InStr(1, "PictureBox,ListBox,TreeView,TextBox,", o.ClassName & ",") > 0 Then 'Only these controls may need scrollbars
                    isWidgetScrl = True
                    If o.ClassName = "TextBox" Then 'Only multiline textboxes support scrolling
                        If Not o.Properties("MultiLine") Then
                            isWidgetScrl = False
                        End If
                    End If
                    If isWidgetScrl Then
                        oX1 = o.Properties("Left") + o.Properties("Width")
                        oY1 = o.Properties("Top")
                        oX2 = oX1
                        oY2 = oY1 + o.Properties("Height")
                        
                        'The first line is the upper right corner judgment, the second line is the lower right corner judgment
                        If (oX1 >= vX1 And oX1 <= vX2 And oY1 >= vY1 And oY1 <= vY2) _
                            And (oX2 >= vX1 And oX2 <= vX2 And oY2 >= vY1 And oY2 <= vY2) Then
                            'Set the yscrollcommand property of the control
                            oName = o.Properties("Name")
                            For I = 1 To ctlNum
                                If g_Comps(I).Name = oName Then
                                    g_Comps(I).SetSingleConfig ("yscrollcommand|" & Obj.Properties("Name") & ".set")
                                    Debug.Print oName & " assigned to " & Obj.Properties("name")
                                    Assigned = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                If Assigned Then
                    Exit For
                End If
            Next
        End If
    Next
    
End Sub

'Create code
Private Sub CmdGenCode_Click()
    Dim I As Long, cnt As Long, o As Object, sysImport As String
    Dim strHead As New cStrBuilder, strOut As New cStrBuilder, strCmd As New cStrBuilder, strI18n As New cStrBuilder, strTmp As New cStrBuilder
    Dim s As String, finalCode As String, sF As String
    Dim OutOnlyV3 As Boolean, OutRelPos As Boolean, i18n As Boolean, usettk As Boolean
    Dim bUnicodePrefix As Boolean  'Temporarily save the UNICODE prefix mode
    Dim aCompsSorted() As Object 'Used for sorting code output
    
    If LstComps.ListCount = 0 Or LstCfg.ItemCount = 0 Or m_curFrm Is Nothing Then
        Exit Sub
    End If
    
    On Error Resume Next
    s = m_curFrm.Name
    If Err.Number Then
        If MsgBox(L("l_msgGetAttrOfFrmFailed", "Failed in getting property of the form, please reopen the vb project and retry.\nRefresh list of forms now?"), vbInformation + vbYesNo) = vbYes Then
            CmdRefsFormsList_Click
        End If
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    OutOnlyV3 = Not mnuV2andV3Code.Checked
    OutRelPos = mnuRelPos.Checked
    i18n = mnuI18n.Checked
    usettk = mnuUseTtk.Checked
    
    'Absolute coordinates BUG prompt
'    If Not OutRelPos And m_curFrm.Properties("ScaleMode") <> vbTwips Then
'        'If using absolute coordinates, the Frame control only supports vbTwips mode
'        For Each o In m_curFrm.Designer.VBControls
'            If o.ClassName = "Frame" Then
'                MsgBox L("l_msgFrameNotSupportInAbs", "The control 'Frame' is not supported when menu 'Use Relative Position' is unchecked."), vbInformation
'                Exit Sub
'            End If
'        Next
'    End If
'
    bUnicodePrefix = g_bUnicodePrefixU 'Temporarily save, restore at the end of the function
    If OutOnlyV3 Then
        g_bUnicodePrefixU = False  'No prefix needed in V3 mode
    End If
    
    'Update the currently displayed data before outputting the code
    UpdateCfgtoCls LstComps.ListIndex
    
    sysImport = IIf(i18n, "import os, sys, gettext", "import os, sys")
    If OutOnlyV3 Then                                                           'Output code only for PYTHON 3.X
        strHead.Append "#!/usr/bin/env python3"
        strHead.Append "#-*- coding:utf-8 -*-" & vbCrLf
        strHead.Append sysImport
        strHead.Append "from tkinter import *"
        strHead.Append "from tkinter.font import Font"
        If usettk Then strHead.Append "from tkinter.ttk import *"
        strHead.Append "#Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel"
        strHead.Append "from tkinter.messagebox import *"
        strHead.Append IIf(m_HasCommonDialog, "", "#") & "from tkinter import filedialog  #.askopenfilename()"
        strHead.Append IIf(m_HasCommonDialog, "", "#") & "from tkinter import simpledialog  #.askstring()"
        If m_HasCommonDialog Then
            strHead.Append "from tkinter import colorchooser  #.askcolor()"
        End If
        strHead.Append vbCrLf
    Else
        strHead.Append "#!/usr/bin/env python"
        strHead.Append "#-*- coding:utf-8 -*-" & vbCrLf
        strHead.Append sysImport
        strHead.Append "if sys.version_info[0] == 2:"
        strHead.Append "    from Tkinter import *"
        strHead.Append "    from tkFont import Font"
        If usettk Then
            strHead.Append "    from ttk import *"
        End If
        strHead.Append "    #Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel"
        strHead.Append "    from tkMessageBox import *"
        strHead.Append "    #Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkFileDialog"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkSimpleDialog"
        If m_HasCommonDialog Then
            strHead.Append "    import tkColorChooser  #askcolor()"
        End If
        strHead.Append "else:  #Python 3.x"
        strHead.Append "    from tkinter import *"
        strHead.Append "    from tkinter.font import Font"
        If usettk Then
            strHead.Append "    from tkinter.ttk import *"
        End If
        strHead.Append "    from tkinter.messagebox import *"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkinter.filedialog as tkFileDialog"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkinter.simpledialog as tkSimpleDialog    #askstring()"
        If m_HasCommonDialog Then
            strHead.Append "    import tkinter.colorchooser as tkColorChooser  #askcolor()"
        End If
        strHead.Append vbCrLf
    End If
    
    'If there is a status bar control, output the class definition of the status bar control first
    For I = 1 To UBound(g_Comps)  '0 is fixed as the form, no need to judge
        If TypeName(g_Comps(I)) = "clsStatusbar" Then
            strHead.Append g_Comps(I).WidgetCode(OutOnlyV3)
            Exit For
        End If
    Next
    
    'If any control has ToolTipText set, output the custom class definition of the Tooptip first
    For I = 1 To UBound(g_Comps)  '0 is fixed as the form, no need to judge
        If g_Comps(I).hasAttribute("tooltip") And g_Comps(I).GetAttrCurrentValue("tooltip") <> "" Then
            CreateTooltipClassCode strHead
            Exit For
        End If
    Next
    
    'Internationalization support function header
    If i18n Then
        strI18n.Append vbCrLf & "    def retranslateUi(self):"
    End If
    
    'Command callback code block
    strTmp.Reset
    If i18n Then
        strTmp.Append vbCrLf & "        tr = gettext.translation('lang', localedir='i18n', languages=['en'], fallback=True)"
        strTmp.Append "        tr.install()"
        strTmp.Append "        self.retranslateUi()"
    End If
    strCmd.Append vbCrLf
    strCmd.Append "class Application(Application_ui):"
    strCmd.Append "    def __init__(self, master):"
    If OutOnlyV3 Then
        strCmd.Append "        super().__init__(master)" & strTmp.toString(vbCrLf) & vbCrLf
    Else
        strCmd.Append "        Application_ui.__init__(self, master)" & strTmp.toString(vbCrLf) & vbCrLf
    End If
    strTmp.Reset
    
    'Main GUI code block
    strOut.Append "class Application_ui(Frame):"
    strOut.Append "    def __init__(self, master):"
    If OutOnlyV3 Then
        strOut.Append "        super().__init__(master)"
    Else
        strOut.Append "        Frame.__init__(self, master)"
    End If
    g_Comps(0).toString strOut, strCmd, strI18n, OutRelPos, usettk  'g_Comps(0) is fixed as Form
    strOut.Append "        self.createWidgets()" & vbCrLf
    
    strOut.Append "    def createWidgets(self):"
    strOut.Append "        self." & WTOP & " = self.winfo_toplevel()" & vbCrLf
    If usettk Then strOut.Append "        self.style = Style()" & vbCrLf
    
    'Sort the controls according to dependencies
    cnt = 0
    For I = 1 To UBound(g_Comps)
        If g_Comps(I).EnableOutByMainForm Then
            ReDim Preserve aCompsSorted(cnt) As Object
            Set aCompsSorted(cnt) = g_Comps(I)
            cnt = cnt + 1
        End If
    Next
    If cnt > 0 Then
        SortWidgets aCompsSorted, cnt
    End If
    
    'Traverse each control, each control outputs its own interface generation code
    For I = 0 To cnt - 1
        aCompsSorted(I).toString strOut, strCmd, strI18n, OutRelPos, usettk
        strOut.Append ""  'Separate two controls with a blank line
    Next
    
    'Concatenate code segments and output to the text box
    If Not i18n Then
        strI18n.Reset
    End If
    strCmd.Append "if __name__ == ""__main__"":"
    strCmd.Append "    " & WTOP & " = Tk()"
    strCmd.Append "    Application(" & WTOP & ").mainloop()"
    strCmd.Append vbCrLf
    finalCode = strHead.toString(vbCrLf) & strOut.toString(vbCrLf) & strI18n.toString(vbCrLf) & strCmd.toString(vbCrLf)
    
    'VB's TEXTBOX supports up to 65K text
    If Len(finalCode) > 65000 Then
        MsgBox L("l_msgCodeTooBig", "Size of generated code is too big to load into TextBox, please choose a file to save it."), vbInformation
        sF = FileDialog(Me, True, L("l_fdSave", "Save file to:"), "*.py", m_prevsf)
        
        If Len(sF) > 0 Then
            If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  'If the filename has no extension, automatically add .py extension
            Utf8File_Write_VB sF, finalCode
            m_prevsf = sF
        End If
        TxtCode.Text = ""
    Else
        TxtCode.Text = finalCode
    End If
    
    strOut.Reset
    strHead.Reset
    strCmd.Reset
    strI18n.Reset
    
    g_bUnicodePrefixU = bUnicodePrefix    'Restore UNICODE prefix mode
End Sub

'Create custom Tooltip class
'Because VB does not allow too many continuation characters, it is divided into three parts and then concatenated in the program
'Usage of this control:
'  self.t = Tooltip(self.widget, 'tooltip text')
Private Sub CreateTooltipClassCode(ByRef strOut As cStrBuilder)
    Const strCls1 As String = "class Tooltip:" & vbCrLf & _
"    def __init__(self, widget, text, bg='#FFFFEA', pad=(5, 3, 5, 3), waittime=500, wraplength=300):" & vbCrLf & _
"        self.waittime = waittime" & vbCrLf & _
"        self.wraplength = wraplength" & vbCrLf & _
"        self.widget = widget" & vbCrLf & _
"        self.text = text" & vbCrLf & _
"        self.widget.bind('<Enter>', self.onEnter)" & vbCrLf & _
"        self.widget.bind('<Leave>', self.onLeave)" & vbCrLf & _
"        self.widget.bind('<ButtonPress>', self.onLeave)" & vbCrLf & _
"        self.bg = bg" & vbCrLf & _
"        self.pad = pad" & vbCrLf & _
"        self.id_ = None" & vbCrLf & _
"        self.tw = None" & vbCrLf & vbCrLf & _
"    def onEnter(self, event=None):" & vbCrLf & _
"        self.schedule()" & vbCrLf & vbCrLf & _
"    def onLeave(self, event=None):" & vbCrLf & _
"        self.unschedule()" & vbCrLf & _
"        self.Hide()" & vbCrLf & vbCrLf & _
"    def schedule(self):" & vbCrLf & _
"        self.unschedule()" & vbCrLf & _
"        self.id_ = self.widget.after(self.waittime, self.Show)" & vbCrLf

    Const strCls2 As String = vbCrLf & _
"    def unschedule(self):" & vbCrLf & _
"        id_ = self.id_" & vbCrLf & _
"        self.id_ = None" & vbCrLf & _
"        if id_:" & vbCrLf & _
"            self.widget.after_cancel(id_)" & vbCrLf & vbCrLf & _
"    def Show(self):" & vbCrLf & _
"        def tip_pos_calculator(widget, label, pad=(5, 3, 5, 3), tip_delta=(15, 10)):" & vbCrLf & _
"            s_width, s_height = widget.winfo_screenwidth(), widget.winfo_screenheight()" & vbCrLf & _
"            width, height = (pad[0] + label.winfo_reqwidth() + pad[2], pad[1] + label.winfo_reqheight() + pad[3])" & vbCrLf & _
"            mouse_x, mouse_y = widget.winfo_pointerxy()" & vbCrLf & _
"            x1, y1 = mouse_x + tip_delta[0], mouse_y + tip_delta[1]" & vbCrLf & _
"            if x1 + width > s_width:" & vbCrLf & _
"                x1 = mouse_x - tip_delta[0] - width" & vbCrLf & _
"            if y1 + height > s_height - 30:" & vbCrLf & _
"                y1 = mouse_y - tip_delta[1] - height" & vbCrLf & _
"                if y1 < 0:" & vbCrLf & _
"                    y1 = 0" & vbCrLf & _
"            return x1, y1" & vbCrLf & vbCrLf & _
"        self.Hide()" & vbCrLf & _
"        self.tw = Toplevel(self.widget)" & vbCrLf & _
"        self.tw.wm_overrideredirect(True)" & vbCrLf & _
"        label = Label(self.tw, text=self.text, justify=LEFT, background=self.bg, relief=RAISED, borderwidth=1, wraplength=self.wraplength)" & vbCrLf & _
"        label.pack(ipadx=1)"

    Const strCls3 As String = vbCrLf & _
"        x, y = tip_pos_calculator(self.widget, label, self.pad)" & vbCrLf & _
"        self.tw.wm_geometry('+%d+%d' % (x, y))" & vbCrLf & vbCrLf & _
"    def Hide(self):" & vbCrLf & _
"        tw = self.tw" & vbCrLf & _
"        if tw:" & vbCrLf & _
"            tw.destroy()" & vbCrLf & _
"        self.tw = None" & vbCrLf & vbCrLf

    strOut.Append strCls1 & strCls2 & strCls3
    
End Sub

Private Sub CmdRefsFormsList_Click()
    
    Dim frm As Object, nScaleMode As Long, emptyForms() As String, emptyFormsCount As Integer, frmName As String, idx As Integer
    
    emptyFormsCount = 0
    ReDim emptyForms(emptyFormsCount) As String
    
    cmbFrms.Clear
    LstComps.Clear
    LstCfg.Clear
    
    If VbeInst.ActiveVBProject Is Nothing Then
        CmdGenCode.Enabled = False
        CmdCopyToClipboard.Enabled = False
        CmdSaveToFile.Enabled = False
        mnuSaveToFile.Enabled = False
        mnuCopyToClipboard.Enabled = False
        mnuAddProperty.Enabled = False
        mnuGenCode.Enabled = False
        Exit Sub
    End If
    
    Me.Caption = m_BriefCaption & " [" & VbeInst.ActiveVBProject.Name & "]"
    
    'Find all forms in the project and add them to the combo box for selection
    For Each frm In VbeInst.ActiveVBProject.VBComponents
        If frm.Type = vbext_ct_VBForm Then
            nScaleMode = frm.Properties("ScaleMode")
            If nScaleMode <> vbTwips And nScaleMode <> vbPoints And nScaleMode <> vbPixels Then
                MsgBox L_F("l_msgFailedScaleMode", "Found form '{0}', but its ScaleMode={1}, now can only support mode 1/2/3.", _
                         frm.Properties("Name"), nScaleMode), vbInformation
            ElseIf frm.Designer.VBControls.Count = 0 Then 'Empty forms without controls are added to the end of the combo box list
                emptyFormsCount = emptyFormsCount + 1
                ReDim Preserve emptyForms(emptyFormsCount) As String
                emptyForms(emptyFormsCount - 1) = frm.Name
            Else
                cmbFrms.AddItem frm.Name
            End If
        End If
    Next
    
    'Add forms without controls to the end of the list
    If emptyFormsCount > 0 Then
        For idx = 0 To UBound(emptyForms)
            frmName = emptyForms(idx)
            If Len(frmName) > 0 Then
                cmbFrms.AddItem frmName
            End If
        Next
    End If
    
    If cmbFrms.ListCount >= 1 Then
        cmbFrms.ListIndex = 0      'Trigger cmbFrms_Click
    Else
        CmdGenCode.Enabled = False
        CmdCopyToClipboard.Enabled = False
        CmdSaveToFile.Enabled = False
        mnuSaveToFile.Enabled = False
        mnuCopyToClipboard.Enabled = False
        mnuAddProperty.Enabled = False
        mnuGenCode.Enabled = False
    End If
    
End Sub

Private Sub CmdSaveToFile_Click()
    Me.PopupMenu mnuSaveToFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim I As Long
    If TxtCode.Width = Me.ScaleWidth Then
        TxtCode_DblClick
        Cancel = True
    ElseIf TxtTips.Width = Me.ScaleWidth Then
        TxtTips_DblClick
        Cancel = True
    ElseIf UBound(m_saTmpFile) > 0 Then 'Delete temporary files used for preview
        On Error Resume Next
        For I = 0 To UBound(m_saTmpFile)
            If Len(m_saTmpFile(I)) Then
                Kill m_saTmpFile(I)
            End If
        Next
        On Error GoTo 0
        ReDim m_saTmpFile(0) As String
    End If
End Sub

Private Sub Form_Resize()
    
    Dim sLstCfgWidth As Single
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If m_TxtCodeExpanded Then
        TxtCode_DblClick
    ElseIf m_TxtTipsExpanded Then
        TxtTips_DblClick
    End If
    
    If Me.Width < 9000 Then Me.Width = 9000
    If Me.Height < 6750 Then Me.Height = 6750
    
    If LstCfg.ItemCount > 0 Then LstCfg.TopRow = 0  'Avoid a bug in this control that may cause partial items to not display completely
    
    ResizeForm Me
    
    'Adjust the column width of the configuration list box
    sLstCfgWidth = GetOrignalWidth(LstCfg)
    If LstCfg.Width < sLstCfgWidth Then
        LstCfg.ColWidth(0) = 2260
        LstCfg.ColWidth(1) = 3450
    ElseIf sLstCfgWidth > 1 Then 'Confirm the width value is valid
        LstCfg.ColWidth(0) = 2260 * LstCfg.Width / sLstCfgWidth
        LstCfg.ColWidth(1) = 3450 * LstCfg.Width / sLstCfgWidth
    End If
    
End Sub

Private Sub LstCfg_ItemChecked(Row As Long)
    If InStr(1, " x, y, relx, rely, width, height, relwidth, relheight,", " " & LstCfg.CellText(Row, 0) & ",") Then
        LstCfg.ItemChecked(Row) = True
    ElseIf LstComps.ListCount > 0 And LstComps.ListIndex >= 0 Then 'Additional mandatory properties
        Select Case Mid$(LstComps.List(LstComps.ListIndex), InStrRev(LstComps.List(LstComps.ListIndex), " ") + 1)
            Case "(TextBox)"
                If LstCfg.CellText(Row, 0) = "textvariable" Then LstCfg.ItemChecked(Row) = True
            Case "(Line)"
                If LstCfg.CellText(Row, 0) = "orient" Then LstCfg.ItemChecked(Row) = True
        End Select
    End If
    
    'Update the values in the list to the instance object and array
    UpdateCfgtoCls m_PrevCompIdx
    
End Sub

Private Sub LstCfg_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = &H41 Or KeyCode = &H61) And Shift = vbCtrlMask Then 'Ctrl+a: Select all, mainly for testing purposes
        Dim I As Long
        LstCfg.Redraw = False
        For I = 0 To LstCfg.ItemCount - 1
            LstCfg.ItemChecked(I) = True
        Next
        LstCfg.Redraw = True
    End If
End Sub

Private Sub LstCfg_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 0 Or InStr(1, " x, y, relx, rely, width, height, relwidth, relheight, ", " " & LstCfg.CellText(Row, 0) & ",") Then
        Cancel = True
    Else
        'Automatically add the required property selection item. If there is an optional list, select from the dropdown list; otherwise, switch to text box
        FillcmbEdit Row, Col
    End If
    
End Sub

Private Sub LstCfg_RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
    If NewValue <> "" Then
        LstCfg.ItemChecked(Row) = True
    End If
End Sub

Private Sub LstCfg_RowColChanged()
    If LstComps.ListIndex >= 0 Then
        TxtTips.Text = g_Comps(LstComps.ListIndex).Tips(LstCfg.CellText(LstCfg.Row, 0))
    End If
End Sub

Private Sub LstComps_Click()
    
    Dim ctl As Object, s As String
    
    If LstComps.ListCount = 0 Or LstComps.ListIndex < 0 Then Exit Sub
    
    On Error Resume Next
    s = m_curFrm.Name
    If Err.Number Then
        If MsgBox(L("l_msgGetAttrOfFrmFailed", "Failed in getting property of the form, please reopen the vb project and retry.\nRefresh list of forms now?"), vbInformation + vbYesNo) = vbYes Then
            CmdRefsFormsList_Click
        End If
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    'Update the values in the list to the instance object
    UpdateCfgtoCls m_PrevCompIdx
    
    FetchCfgFromCls LstComps.ListIndex
    
    m_PrevCompIdx = LstComps.ListIndex
    
    'Display control description
    TxtTips.Text = g_Comps(LstComps.ListIndex).Description
    
    'Select the corresponding control
    m_curFrm.Designer.SelectedVBControls.Clear
    For Each ctl In m_curFrm.Designer.VBControls
        If ctl.Properties("Name") = Left(LstComps.List(LstComps.ListIndex), InStr(1, LstComps.List(LstComps.ListIndex), " ") - 1) Then
            ctl.InSelection = True
            Exit For
        End If
    Next
    
End Sub

'Fetch configuration information from the object to the list box
Private Sub FetchCfgFromCls(idx As Long)
    
    Dim nRow As Long, cfg As Variant, cItms As Collection
    
    If idx < 0 Or idx > UBound(g_Comps) Then Exit Sub
    
    LstCfg.Redraw = False
    If LstCfg.ItemCount > 0 Then LstCfg.TopRow = 0  'This operation is to avoid a bug in GridOcx where some items cannot be fully displayed after switching controls
    LstCfg.Clear
    Set cItms = g_Comps(idx).Allitems()
    For Each cfg In cItms
        nRow = LstCfg.AddItem(Left(cfg, InStr(1, cfg, "|") - 1))
        LstCfg.CellText(nRow, 1) = Mid(cfg, InStr(1, cfg, "|") + 1, InStrRev(cfg, "|") - InStr(1, cfg, "|") - 1)
        LstCfg.ItemChecked(nRow) = CLng(Mid(cfg, InStrRev(cfg, "|") + 1))
    Next
    LstCfg.Redraw = True
    
End Sub

'Update configuration to the instance object, idx indicates which control the properties displayed on LstCfg belong to.
Private Sub UpdateCfgtoCls(idx As Long)
    Dim s As String, I As Long
    
    If idx < 0 Or idx > UBound(g_Comps) Then Exit Sub
    
    LstCfg.UpdateIfPending
    LstCfg.Refresh
    
    s = ""
    For I = 0 To LstCfg.ItemCount - 1
        If LstCfg.ItemChecked(I) Then
            s = s & IIf(I > 0, "|", "") & LstCfg.CellText(I, 0) & "|" & LstCfg.CellText(I, 1)
        End If
    Next
    
    If Len(s) Then g_Comps(idx).SetConfig s
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

'Add custom configuration
Private Sub mnuAddProperty_Click()
    
    Dim s As String, sa() As String, nRow As Long, I As Long
    
    If LstCfg.ItemCount <= 0 Then Exit Sub
    
    s = InputBox(L("l_ProForAddAttr", "Please input string format as 'Property=Value', for example x=20.\nCase sensitive."), App.Title)
    s = Trim(s)
    If Len(s) <= 0 Then
        Exit Sub
    End If
    
    sa = Split(s, "=")
    If UBound(sa) < 1 Then Exit Sub
    
    ' If the entered property already exists, overwrite the original value
    sa(0) = Trim(sa(0))
    For I = 0 To LstCfg.ItemCount - 1
        If LstCfg.CellText(I, 0) = sa(0) Then
            LstCfg.CellText(I, 1) = Trim(sa(1))
            Exit For
        End If
    Next
    'Add a new property
    If I >= LstCfg.ItemCount Then
        I = LstCfg.AddItem(Trim(sa(0)))
        LstCfg.CellText(I, 1) = Trim(sa(1))
    End If
    
    LstCfg.ItemChecked(I) = True
    UpdateCfgtoCls m_PrevCompIdx
    
End Sub

Private Sub mnuCopyToClipAll_Click()
    Clipboard.Clear
    Clipboard.SetText TxtCode.Text
End Sub

Private Sub mnuCopyToClipUiOnly_Click()
    
    Dim s As String, nui As Long, napp As Long, idx As Long
    
    'Analyze the code, only extract Application_ui(), regular expressions can also be used, but here simple string analysis is used
    s = TxtCode.Text
    nui = InStr(1, s, "class Application_ui(Frame):")
    napp = InStr(1, s, "class Application(Application_ui):")
    If nui > 0 And napp > 0 Then
        Clipboard.Clear
        Clipboard.SetText Mid(s, nui, napp - nui - 2) 'Remove one newline
    Else
        MsgBox L("l_msgNoClsUi", "Class 'Application_ui' not found in code!"), vbInformation
    End If
    
End Sub

'Encode file with base64
Private Sub mnuEncodeAFile_Click()
    frmEncodeAFile.Show vbModal
End Sub

'Check for updates
Private Sub mnuCheckUpdate_Click()
    Dim data As String, ver As String, currVer As String, skipVer As String
    Dim info As Variant, latestVerInfo As Variant
    data = HttpGetResponse(OFFICIAL_UPDATE_INFO)
    If Len(data) = 0 Then
        MsgBox L("l_msgFetchVerInfoFail", "Error fetching version update information."), vbInformation
        Exit Sub
    End If
    data = Trim(data)
    ParseJSONString2 data, info
    If IsArray(info) = False Then
        MsgBox L("l_msgParseVerInfoFail", "Unable to parse the version information."), vbInformation
        Exit Sub
    End If
    If IsObject(info(0)) = False Then
        MsgBox L("l_msgParseVerInfoFail", "Unable to parse the version information."), vbInformation
        Exit Sub
    End If
    
    On Error GoTo verErrHandler:
    Set latestVerInfo = info(0)  'The first element is the latest version, which is a dictionary object
    ver = latestVerInfo("tag_name") 'The author uses tag_name to save the version number when releasing a version on GitHub
    If Len(ver) > 0 Then
        currVer = g_AppVerString
        If isVersionNewerThan(ver, currVer) Then
            Load frmNewVer
            frmNewVer.lblInfo.Caption = L("l_msgFoundNewVersion", "Found new version: ") & ver
            frmNewVer.lblInfo.Tag = ver
            frmNewVer.Show vbModal
            Exit Sub
        End If
    End If
verErrHandler:
    MsgBox L("l_msgYourVerIsLastest", "Your version is the latest."), vbInformation
End Sub

Private Sub mnuFile_Click()
    mnuGenCode.Enabled = LstComps.ListCount > 0
End Sub

Private Sub mnuGenCode_Click()
    CmdGenCode_Click
End Sub

Private Sub mnuI18n_Click()
    Dim o As Object
    mnuI18n.Checked = Not mnuI18n.Checked
    SaveSetting App.Title, "Settings", "i18n", IIf(mnuI18n.Checked, "1", "0")
End Sub

Private Sub mnuLng_Click(Index As Integer)
    Dim I As Long
    
    If m_nLngNum = 0 Then Exit Sub
    
    For I = 0 To m_nLngNum - 1
        mnuLng(I).Checked = False
    Next
    
    mnuLng(Index).Checked = True
    SaveSetting App.Title, "Settings", "Language", mnuLng(Index).Caption
    
    ChangeLanguage (mnuLng(Index).Caption)
End Sub

Private Sub mnuSaveAll_Click()
    Dim sF As String
    sF = FileDialog(Me, True, L("l_fdSave", "Save file to:"), "*.py", m_prevsf)
    
    If Len(sF) Then
        If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  'If the filename has no extension, automatically add .py extension
        Utf8File_Write_VB sF, TxtCode.Text
    End If
    
    m_prevsf = sF
End Sub

'Only output the UI generation class, used when the framework has been built and some code has been written. Now modifying the layout without affecting other code
Private Sub mnuSaveUiOnly_Click()
    
    Dim sF As String, s As String, nui As Long, napp As Long
    
    'Analyze the code, only extract main(), regular expressions can also be used, but here simple string analysis is used
    s = TxtCode.Text
    nui = InStr(1, s, "class Application_ui(Frame):")
    napp = InStr(1, s, "class Application(Application_ui):")
    If nui > 0 And napp > 0 Then
        sF = FileDialog(Me, True, L("l_fdSave", "Save file to:"), "*.py", m_prevsf)
        If Len(sF) Then
            If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  'If the filename has no extension, automatically add .py extension
            Utf8File_Write_VB sF, Mid(s, nui, napp - nui)
        End If
    Else
        MsgBox L("l_msgNoClsUi", "Class 'Application_ui' not found in code!"), vbInformation
    End If
    
    m_prevsf = sF
    
End Sub

Private Sub mnuPreview_Click()
    
    Dim bExeExisted As Boolean, sTmpFile As String
    Dim nIdxIcon1 As Long, nIdxIcon2 As Long, s As String, sCode As String, sFrmFile As String
    
    'First check if PYTHONEXE exists
    If Len(g_PythonExe) = 0 Then
        mnuPythonExe_Click
    ElseIf Dir(g_PythonExe) = "" Then
        g_PythonExe = ""
        mnuPythonExe_Click  'If the location of python.exe is not set, open the settings interface to set it first
    Else
        bExeExisted = True
    End If
    
    'Check again after the previous settings, if incorrect, exit
    If Not bExeExisted And Len(g_PythonExe) = 0 Then
        Exit Sub
    End If
    
    'Create a temporary file
    sTmpFile = CreateTempFile("vt")
    If Len(sTmpFile) Then
        'Temporarily save the temporary filename, all deleted when the ADDIN exits
        ReDim Preserve m_saTmpFile(UBound(m_saTmpFile) + 1) As String
        m_saTmpFile(UBound(m_saTmpFile)) = sTmpFile
        
        sCode = TxtCode.Text
        
        'For a better experience, if an icon is set, copy the icon to the temporary folder [assuming the icon is in the same directory as the form file]
        nIdxIcon1 = InStr(1, sCode, ".iconbitmap(default=r'")
        If nIdxIcon1 > 0 Then
            nIdxIcon1 = nIdxIcon1 + Len(".iconbitmap(default=r'")
            nIdxIcon2 = InStr(nIdxIcon1 + 1, sCode, "'")
            If nIdxIcon2 > 0 And (nIdxIcon2 - nIdxIcon1 < 256) Then
                s = Mid$(sCode, nIdxIcon1, nIdxIcon2 - nIdxIcon1)
                On Error Resume Next
                sFrmFile = m_curFrm.FileNames(1)
                On Error GoTo 0
                If Len(sFrmFile) Then
                    On Error Resume Next
                    FileCopy AddSlash(PathName(sFrmFile)) & s, AddSlash(PathName(sTmpFile)) & s
                    If Err.Number = 0 Then
                        ReDim Preserve m_saTmpFile(UBound(m_saTmpFile) + 1) As String
                        m_saTmpFile(UBound(m_saTmpFile)) = AddSlash(PathName(sTmpFile)) & s
                        sCode = Left$(sCode, nIdxIcon1 - 1) & m_saTmpFile(UBound(m_saTmpFile)) & Mid$(sCode, nIdxIcon2)
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
        
        Utf8File_Write_VB sTmpFile, sCode
        Shell Chr(34) & g_PythonExe & """ """ & sTmpFile & Chr(34)
    Else
        MsgBox L("l_msgCreateTempFileFailed", "Failed in creating a temp file."), vbInformation
    End If
    
End Sub

Private Sub mnuPythonExe_Click()
    Dim sExe As String, sExes() As String, I As Long
    
    Load frmOption
    
    'Add all installed Pythons in the system to the combo box for selection
    frmOption.cmbPythonExe.Clear
    sExes = GetAllInstalledPython()
    If UBound(sExes) >= 0 Then
        For I = 0 To UBound(sExes)
            frmOption.cmbPythonExe.AddItem sExes(I)
        Next
    End If
    
    'Read the saved configuration from the registry
    sExe = GetSetting(App.Title, "Settings", "PythonExe", "")
    If Len(sExe) > 0 Then
        If Dir(sExe) = "" Then
            sExe = ""
        End If
    End If
    
    If Len(sExe) Then
        frmOption.cmbPythonExe.Text = sExe
    ElseIf frmOption.cmbPythonExe.ListCount > 0 Then
        frmOption.cmbPythonExe.ListIndex = 0
    End If
    
    frmOption.Show vbModal  'In the frmOption form, the g_PythonExe variable will be set and saved to the registry
    
End Sub

Private Sub mnuQuit_Click()
    mConnect.Hide
End Sub

Private Sub mnuRefreshForms_Click()
    CmdRefsFormsList_Click
End Sub

Private Sub mnuRelPos_Click()
    Dim o As Object
    mnuRelPos.Checked = Not mnuRelPos.Checked
    
    'Absolute coordinates
'    If Not m_curFrm Is Nothing Then
'        If Not mnuRelPos.Checked And m_curFrm.Properties("ScaleMode") <> vbTwips Then
'            'If using absolute coordinates, Frame control is not supported
'            For Each o In m_curFrm.Designer.VBControls
'                If o.ClassName = "Frame" Then
'                    MsgBox L("l_msgFrameNotSupportInAbs", "The control 'Frame' is not supported when menu 'Use Relative Position' is unchecked."), vbInformation
'                    mnuRelPos.Checked = True
'                    Exit For
'                End If
'            Next
'        End If
'    End If
'
    SaveSetting App.Title, "Settings", "RelPos", IIf(mnuRelPos.Checked, "1", "0")
End Sub

Private Sub mnuUnicodePrefixU_Click()
    
    If MsgBox(L("l_msgChangePrefixU", "Changing the prefix of UNICODE string will refresh all widgets of the form, \nContinue?"), vbQuestion + vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    mnuUnicodePrefixU.Checked = Not mnuUnicodePrefixU.Checked
    g_bUnicodePrefixU = mnuUnicodePrefixU.Checked
    SaveSetting App.Title, "Settings", "UnicodePrefix", IIf(mnuUnicodePrefixU.Checked, "1", "0")
    
    CmdRefsFormsList_Click
    
End Sub

Private Sub mnuUseTtk_Click()
    Dim I As Long, s As String
    
    If LstComps.ListCount > 0 And LstComps.ListIndex >= 0 Then
        If InStr(1, LstComps.List(LstComps.ListIndex), "ComboBox") Then
            LstComps_Click                                                      'Save the configuration first to avoid incorrect configuration after switching the combo box
        End If
    End If
    
    mnuUseTtk.Checked = Not mnuUseTtk.Checked
    
    'Check if there are controls that require TTK, if any, do not allow unchecking the TTK option
    If Not mnuUseTtk.Checked Then
        For I = 0 To LstComps.ListCount - 1
            s = Mid(LstComps.List(I), InStr(1, LstComps.List(I), "(") + 1)
            s = Left(s, Len(s) - 1)
            If InStr(1, " ProgressBar, TreeView, TabStrip, Line, ", " " & s & ",") > 0 Then
                MsgBox L("l_msgCantCancelTTK", "Can't uncheck the menu 'Use TTK Themed Library' because some widgets require TTK."), vbInformation
                mnuUseTtk.Checked = True
                Exit For
            End If
        Next
    End If
    
    'Switch the TTK property of the combo box adapter
    If LstComps.ListCount > 0 Then
        For I = 0 To UBound(g_Comps)
            If TypeName(g_Comps(I)) = "clsComboboxAdapter" Then
                g_Comps(I).TTK = mnuUseTtk.Checked
            End If
        Next
        
        If LstComps.ListIndex >= 0 Then
            If InStr(1, LstComps.List(LstComps.ListIndex), "ComboBox") Then
                FetchCfgFromCls LstComps.ListIndex                              'Retrieve combo box information again
            End If
        End If
        LstComps_Click
    End If
    
    SaveSetting App.Title, "Settings", "UseTtk", IIf(mnuUseTtk.Checked, "1", "0")
    
End Sub

Private Sub mnuV2andV3Code_Click()
    mnuV2andV3Code.Checked = Not mnuV2andV3Code.Checked
    SaveSetting App.Title, "Settings", "V2andV3Code", IIf(mnuV2andV3Code.Checked, "1", "0")
End Sub

'Automatically fill the combo box content for editing
Private Sub FillcmbEdit(Row As Long, Col As Long)
    
    Dim sa() As String, I As Long, nEditType As Long, fn As String
    Static s_NoFirstcmbEditList As Boolean, s_NoFirstcmbEditCombo As Boolean
    
    If LstComps.ListCount = 0 Or LstComps.ListIndex < 0 Then Exit Sub
    
    '0 indicates built-in text editor, 1 indicates dropdown list only, 2 indicates dropdown list with text input
    nEditType = g_Comps(LstComps.ListIndex).GetAttrValueList(LstCfg.CellText(Row, 0), sa)
    
    If nEditType = 1 Then
        LstCfg.BindControl 1, cmbEditList
        cmbEditList.Clear
        cmbEditList.AddItem "" 'Place an empty string in the first line, so the corresponding parameter can be unset.
        For I = 0 To UBound(sa)
            cmbEditList.AddItem sa(I)
        Next
        For I = 0 To cmbEditList.ListCount - 1
            If cmbEditList.List(I) = LstCfg.CellText(Row, Col) Then
                cmbEditList.ListIndex = I
                Exit For
            End If
        Next
        cmbEditList.Refresh
    ElseIf nEditType = 2 Then
        LstCfg.BindControl 1, cmbEditCombo
        cmbEditCombo.Clear
        For I = 0 To UBound(sa)
            cmbEditCombo.AddItem sa(I)
        Next
        cmbEditCombo.Text = LstCfg.CellText(Row, Col)
        cmbEditCombo.Refresh
    Else
        LstCfg.BindControl 1, Nothing  'Use built-in text editor
    End If
    
End Sub

Private Sub stabar_DblClick()
    MsgBox L("l_msgCtlsSupported", "Controls supported:") & vbCrLf & "Menu, Label, TextBox, PictureBox, Frame, CommandButton, CheckBox, OptionButton, ComboBox," & vbCrLf & _
            "ListBox, HScrollBar, VScrollBar, Slider, ProgressBar, TreeView, StatusBar, CommonDialog, Line" & vbCrLf & vbCrLf
End Sub

Private Sub stabar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetText OFFICIAL_SITE
    End If
End Sub

Private Sub TxtCode_Change()
    mnuPreview.Enabled = (Len(TxtCode.Text) > 0)
End Sub

Private Sub TxtCode_DblClick()
    Static s_l As Single, s_t As Single, s_w As Single, s_h As Single
    Static s_txt As String
    
    If m_TxtCodeExpanded Then
        TxtCode.Move s_l, s_t, s_w, s_h
        m_TxtCodeExpanded = False
    Else
        s_l = TxtCode.Left
        s_t = TxtCode.Top
        s_w = TxtCode.Width
        s_h = TxtCode.Height
        TxtCode.ZOrder 0
        TxtCode.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        m_TxtCodeExpanded = True
    End If
End Sub

Private Sub TxtCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And TxtCode.Width = Me.ScaleWidth Then
        TxtCode_DblClick
    End If
End Sub

'View mouse and key message details
Private Sub TxtTips_DblClick()
    Static s_l As Single, s_t As Single, s_w As Single, s_h As Single
    Static s_txt As String
    
    Dim s As String
    s = TxtTips.Text
    If Len(s) Then
        If Left(s, Len("bindcommand")) = "bindcommand" Then
            s_l = TxtTips.Left
            s_t = TxtTips.Top
            s_w = TxtTips.Width
            s_h = TxtTips.Height
            TxtTips.ZOrder 0
            TxtTips.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            s_txt = TxtTips.Text
            m_TxtTipsExpanded = True
            TxtTips.Text = "<Double-click again to return>" & vbCrLf & _
            "bindcommand" & vbCrLf & _
            "Events bound using bind(), if multiple bindings are needed, use commas to separate them. Leave blank if not needed." & vbCrLf & _
            "All event lists are as follows:" & vbCrLf & _
            "<ButtonPress-n> : Mouse button n pressed, n: 1 (left button); 2 (middle button); 3 (right button)" & vbCrLf & _
            "<Button-n>,<n> : Both are simplified forms of <ButtonPress-n>" & vbCrLf & _
            "<ButtonRelease-n> : Mouse button n released" & vbCrLf & _
            "<Bn-Motion> : Mouse moves while holding down button n" & vbCrLf & _
            "<Double-Button-n> : Mouse button n double-click" & vbCrLf & _
            "<Triple-Button-n> : Mouse button n triple-click" & vbCrLf & _
            "<Enter> : Mouse pointer enters the component" & vbCrLf & _
            "<Leave> : Mouse pointer leaves the component" & vbCrLf & _
            "<FocusIn> / <FocusOut> : Component gains or loses focus" & vbCrLf & _
            "<KeyPress> : Any key pressed" & vbCrLf & _
            "<KeyRelease> : Any key released" & vbCrLf & _
            "<KeyPress-key> : Key pressed, for example, <KeyPress-H> means the H key is pressed, which can be simplified by using double quotes instead of angle brackets to enclose the character, for example: ""H""." & vbCrLf & _
            "<KeyRelease-key> : Key released" & vbCrLf & _
            "<Key> : Any key pressed." & vbCrLf & _
            "<Key-key> : Simplified form of <KeyPress-key>, for example, <Key-H>." & vbCrLf & _
            "<key> : Use the special key definition attached to replace the key, indicating a specific key pressed." & vbCrLf & _
            "<Prefix-key> : Hold down the Prefix while pressing the key, can use single or combination of Alt, Shift, Control, for example <Control-Alt-key>" & vbCrLf
            
            TxtTips.Text = TxtTips.Text & "<Configure> : Triggered after the control size changes." & vbCrLf & _
            "All special key definitions are as follows:" & vbCrLf & _
            "Cancel, Break, BackSpace, Tab, Return," & vbCrLf & _
            "Shift_L, Shift_R, Control_L, Control_R, Alt_L, Alt_R, Pause" & vbCrLf & _
            "Caps_Lock, Escape, Prior (PageUp), Next (PageDown), End, Home, Left, Up, Right, Down, Print," & vbCrLf & _
            "Insert, Delete, F1-12, Num_Lock, Scroll_Lock, space, less"
            TxtTips.SelStart = 1
            TxtTips.SelLength = 0
        ElseIf Left(s, Len("<Double-click again to return>")) = "<Double-click again to return>" Then
            TxtTips.Move s_l, s_t, s_w, s_h
            TxtTips.Text = s_txt
            m_TxtTipsExpanded = False
        End If
    End If
End Sub

Private Sub TxtTips_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And TxtTips.Width = Me.ScaleWidth Then
        TxtTips_DblClick
    End If
End Sub

Private Sub TxtTips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staTips", "Tips of properties.")
End Sub

Private Sub LstComps_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staComps", "List of controls.")
End Sub

Private Sub cmbFrms_GotFocus()
    stabar.SimpleText = L("l_staFrms", "List of forms.")
End Sub

Private Sub CmdCopyToClipboard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCopyCode", "Copy code to clipboard.")
End Sub

Private Sub CmdGenCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCmdGenCode", "Generate python code after confirm properties of controls.")
End Sub

Private Sub CmdQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staQuit", "Quit!")
End Sub

Private Sub CmdRefsFormsList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staRefsFrms", "Resfresh list of forms and controls.")
End Sub

Private Sub CmdSaveToFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCmdSaveFile", "Save code to file (format utf-8 with BOM).")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = ""
End Sub

Private Sub LstCfg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staLstCfg", "List of properties of control, F2/Return/DblClick to modify a property.")
End Sub

Private Sub TxtCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staTxtCode", "Preview python code here, DblClick to zoom out/in.")
End Sub
 
Private Sub stabar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = OFFICIAL_SITE & "  ['Ctrl+Click' copy url to clipboard]"
End Sub

