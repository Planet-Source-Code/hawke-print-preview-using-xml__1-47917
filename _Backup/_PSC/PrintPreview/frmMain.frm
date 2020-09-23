VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRINT PREVIEW"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Preview"
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dtgResult 
      Bindings        =   "frmMain.frx":0000
      Height          =   5415
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoResult 
      Height          =   375
      Left            =   6120
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebPreview 
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   1560
      Width           =   615
      ExtentX         =   1085
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strXMLPath As String
Private m_objRS As ADODB.Recordset  'this recordset will be saved to xml file

Private Const navNoHistory As Integer = 2
Private Const navNoWriteToCache As Integer = 8

Private Sub Form_Load()

    Dim strConnString As String
    'change the path of the database according to suit your needs
    strConnString = ConstructConnString
    
    'change the sql statement accordingly to suit your needs
    adoResult.RecordSource = "SELECT * FROM Accounts"
    adoResult.ConnectionString = strConnString
    adoResult.CursorType = adOpenStatic
    adoResult.LockType = adLockBatchOptimistic
    
    Set dtgResult.DataSource = adoResult
    Set m_objRS = adoResult.Recordset.Clone 'create a clone of the recordset from data control
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_objRS Is Nothing Then Set m_objRS = Nothing
    Set frmMain = Nothing
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err_handler

    If IsFileExist(m_strXMLPath) Then Kill m_strXMLPath   'delete xml file
    
    If TransformXML Then
        Dim objReg As clsRegistry
        Set objReg = New clsRegistry
        
        'the footer setting is stored in the registry.
        'set footer to empty string to prevent user from printing the xml file path
        Call objReg.ResetPrintFooter
        Set objReg = Nothing
    
        WebPreview.Navigate2 m_strXMLPath, navNoHistory & navNoWriteToCache
    Else
        MsgBox "Unable to show print preview screen!", vbExclamation
    End If
    Exit Sub
    
err_handler:
    If Not objReg Is Nothing Then Set objReg = Nothing
    
End Sub

'this function insert stylesheet tag into the xml file to render it as html file
Private Function TransformXML() As Boolean
On Error GoTo err_handler

    Dim strCode As String
    Dim strTempPath As String
    Dim strUnique As String
    Dim objFSO As Scripting.FileSystemObject
    Dim objTextStream As TextStream
    
    'the xml file will be stored in system temp folder
    strTempPath = GetTempDirectory

    'this is to ensure the xml file name is unique and will not conflict with any existing files
    strUnique = CreateGUID
    If strUnique = "ERROR" Then strUnique = "Temp"
    
    m_strXMLPath = strTempPath & strUnique & ".xml"
    m_objRS.Save m_strXMLPath, adPersistXML     'save recordset to the given xml file name
    
    Set objFSO = New Scripting.FileSystemObject
    Set objTextStream = objFSO.OpenTextFile(m_strXMLPath, ForReading, False)
    strCode = objTextStream.ReadAll
    
    'apply the included style sheet (style.xsl) to render the xml file as html
    'the stylesheet header (<?xml-stylesheet type="text/xsl" href=path of stylesheet\style.xsl"?>)
    'must be added to the first line of the xml file
    'edit style.xsl to change the look and feel of the displayed html file
    objTextStream.Close
    Set objTextStream = objFSO.OpenTextFile(m_strXMLPath, ForWriting, False)
    objTextStream.Write "<?xml-stylesheet type=""text/xsl"" href=""" & App.Path & "\style.xsl""?>" & vbCrLf & strCode
    
    objTextStream.Close
    Set objTextStream = Nothing
    Set objFSO = Nothing
    
    TransformXML = True
    Exit Function
    
err_handler:
    If Not objTextStream Is Nothing Then Set objTextStream = Nothing
    If Not objFSO Is Nothing Then Set objFSO = Nothing
    
    TransformXML = False
    Call WriteToEventViewer("TransformXML", Err.Description, Err.Number, Err.Source)
    
End Function

'*** show the IE print preview screen only after navigation is completed.
'when form is first loaded, the url is "http:///". shld prevent the print preview
'screen from loading if thats the case
Private Sub webPreview_DownloadComplete()
    If WebPreview.LocationURL <> "http:///" Then WebPreview.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
End Sub

'delete xml file when user closes print preview screen
Private Sub WebPreview_PrintTemplateTeardown(ByVal pDisp As Object)
    If IsFileExist(m_strXMLPath) Then Kill m_strXMLPath   'delete xml file
End Sub
