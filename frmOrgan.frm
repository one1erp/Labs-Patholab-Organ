VERSION 5.00
Begin VB.Form frmOrgan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox PicRemark 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   7080
      Picture         =   "frmOrgan.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   3817
      Width           =   540
   End
   Begin VB.TextBox txtRemark 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   6885
   End
   Begin VB.ComboBox cmbOrgan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   6885
   End
   Begin VB.ComboBox cmbSide 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3120
      Width           =   6885
   End
   Begin VB.ComboBox cmbTopography 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   6885
   End
   Begin VB.ComboBox cmbProcedureCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   6885
   End
   Begin VB.ComboBox cmbSample 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   5205
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "אישור"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ביטול"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   3855
   End
   Begin VB.Label lblSampleCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   1005
      Width           =   375
   End
   Begin VB.Label lblSampleCountHeader 
      Alignment       =   1  'Right Justify
      Caption         =   "סה""כ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   1005
      Width           =   495
   End
   Begin VB.Label lblSampleType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   6555
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblInternalNumber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblPatient 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblOrgan 
      Alignment       =   1  'Right Justify
      Caption         =   "איבר"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7275
      TabIndex        =   13
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label lblSide 
      Alignment       =   1  'Right Justify
      Caption         =   "צד"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7275
      TabIndex        =   12
      Top             =   3165
      Width           =   975
   End
   Begin VB.Label lblTopography 
      Alignment       =   1  'Right Justify
      Caption         =   "מיקום"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7275
      TabIndex        =   11
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label lblRemark 
      Alignment       =   1  'Right Justify
      Caption         =   "הערה"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7635
      TabIndex        =   10
      Top             =   3885
      Width           =   615
   End
   Begin VB.Label lblProcedureCode 
      Alignment       =   1  'Right Justify
      Caption         =   "קוד טיפול"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7275
      TabIndex        =   9
      Top             =   4725
      Width           =   975
   End
   Begin VB.Label lblSample 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "דגימה"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7275
      TabIndex        =   8
      Top             =   1005
      Width           =   975
   End
End
Attribute VB_Name = "frmOrgan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'constants:
Private Const STATUS_COMPLETE = "C"
Private Const STATUS_INCOMPLETE = "P"
Private Const INVALID_STATUS_FOR_APPROVAL = "'A','R','X'"
Private Const INVALID_STATUS_FOR_SELECTION = "'R','X'"


Private Const CB_FINDSTRING = &H14C
Private Const CB_SHOWDROPDOWN = &H14F
Private Const LB_FINDSTRING = &H18F
Private Const CB_ERR = (-1)

Private Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) _
    As Long
    
    
'input fields:
Public con As Connection
Public strSampleId As String
Public strSdgId As String
Public strOperatorName As String
Public sdg_log As New SdgLog.CreateLog



'computed / input fields:
'1. where there is a code and a name to the field, we hold here the CODE:
'2. in general, these variables reflect the values saved in the DB only,
'   not temporary GUI selections
Private strOrgan As String
Private strSide As String
Private strTopography As String
Private strRemark As String
Private strProcedureCode As String
'Private strBillingCode As String
Public strSampleStatus As String

'Private okClicked As Boolean

'Public strSnomedT As String


'the control status:
'for Sample mode: incomplete if the sample has no Billing Code
'for SDG mode: incomplete if at least one sample has no Billing Code
Private strStatus As String

'changes dynamicaly, without update.
'the organ selection determines it's value:
Private IsRemarkNeeded As Boolean


'was the control opened for a Sample / SDG
Private IsSampleMode As Boolean

'collections to hold the corrent selections
'for the different GUI fields:
Private dicSamplesTypeToId As New Dictionary
Private dicSamplesIdToType As New Dictionary
Private dicOrganNameToCode As New Dictionary
Private dicOrganCodeToName As New Dictionary
Private dicSideNameToCode As New Dictionary
Private dicSideCodeToName As New Dictionary
Private dicTopographyNameToCode As New Dictionary
Private dicTopographyCodeToName As New Dictionary
Private dicProcedureCodeToName As New Dictionary
Private dicProcedureNameToCode As New Dictionary


'if true - should show the sample fields value in the text of the combo box:
'(when the user makes new selection this is NOT the case)
Private IsAutoSelection As Boolean





Public Sub Initialize()
    
    IsAutoSelection = False

    Call ClearMemoryFields

    Call InitSamples
    
    Call ResetScreenBySample
    
    
    
    'Call ReadTopography
    'Call ReadSide
    'Call ReadProcedures
    
End Sub


Private Sub ClearMemoryFields()
    strStatus = ""
    'strSnomedT = ""
    strOrgan = ""
    strSide = ""
    strTopography = ""
    strRemark = ""
    strProcedureCode = ""
    'strBillingCode = ""
End Sub


Public Sub ResetScreenBySample()
    IsAutoSelection = True
    
    Call InitChosenSampleData(strSampleId)
    Call InitRequestData
    Call ReadOrgans
    
    IsAutoSelection = False
End Sub


Private Sub InitSamples()
On Error GoTo ERR_InitSamples
    
    Dim rs As Recordset
    Dim sql As String

    IsSampleMode = strSampleId <> ""
    
    cmbSample.Enabled = Not IsSampleMode

    If strSampleId <> "" And strSampleId <> "0" Then
        strSdgId = GetSdgId(strSampleId)
'    Else
'        strSampleId = GetFirstSampleId(strSdgId)
    End If
 
    sql = " select s.SAMPLE_ID, s.SAMPLE_TYPE, s.NAME "
    sql = sql & " from lims_sys.sample s "
    sql = sql & " where s.SDG_ID = '" & strSdgId & "'"
    sql = sql & " and   s.status not in (" & INVALID_STATUS_FOR_SELECTION & ")  "
    sql = sql & " order by s.SAMPLE_ID"
        
    Set rs = con.Execute(sql)
    
    Call dicSamplesTypeToId.RemoveAll
    Call dicSamplesIdToType.RemoveAll
    cmbSample.Clear
    
    While Not rs.EOF
         
        'Dim strSampleNumber As String
    
        Call dicSamplesTypeToId.Add(nte(rs("NAME")), nte(rs("SAMPLE_ID")))
        Call dicSamplesIdToType.Add(nte(rs("SAMPLE_ID")), nte(rs("NAME")))
        Call cmbSample.AddItem(nte(rs("NAME")))
    
        'strSampleNumber = CStr(dicSamplesTypeToId.Count + 1)
        'Call dicSamplesTypeToId.Add(strSampleNumber, nte(rs("SAMPLE_ID")))
        'Call dicSamplesIdToType.Add(nte(rs("SAMPLE_ID")), strSampleNumber)
        'Call cmbSample.AddItem(strSampleNumber)
        
        
        
        
        'Call dicSamplesTypeToId.Add(nte(rs("SAMPLE_TYPE")), nte(rs("SAMPLE_ID")))
        'Call dicSamplesIdToType.Add(nte(rs("SAMPLE_ID")), nte(rs("SAMPLE_TYPE")))
        'Call cmbSample.AddItem(nte(rs("SAMPLE_TYPE")))
        
'        If strSampleId = nte(rs("SAMPLE_ID")) Then
'            cmbSample.Text = nte(rs("SAMPLE_TYPE"))
'        End If
        
        rs.MoveNext
        
    Wend

    'get relevant sample id
    'if control initialized for SDG - 1st sample
    If Not IsSampleMode Then
        strSampleId = dicSamplesTypeToId.Items(0)
    End If

    cmbSample.Text = dicSamplesIdToType(strSampleId)
    
    lblSampleCount.Caption = dicSamplesIdToType.Count

    Exit Sub
ERR_InitSamples:
MsgBox "ERR_InitSamples" & vbCrLf & Err.Description
End Sub


Private Sub InitRequestData()
On Error GoTo ERR_InitRequestData

    Dim rs As Recordset
    Dim sql As String
    
        'Ashi

    lblPatient.Caption = ""
    lblInternalNumber.Caption = ""
        If strSampleId = "0" Then
            Exit Sub
        End If
'--------------
    
    'the patient label:
    sql = " select cu.U_LAST_NAME, cu.U_FIRST_NAME"
    sql = sql & " from lims_sys.client_user cu,"
    sql = sql & "      lims_sys.sdg_user du"
    sql = sql & " where cu.CLIENT_ID = du.U_PATIENT"
    sql = sql & " and   du.SDG_ID='" & strSdgId & "'"

    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        lblPatient.Caption = nte(rs("U_FIRST_NAME")) & " " & nte(rs("U_LAST_NAME"))
    End If
    
    
    
    'the internal number label:
    lblInternalNumber.Caption = GetSdgName(strSdgId)
    

    Exit Sub
ERR_InitRequestData:
MsgBox "ERR_InitRequestData" & vbCrLf & Err.Description
End Sub


Private Sub ReadOrgans()
On Error GoTo ERR_ReadAllORgans

    Dim rs As Recordset
    Dim sql As String
    
    sql = "  select ou.U_ORGAN_CODE, ou.U_ORGAN_HEBREW_NAME"
    sql = sql & "  from lims_sys.u_norgan o,"
    sql = sql & "       lims_sys.u_norgan_user ou"
    sql = sql & "  where ou.U_NORGAN_ID=o.U_NORGAN_ID"
    sql = sql & "  and   o.VERSION_STATUS='A'    "
    sql = sql & "  order by ou.U_ORGAN_HEBREW_NAME"
    
    Set rs = con.Execute(sql)
    
    
    Call dicOrganCodeToName.RemoveAll
    Call dicOrganNameToCode.RemoveAll
    Call cmbOrgan.Clear
    
    While Not rs.EOF
    
        Call dicOrganCodeToName.Add(nte(rs("U_ORGAN_CODE")), nte(rs("U_ORGAN_HEBREW_NAME")))
        Call dicOrganNameToCode.Add(nte(rs("U_ORGAN_HEBREW_NAME")), nte(rs("U_ORGAN_CODE")))
        Call cmbOrgan.AddItem(nte(rs("U_ORGAN_HEBREW_NAME")))
        rs.MoveNext
        
    Wend
    
    If strOrgan <> "" And IsAutoSelection Then
        cmbOrgan.Text = dicOrganCodeToName(strOrgan)
    End If

    Call ReadTopography

    Exit Sub
ERR_ReadAllORgans:
MsgBox "ERR_ReadAllORgans" & vbCrLf & Err.Description
End Sub


'consider an organ may or may not be already chosen:
Private Sub ReadTopography()
On Error GoTo ERR_ReadTopography

    Dim rs As Recordset
    Dim sql As String
     
    cmbTopography.Clear
    Call dicTopographyCodeToName.RemoveAll
    Call dicTopographyNameToCode.RemoveAll
     
    If cmbOrgan.Text = "" Then
        ReadSide
        ReadProcedures
        ReadRemarkNeeded
        Exit Sub
    End If
     
   
     
     
    sql = " select otu.U_TOPOGRAPHY_CODE, tu.U_TOPOGRAPHY_HEBREW_NAME"
    sql = sql & " from lims_sys.u_norgan_topography ot,"
    sql = sql & "      lims_sys.u_norgan_topography_user otu,"
    'sql = sql & "      lims_sys.u_ntopography t,"
    sql = sql & "      lims_sys.u_ntopography_user tu"
    sql = sql & " where otu.U_NORGAN_TOPOGRAPHY_ID=ot.U_NORGAN_TOPOGRAPHY_ID"
    'sql = sql & " and   tu.U_NTOPOGRAPHY_ID = t.U_NTOPOGRAPHY_ID"
    'sql = sql & " and   t.VERSION_STATUS='A'"
    sql = sql & " and   tu.U_TOPOGRAPHY_CODE=otu.U_TOPOGRAPHY_CODE"
    sql = sql & " and   otu.u_organ_code='" & dicOrganNameToCode(cmbOrgan.Text) & "'"
    sql = sql & " order by  tu.U_TOPOGRAPHY_HEBREW_NAME"

    Set rs = con.Execute(sql)

    While Not rs.EOF
        
        If Not dicTopographyCodeToName.Exists(nte(rs("U_TOPOGRAPHY_CODE"))) Then

            Call dicTopographyCodeToName.Add(nte(rs("U_TOPOGRAPHY_CODE")), nte(rs("U_TOPOGRAPHY_HEBREW_NAME")))
            Call dicTopographyNameToCode.Add(nte(rs("U_TOPOGRAPHY_HEBREW_NAME")), nte(rs("U_TOPOGRAPHY_CODE")))
            Call cmbTopography.AddItem(nte(rs("U_TOPOGRAPHY_HEBREW_NAME")))

        End If
        
        rs.MoveNext
        
    Wend
    
     
    Call dicTopographyCodeToName.Add("?", "?")
    Call dicTopographyNameToCode.Add("?", "?")
    Call cmbTopography.AddItem("?")
    
    If strTopography <> "" And IsAutoSelection Then
        cmbTopography.Text = dicTopographyCodeToName(strTopography)
    ElseIf dicTopographyCodeToName.Count = 1 Then
        cmbTopography.Text = dicTopographyCodeToName.Items(0)
    ElseIf dicTopographyCodeToName.Count = 2 Then
        cmbTopography.Text = dicTopographyCodeToName.Items(0)
    End If
   
   
    
    ReadSide
    ReadProcedures
    ReadRemarkNeeded

    Exit Sub
ERR_ReadTopography:
MsgBox "ERR_ReadTopography" & vbCrLf & Err.Description
End Sub


'get possible sides by selected organ and topography:
Private Sub ReadSide()
On Error GoTo ERR_ReadSide
    
    Dim rs As Recordset
    Dim sql As String
    Dim dAllSides As New Dictionary
    
    cmbSide.Clear
    Call dicSideCodeToName.RemoveAll
    Call dicSideNameToCode.RemoveAll
    Call dAllSides.RemoveAll
    
    If cmbOrgan.Text = "" Or cmbTopography.Text = "" Then
        Exit Sub
    End If
    

    'read all side codes:
    sql = " select sd.NAME, sd.DESCRIPTION"
    sql = sql & " from lims_sys.u_nside sd"
    sql = sql & " order by sd.U_NSIDE_ID"
        
    Set rs = con.Execute(sql)
    
    While Not rs.EOF
    
        Call dAllSides.Add(nte(rs("NAME")), nte(rs("DESCRIPTION")))
        rs.MoveNext
        
    Wend
    

    'read relevant sides for current organ & topography:
    sql = " select otu.U_R_SNOMED, otu.U_L_SNOMED, otu.U_O_SNOMED, otu.U_RL_SNOMED, otu.U_NS_SNOMED "
    sql = sql & " from lims_sys.u_norgan_topography_user otu"
    sql = sql & " where otu.U_ORGAN_CODE='" & dicOrganNameToCode(cmbOrgan.Text) & "'"
    sql = sql & " and   otu.U_TOPOGRAPHY_CODE='" & dicTopographyNameToCode(cmbTopography.Text) & "'"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        
        Dim i As Integer
        
        For i = 0 To rs.Fields.Count - 1
            
            If nte(rs.Fields(i)) <> "" Then
                
                Call dicSideCodeToName.Add(dAllSides.Keys(i), dAllSides.Items(i))
                Call dicSideNameToCode.Add(dAllSides.Items(i), dAllSides.Keys(i))
                Call cmbSide.AddItem(dAllSides.Items(i))
                
            End If
            
        Next i
        
    End If

    Call dicSideCodeToName.Add("?", "?")
    Call dicSideNameToCode.Add("?", "?")
    Call cmbSide.AddItem("?")
    
    If strSide <> "" And IsAutoSelection Then
        cmbSide.Text = dicSideCodeToName(strSide)
    ElseIf dicSideCodeToName.Count = 1 Then
        cmbSide.Text = dicSideCodeToName.Items(0)
     ElseIf dicSideCodeToName.Count = 2 Then
        cmbSide.Text = dicSideCodeToName.Items(0)
    End If
    
  
    
    
    Exit Sub
ERR_ReadSide:
MsgBox "ERR_ReadSide" & vbCrLf & Err.Description
End Sub


Private Sub ReadProcedures()
On Error GoTo ERR_ReadProcedures

    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    
    Call dicProcedureCodeToName.RemoveAll
    Call dicProcedureNameToCode.RemoveAll
    Call cmbProcedureCode.Clear
    
    
    If cmbOrgan.Text = "" Or cmbTopography.Text = "" Or _
        cmbTopography.Text = "?" Then
        'Call ReadBillingCode
        Exit Sub
    End If
    
    sql = " select otu.U_PROCEDURES "
    sql = sql & " from lims_sys.u_norgan_topography_user otu"
    sql = sql & " where otu.U_ORGAN_CODE='" & dicOrganNameToCode(cmbOrgan.Text) & "'"
    sql = sql & " and   otu.U_TOPOGRAPHY_CODE='" & dicTopographyNameToCode(cmbTopography.Text) & "'"
    
    Set rs = con.Execute(sql)

    If Not rs.EOF Then
    
        Dim strProcedures As String
        Dim strProcedureCode_ As String
        Dim strProcedureName_ As String
        
        strProcedures = nte(rs("U_PROCEDURES"))
        
        While strProcedures <> ""
        
            strProcedureCode_ = Trim(getNextStr(strProcedures, ","))
            strProcedureName_ = GetProcedureName(strProcedureCode_)
             
            Call dicProcedureCodeToName.Add(strProcedureCode_, _
                                            strProcedureCode_ & " - " & strProcedureName_)
            Call dicProcedureNameToCode.Add(strProcedureCode_ & " - " & strProcedureName_, _
                                            strProcedureCode_)
            'Call cmbProcedureCode.AddItem(strProcedureCode_ & " - " & strProcedureName_)
            
            
            
            'Call cmbProcedureCode.AddItem(strProcedureName_)
            
        Wend
        
        
    End If
    
    
    'sort a dictionary in favor of the procedure combobox order:
    Dim dicProceduresCombo As New Dictionary
    Dim h As New Heap
    
    '1. build a heap:
    For i = 0 To dicProcedureCodeToName.Count - 1
    
        Call h.Enter("", dicProcedureCodeToName.Keys(i))
    
    Next i
    
    '2. extract from the heap:
    While h.Leave("", strProcedureCode_)
    
        Call dicProceduresCombo.Add(strProcedureCode_, strProcedureCode_)
    
    Wend
    
    '3. fill the combo:
    For i = dicProceduresCombo.Count - 1 To 0 Step -1
    
        Call cmbProcedureCode.AddItem(dicProcedureCodeToName(dicProceduresCombo.Keys(i)))
    
    Next i
    
    
    
    
    If strProcedureCode <> "" And IsAutoSelection Then
        'cmbProcedureCode.Text = strProcedureCode
        cmbProcedureCode.Text = dicProcedureCodeToName(strProcedureCode)
    ElseIf dicProcedureCodeToName.Count = 1 Then
        'cmbProcedureCode.Text = dicProcedureCodeToName.Keys(0)
        cmbProcedureCode.Text = dicProcedureCodeToName.Items(0)
    End If
    
    
    'Call ReadBillingCode

    Exit Sub
ERR_ReadProcedures:
MsgBox "ERR_ReadProcedures" & vbCrLf & Err.Description
End Sub


''if there is only one choice (there is a chosen Procedure Code)
''write it in the Billing Code text box;
'Private Sub ReadBillingCode()
'On Error GoTo ERR_ReadBillingCode
'
'    Dim rs As Recordset
'    Dim sql As String
'    Dim dicProcedures As Dictionary
'    Dim dicBillingCodes As Dictionary
'    Dim i As Integer
'
'    txtBillingCode.Text = ""
'
'    If cmbOrgan.Text = "" Or _
'       cmbTopography.Text = "" Or _
'       cmbProcedureCode.Text = "" Then
'
'        Exit Sub
'
'    End If
'
'    sql = " select otu.U_PROCEDURES, otu.U_BILLING_CODES"
'    sql = sql & " from lims_sys.u_norgan_topography_user otu "
'    sql = sql & " where otu.U_ORGAN_CODE='" & dicOrganNameToCode(cmbOrgan.Text) & "'"
'    sql = sql & " and   otu.U_TOPOGRAPHY_CODE='" & dicTopographyNameToCode(cmbTopography.Text) & "'"
'
'    Set rs = con.Execute(sql)
'
'    If rs.EOF Then
'        Exit Sub
'    End If
'
'    Set dicProcedures = StringToDictionary(nte(rs("U_PROCEDURES")), ",")
'    Set dicBillingCodes = StringToDictionary(nte(rs("U_BILLING_CODES")), ",")
'
'    'location of the selected procedure in peocedures collection;
'    'the location of the needed item in the billing codes collection is the same:
'    i = dicProcedures(cmbProcedureCode.Text)
'
'
''    If IsAutoSelection And strBillingCode <> "" Then
''
''        txtBillingCode.Text = strBillingCode
''
''    Else
''
''        If dicBillingCodes.Count > i Then
''
''            txtBillingCode.Text = dicBillingCodes.Keys(i)
''
''        End If
''
''    End If
'
'
'    If dicBillingCodes.Count > i Then
'
'        txtBillingCode.Text = dicBillingCodes.Keys(i)
'        txtBillingCode.Enabled = False
'
'        'there is no billing code to that
'        'procedore code selection:
'        If txtBillingCode.Text = "" Then
'
'            txtBillingCode.Enabled = True
'
'            If IsAutoSelection Then
'
'                txtBillingCode.Text = strBillingCode
'
'            End If
'
'        End If
'
'    End If
'
'    Exit Sub
'ERR_ReadBillingCode:
'MsgBox "ERR_ReadBillingCode" & vbCrLf & Err.Description
'End Sub


Public Function GetProcedureName(strProcedureCode As String) As String

    Dim rs As Recordset
    Dim sql As String
    
    sql = " select pu.U_PROCEDURE_NAME"
    sql = sql & " from lims_sys.u_nprocedure_user pu"
    sql = sql & " where pu.U_PROCEDURE_CODE='" & strProcedureCode & "'"
    
    Set rs = con.Execute(sql)

    If Not rs.EOF Then
        GetProcedureName = nte(rs("U_PROCEDURE_NAME"))
'    Else
'        GetProcedureName = "קוד לא קיים"
    End If

End Function


Private Sub ReadRemarkNeeded()

    Dim rs As Recordset
    Dim sql As String
    
    txtRemark.Text = ""
    IsRemarkNeeded = False
    PicRemark.Visible = False
    
    sql = " select tu.U_REMARK_NEEDED"
    sql = sql & " from lims_sys.u_ntopography_user tu"
    sql = sql & " where tu.U_TOPOGRAPHY_CODE='" & dicTopographyNameToCode(cmbTopography.Text) & "'"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
    
        If strRemark <> "" And IsAutoSelection Then
                
            txtRemark.Text = strRemark
                
        End If
            
        If nte(rs("U_REMARK_NEEDED")) = "T" Then
        
            PicRemark.Visible = True
            IsRemarkNeeded = True
            
        End If
    
    End If

End Sub



Private Sub InitChosenSampleData(strSampleId As String)
On Error GoTo ERR_InitChosenSampleData

    Dim rs As Recordset
    Dim sql As String
    
    'Ashi
    ClearMemoryFields
    lblSampleType.Caption = ""
        If strSampleId = "0" Then
            Exit Sub
        End If
'--------------

    sql = " select su.U_ORGAN_CODE, "
    sql = sql & "  su.U_ORGAN_SIDE, "
    sql = sql & "  su.U_TOPOGRAPHY_CODE, "
    sql = sql & "  su.U_ORGAN_REMARK, "
    sql = sql & "  su.U_PROCEDURE_CODE, "
    'sql = sql & "  su.U_BILLING_CODE, "
    sql = sql & "  s.STATUS, "
    sql = sql & "  s.SAMPLE_TYPE "
    sql = sql & " from  lims_sys.sample s, "
    sql = sql & "       lims_sys.sample_user su "
    sql = sql & " where su.SAMPLE_ID='" & strSampleId & "'"
    sql = sql & " and   su.SAMPLE_ID=s.SAMPLE_ID  "

    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
    
        strOrgan = (nte(rs("U_ORGAN_CODE")))
        strTopography = (nte(rs("U_TOPOGRAPHY_CODE")))
        strSide = (nte(rs("U_ORGAN_SIDE")))
        strRemark = nte(rs("U_ORGAN_REMARK"))
        'txtRemark.Text = strRemark
        strProcedureCode = nte(rs("U_PROCEDURE_CODE"))
        'strBillingCode = nte(rs("U_BILLING_CODE"))
        strSampleStatus = nte(rs("STATUS"))
        
        
        lblSampleType.Caption = cmbSample.Text
        'lblSampleType.Caption = nte(rs("SAMPLE_TYPE"))
    
    End If
    
    Call ComputeStatus

    Exit Sub
ERR_InitChosenSampleData:
MsgBox "ERR_InitChosenSampleData" & vbCrLf & Err.Description
End Sub


Private Sub ComputeStatus()
On Error GoTo ERR_ComputeStatus

    Dim rs As Recordset
    Dim sql As String
    Dim sql2 As String

    If IsSampleMode Then
    
        sql = sql & "select 1 "
        sql = sql & " from lims_sys.sample_user su"
        sql = sql & " where su.SAMPLE_ID='" & strSampleId & "'"
        sql = sql & " and su.U_PROCEDURE_CODE is null "
        
     'checks if there is a "?" in topo code or side code
        sql2 = sql2 & "select 2 "
        sql2 = sql2 & " from lims_sys.sample_user su"
        sql2 = sql2 & " where su.SAMPLE_ID='" & strSampleId & "'"
        sql2 = sql2 & "and  (su.U_TOPOGRAPHY_CODE='?' or su.U_ORGAN_SIDE= '?') "
        
    
    Else
    
        Dim strSampleIds As String
        strSampleIds = DictionaryToString(dicSamplesIdToType, ",")
    
        sql = " select 1"
        sql = sql & " from lims_sys.sample s,"
        sql = sql & "      lims_sys.sample_user su"
        sql = sql & " where su.SAMPLE_ID=s.SAMPLE_ID"
        sql = sql & " and   su.U_PROCEDURE_CODE is null"
        sql = sql & " and   s.SAMPLE_ID in (" & strSampleIds & ") "
        
    
    
   'checks if there is a "?" in topo code or side code
        sql2 = " select 2"
        sql2 = sql2 & " from lims_sys.sample s,"
        sql2 = sql2 & "      lims_sys.sample_user su"
        sql2 = sql2 & " where su.SAMPLE_ID=s.SAMPLE_ID"
        sql2 = sql2 & " and   s.SAMPLE_ID in (" & strSampleIds & ") "
        sql2 = sql2 & "and  (su.U_TOPOGRAPHY_CODE='?' or su.U_ORGAN_SIDE= '?') "
    
'        sql = " select 1"
'        sql = sql & " from lims_sys.sample s,"
'        sql = sql & "      lims_sys.sample_user su"
'        sql = sql & " where su.SAMPLE_ID=s.SAMPLE_ID"
'        sql = sql & " and   s.STATUS not in (" & INVALID_STATUS_FOR_SELECTION & ") "
'        sql = sql & " and   su.U_BILLING_CODE is null"
'        sql = sql & " and   s.SDG_ID='" & strSdgId & "' "

    End If

    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        strStatus = STATUS_INCOMPLETE
    Else
        strStatus = STATUS_COMPLETE
    End If
    
     Set rs = con.Execute(sql2)
     If Not rs.EOF Then
        strStatus = "?"
    End If
    
    Exit Sub
ERR_ComputeStatus:
MsgBox "ERR_ComputeStatus" & vbCrLf & Err.Description
End Sub


Private Function GetSdgId(strSampleId As String) As String
    
    Dim rs As Recordset
    Dim sql As String
    
    sql = " select s.SDG_ID"
    sql = sql & " from lims_sys.sample s"
    sql = sql & " where s.SAMPLE_ID='" & strSampleId & "'"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        GetSdgId = nte(rs("SDG_ID"))
    End If
    
End Function


Private Function GetFirstSampleId(strSdgId As String) As String

    Dim rs As Recordset
    Dim sql As String
    
    sql = " select s.SAMPLE_ID"
    sql = sql & " from lims_sys.sample s"
    sql = sql & " where s.SDG_ID = '" & strSdgId & "'"
    sql = sql & " order by s.SAMPLE_ID"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        GetFirstSampleId = nte(rs("SAMPLE_ID"))
    End If

End Function


Private Function GetSdgName(strSdgId As String)

    Dim rs As Recordset
    Dim sql As String
    
    sql = " select d.NAME"
    sql = sql & " from lims_sys.sdg d"
    sql = sql & " where d.SDG_ID = '" & strSdgId & "'"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        GetSdgName = nte(rs("NAME"))
    End If

End Function


Private Function nte(e As Variant) As String
    
    nte = IIf(IsNull(e), "", e)
    
End Function

Private Sub cmbOrgan_Change()
    Call ReadTopography
End Sub

Private Sub cmbOrgan_Click()
    Call ReadTopography
End Sub

Private Sub cmbOrgan_GotFocus()
    Call zLang.Hebrew
End Sub

Private Sub cmbOrgan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub

Private Sub cmbProcedureCode_Change()
    'Call ReadBillingCode
End Sub

Private Sub cmbProcedureCode_Click()
    'Call ReadBillingCode
End Sub

Private Sub cmbProcedureCode_GotFocus()
    Call zLang.English
End Sub

Private Sub cmbProcedureCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub

Private Sub cmbProcedureCode_KeyPress(KeyAscii As Integer)
 Dim CB As Long
    Dim FindString As String
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    If cmbProcedureCode.SelLength = 0 Then
        FindString = cmbProcedureCode.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(cmbProcedureCode.Text, cmbProcedureCode.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage cmbProcedureCode.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(cmbProcedureCode.hWnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        cmbProcedureCode.ListIndex = CB
        cmbProcedureCode.SelStart = Len(FindString)
        cmbProcedureCode.SelLength = Len(cmbProcedureCode.Text) - cmbProcedureCode.SelStart
    End If
    
    KeyAscii = 0
End Sub

Private Sub cmbSample_Click()
    strSampleId = dicSamplesTypeToId(cmbSample.Text)
    Call ClearMemoryFields
    Call ResetScreenBySample
End Sub

Private Sub cmbSample_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub

Private Sub cmbSide_Change()
'    If cmbSide.Text = "?" Then
'         cmbTopography.enabled = false
'    Else
'          cmbTopography.enabled = True
''        Call ReadProcedures
''        Call ReadRemarkNeeded
'    End If
End Sub

Private Sub cmbSide_GotFocus()
    Call zLang.Hebrew
End Sub

Private Sub cmbSide_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub

Private Sub TopographyChange()
    'If strSide = "?" Then strSide = ""
    Call ReadSide
    Call ReadProcedures
    Call ReadRemarkNeeded
End Sub

Private Sub cmbTopography_Change()
    Call TopographyChange
End Sub

Private Sub cmbTopography_Click()
    Call TopographyChange
End Sub

Private Sub cmbTopography_GotFocus()
    Call zLang.Hebrew
End Sub

Private Sub cmbTopography_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ERR_cmdCancel_Click

'    Call Unload(Me)
    Me.Hide

    Exit Sub
ERR_cmdCancel_Click:
MsgBox "ERR_cmdCancel_Click" & vbCrLf & Err.Description
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub

Private Sub cmdOK_Click()
On Error GoTo ERR_cmdOK_Click

   If strSampleId = "0" Then
    strOrgan = dicOrganNameToCode(cmbOrgan.Text)
    strSide = IIf(cmbSide.Text = "?", "?", dicSideNameToCode(cmbSide.Text))
    strTopography = IIf(cmbTopography.Text = "?", "?", dicTopographyNameToCode(cmbTopography.Text))
    strRemark = txtRemark.Text
    'strProcedureCode = cmbProcedureCode.Text
    strProcedureCode = IIf(cmbProcedureCode.Text = "", "", dicProcedureNameToCode(cmbProcedureCode.Text))
    'strBillingCode = txtBillingCode.Text
        Call Me.Hide
        Exit Sub
     End If
     
    If Not IsInputValidForApproval Then
        Exit Sub
    End If
    
    Call Update
    Call ComputeStatus
'    okClicked = True
    strStatus = strStatus & "+"
    
    Call sdg_log.InsertLog(CLng(strSdgId), "ORGAN.UPD", "")
    
    Call Me.Hide
    
   

    Exit Sub
ERR_cmdOK_Click:
MsgBox "ERR_cmdOK_Click" & vbCrLf & Err.Description
End Sub


Private Function IsInputValidForApproval() As Boolean
On Error GoTo ERR_IsInputValidForApproval

    IsInputValidForApproval = False
    
    If InStr(1, INVALID_STATUS_FOR_APPROVAL, strSampleStatus) Then
        MsgBox " סטאטוס הדגימה איננו חוקי עבור אישור "
        Exit Function
    End If


    If cmbOrgan.Text = "" Or dicOrganNameToCode.Exists(cmbOrgan.Text) = False Then
        MsgBox " איבר חסר "
        Call cmbOrgan.SetFocus
        Exit Function
    End If
    
     If dicOrganNameToCode(cmbOrgan.Text) = "" Then
        MsgBox " איבר חסר "
        Call cmbOrgan.SetFocus
        Exit Function
    End If
    
'     If cmbTopography.Text = "" Or _
'    (dicTopographyNameToCode.Exists(cmbTopography.Text) And cmbTopography.Text <> "?") = False Then

    If cmbTopography.Text = "" Or _
    (dicTopographyNameToCode.Exists(cmbTopography.Text)) = False Then
        
        MsgBox " מיקום חסר "
        Call cmbTopography.SetFocus
        Exit Function
    End If
    
'      If dicTopographyNameToCode(cmbTopography.Text) = "" And cmbTopography.Text <> "?" Then
     If dicTopographyNameToCode(cmbTopography.Text) = "" Then
        MsgBox " מיקום חסר "
        Call cmbTopography.SetFocus
        Exit Function
    End If
'     If cmbSide.Text = "" Or (dicSideNameToCode.Exists(cmbSide.Text) = False _
'                            And cmbSide.Text <> "?") Then
    If cmbSide.Text = "" Or (dicSideNameToCode.Exists(cmbSide.Text) = False _
                            ) Then
        MsgBox " צד חסר "
        Call cmbSide.SetFocus
        Exit Function
    End If
    
'     If dicSideNameToCode(cmbSide.Text) = False And cmbSide.Text <> "?" Then
    
     If dicSideNameToCode(cmbSide.Text) = False Then
        MsgBox " צד חסר "
        Call cmbSide.SetFocus
        Exit Function
    End If

    If IsRemarkNeeded And txtRemark.Text = "" Then
        MsgBox " הערה חסרה "
        Call txtRemark.SetFocus
        Exit Function
    End If

    'If cmbProcedureCode.Text = "" Or dicProcedureCodeToName.Exists(cmbProcedureCode.Text) = False Then
    If (cmbProcedureCode.Text = "" Or dicProcedureNameToCode.Exists(cmbProcedureCode.Text) = False) _
         And cmbSide.Text <> "?" And cmbTopography.Text <> "?" Then
        MsgBox " קוד טיפול חסר "
        Call cmbProcedureCode.SetFocus
        Exit Function
    End If
    
'    If txtBillingCode.Text = "" Then
'        MsgBox " קוד חיוב חסר "
'        Call txtBillingCode.SetFocus
'        Exit Function
'    End If

    IsInputValidForApproval = True

    Exit Function
ERR_IsInputValidForApproval:
MsgBox "ERR_IsInputValidForApproval" & vbCrLf & Err.Description
End Function


Private Sub Update()
On Error GoTo ERR_Update

    Dim rs As Recordset
    Dim sql As String

    strOrgan = dicOrganNameToCode(cmbOrgan.Text)
    strSide = IIf(cmbSide.Text = "?", "?", dicSideNameToCode(cmbSide.Text))
    strTopography = IIf(cmbTopography.Text = "?", "?", dicTopographyNameToCode(cmbTopography.Text))
    strRemark = txtRemark.Text
    'strProcedureCode = cmbProcedureCode.Text
    strProcedureCode = IIf(cmbProcedureCode.Text = "", "", dicProcedureNameToCode(cmbProcedureCode.Text))
    'strBillingCode = txtBillingCode.Text
    
    sql = " update lims_sys.sample_user su"
    sql = sql & " set su.U_ORGAN_CODE       = '" & DoubleApostrphe(strOrgan) & "',"
    sql = sql & "     su.U_ORGAN_SIDE       = '" & DoubleApostrphe(strSide) & "',"
    sql = sql & "     su.U_TOPOGRAPHY_CODE  = '" & DoubleApostrphe(strTopography) & "',"
    sql = sql & "     su.U_ORGAN_REMARK     = '" & DoubleApostrphe(strRemark) & "',"
    sql = sql & "     su.U_PROCEDURE_CODE   = '" & DoubleApostrphe(strProcedureCode) & "',"
    sql = sql & "     su.U_ORGAN            = '" & DoubleApostrphe(GetOrganName) & "',"
    sql = sql & "     su.U_TOPOGRAPHY       = '" & DoubleApostrphe(GetTopographyName) & "'"
    'sql = sql & "     su.U_BILLING_CODE     = '" & strBillingCode & "'"
    sql = sql & " where su.SAMPLE_ID        = '" & DoubleApostrphe(strSampleId) & "' "
 
   
    Call con.Execute(sql)

    Exit Sub
ERR_Update:
MsgBox "ERR_Update" & vbCrLf & Err.Description
End Sub

Private Function DoubleApostrphe(e As Variant) As String
'double the apostrophes in astring
    Dim temp As String
     temp = IIf(IsNull(e), "", e)
    If InStr(1, temp, "'") > 0 Then
        temp = Replace(temp, "'", "''")
    End If
    DoubleApostrphe = temp
End Function

Public Function GetOrganName() As String
    
    Dim rs As Recordset
    Dim sql As String
    
    sql = " select ou.U_ORGAN_HEBREW_NAME"
    sql = sql & " from lims_sys.u_norgan_user ou"
    sql = sql & " where ou.U_ORGAN_CODE='" & strOrgan & "'"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        GetOrganName = nte(rs("U_ORGAN_HEBREW_NAME"))
    End If
    
End Function

Public Function GetSideName() As String
    
    Dim rs As Recordset
    Dim sql As String
    
    sql = " select sd.DESCRIPTION"
    sql = sql & " from lims_sys.u_nside sd"
    sql = sql & " where sd.NAME='" & strSide & "'"
    
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
        GetSideName = nte(rs("DESCRIPTION"))
    End If
    
End Function



'return the remark if mandatory (describes the topography):
Public Function GetTopographyName() As String
    
    Dim rs As Recordset
    Dim sql As String
        
    sql = " select tu.U_TOPOGRAPHY_HEBREW_NAME, tu.U_REMARK_NEEDED "
    sql = sql & " from lims_sys.u_ntopography_user tu"
    sql = sql & " where tu.U_TOPOGRAPHY_CODE='" & strTopography & "'"
        
    Set rs = con.Execute(sql)
    
    If Not rs.EOF Then
    
'        If nte(rs("U_REMARK_NEEDED")) = "T" Then
'            GetTopographyName = strRemark
'        Else
        GetTopographyName = nte(rs("U_TOPOGRAPHY_HEBREW_NAME")) & " " & strRemark
'        End If
        
    End If
    
End Function

Public Function GetTopographyCode() As String

    GetTopographyCode = strTopography
    
End Function
Public Function GetSampleId() As String

    GetSampleId = strSampleId

End Function
Public Function GetOrganCode() As String

    GetOrganCode = strOrgan

End Function


Public Function GetProcedureCode() As String
    
    GetProcedureCode = strProcedureCode
    
End Function

Public Function GetProcedure() As String

    GetProcedure = GetProcedureName(strProcedureCode)

End Function

Public Function GetBillingCode() As String

    'GetBillingCode = strBillingCode

End Function


Public Function GetRemark() As String

    GetRemark = strRemark

End Function

Public Function GetStatus() As String
    GetStatus = strStatus
End Function


'compute and return Snomed-T codes for the SDG:
Public Function GetSnomedT() As String
On Error GoTo ERR_GetSnomedT

    Dim rs As Recordset
    Dim sql As String
    Dim dAllSides As New Dictionary
    Dim dicSnomedT As New Dictionary
    Dim strSnomedT As String
    Dim i As Integer
    
    Call dAllSides.RemoveAll
    
    'read all side codes:
    sql = " select sd.NAME, sd.DESCRIPTION"
    sql = sql & " from lims_sys.u_nside sd"
    sql = sql & " order by sd.U_NSIDE_ID"
        
    Set rs = con.Execute(sql)
    
    While Not rs.EOF
    
        Call dAllSides.Add(nte(rs("NAME")), dAllSides.Count)
        rs.MoveNext
        
    Wend

    
    'get possible snomed codes for the
    'organ & topography selections of each sample:
    sql = " select tu.U_R_SNOMED, tu.U_L_SNOMED, tu.U_O_SNOMED, TU.U_RL_SNOMED ,tu.U_NS_SNOMED, su.U_ORGAN_SIDE "
    sql = sql & " from lims_sys.u_norgan_topography_user tu,"
    sql = sql & "      lims_sys.sample s,"
    sql = sql & "      lims_sys.sample_user su "
    sql = sql & " where tu.U_ORGAN_CODE      = su.U_ORGAN_CODE"
    sql = sql & " and   tu.U_TOPOGRAPHY_CODE = su.U_TOPOGRAPHY_CODE"
    sql = sql & " and   su.SAMPLE_ID         = s.SAMPLE_ID"
    sql = sql & " and   s.SDG_ID             = '" & strSdgId & "'"
    sql = sql & " and   su.U_ORGAN_SIDE is not null "

    Set rs = con.Execute(sql)
    
    
    'for each sample, get the relevant snomed field
    'according to the organ-side selection;
    'side-code => number of snomed field => snomed value;
    While Not rs.EOF
    
        Dim strSampleSnomedT As String
        Dim strSide As String
        Dim nSnomedTNumber As Integer
        
        strSide = nte(rs("U_ORGAN_SIDE"))
        nSnomedTNumber = dAllSides(strSide)
        strSampleSnomedT = nte(rs(nSnomedTNumber))
    
        strSnomedT = strSnomedT & strSampleSnomedT & ","
    
        rs.MoveNext
    
    Wend
     
     
    'create a colection of unique snomed values:
    While strSnomedT <> ""
        
        Dim strSingleSnomed As String
        
        strSingleSnomed = Trim(getNextStr(strSnomedT, ","))
        
        If Not dicSnomedT.Exists(strSingleSnomed) Then
            Call dicSnomedT.Add(strSingleSnomed, strSingleSnomed)
        End If
        
    Wend
     
     
    'create a return string of snomed values:
    strSnomedT = ""
    
    For i = 0 To dicSnomedT.Count - 1
    
         strSnomedT = "," & dicSnomedT.Keys(i) & strSnomedT
    
    Next i
    
    strSnomedT = Mid(strSnomedT, 2)
     
    GetSnomedT = strSnomedT

    Exit Function
ERR_GetSnomedT:
MsgBox "ERR_GetSnomedT" & vbCrLf & Err.Description
End Function


Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call Me.Hide
    End If
End Sub

Private Sub Form_Load()

    Call EnableCloseButton(Me.hWnd, False)
    'Call cmbOrgan.SetFocus
    
    
        
    'MsgBox strSampleId & vbCrLf & strSdgId & vbCrLf & dicSamplesIdToType.Count
End Sub


Private Sub Form_Paint()
On Error GoTo ERR_Form_Paint

    'each time the form is re presented
    'set the focus for the organ list:
    Call cmbOrgan.SetFocus


    Exit Sub
ERR_Form_Paint:
MsgBox "ERR_Form_Paint" & vbCrLf & Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    If IsSampleMode Then
'        strSdgId = ""
'    Else
'        strSampleId = ""
'    End If
End Sub





'Private Sub txtBillingCode_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call KeyDown(KeyCode)
'End Sub

Private Sub txtRemark_GotFocus()
    Call zLang.English
End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode)
End Sub


Private Function DictionaryToString(d As Dictionary, strDelimiter As String) As String
On Error GoTo ERR_DictionaryToString

    Dim i As Integer
    Dim str As String
    
    For i = 0 To d.Count - 1
        str = str & d.Keys(i) & strDelimiter
    Next i

    If str <> "" Then
        str = Mid(str, 1, Len(str) - 1)
    End If
    
    DictionaryToString = str

    Exit Function
ERR_DictionaryToString:
MsgBox "ERR_DictionaryToString" & vbCrLf & Err.Description
End Function

Private Function StringToDictionary(ByVal str As String, strDelimiter As String) As Dictionary
On Error GoTo ERR_StringToDictionary

    Dim d As New Dictionary
    Dim s As String
    
    While str <> ""
        s = Trim(getNextStr(str, strDelimiter))
        
        If Not d.Exists(s) Then
            Call d.Add(s, d.Count)
        End If
    Wend
    
    Set StringToDictionary = d
    
    Exit Function
ERR_StringToDictionary:
MsgBox "ERR_StringToDictionary" & vbCrLf & Err.Description
End Function


Private Function getNextStr(ByRef s As String, c As String)
    Dim p
    Dim res
    p = InStr(1, s, c)
    If (p = 0) Then
        res = s
        s = ""
        getNextStr = res
    Else
        res = Mid$(s, 1, p - 1)
        s = Mid$(s, p + Len(c), Len(s))
        getNextStr = res
    End If
End Function

Private Sub KeyDown(KeyCode As Integer)
 On Error GoTo ERR_KeyDown

    If KeyCode = vbKeyEscape Then
        Call Me.Hide
    End If

    Exit Sub
ERR_KeyDown:
MsgBox "ERR_KeyDown" & vbCrLf & Err.Description
End Sub



Public Function GetOrgans() As Dictionary
On Error GoTo ERR_GetOrgans
        
    Dim rs As Recordset
    Dim sql As String
    Dim Dic As New Dictionary
    
    If strSampleId <> "" Then
        strSdgId = GetSdgId(strSampleId)
    End If
 
    sql = " select ou.U_ORGAN_HEBREW_NAME"
    sql = sql & " from lims_sys.u_norgan_user ou"
    sql = sql & " where ou.U_ORGAN_CODE='" & strOrgan & "'"
 
 
    sql = " select s.SAMPLE_ID, ou.U_ORGAN_HEBREW_NAME"
    sql = sql & " from lims_sys.sample s, lims_sys.u_norgan_user ou, lims_sys.sample_user su"
    sql = sql & " where s.SDG_ID = '" & strSdgId & "'"
    sql = sql & " and s.status not in (" & INVALID_STATUS_FOR_SELECTION & ")  "
    sql = sql & " and ou.U_ORGAN_CODE = su.u_organ_code"
    sql = sql & " and su.sample_id = s.sample_id"
    sql = sql & " order by s.SAMPLE_ID"
    
        
    Set rs = con.Execute(sql)
    While Not rs.EOF
    
        Dic.Add nte(rs("SAMPLE_ID")), nte(rs("U_ORGAN_HEBREW_NAME"))
        rs.MoveNext
        
    Wend
    
    Set GetOrgans = Dic

    Exit Function
ERR_GetOrgans:
MsgBox "ERR_GetOrgans" & vbCrLf & Err.Description

End Function
