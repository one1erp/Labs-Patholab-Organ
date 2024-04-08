VERSION 5.00
Begin VB.UserControl OrganCtrl 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   1215
   Begin VB.CommandButton cmdOpenForm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "איבר"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "OrganCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click()
Public Event StatusChanged(strStatus As String)


Dim frmOrgan_ As New frmOrgan


Private Sub cmdOpenForm_Click()
    'Call frmOrgan.Initialize
    If frmOrgan_.strSampleStatus = "A" Then MsgBox "Request is authorised, no changes can be made."
    Call frmOrgan_.ResetScreenBySample  'show sample updated data,
                                       'not last selection that was not saved
    frmOrgan_.Show (vbModal)
    
    Call SetColor
    RaiseEvent Click
    RaiseEvent StatusChanged(status)
End Sub


'initialize the compunent after the
'relevant input was given:
Public Sub Initialize()
    cmdOpenForm.Font.Charset = 177
    Call frmOrgan_.Initialize
    cmdOpenForm.ToolTipText = " Organ : " & frmOrgan_.GetOrganName & _
                              ", Topography : " & frmOrgan_.GetTopographyName & _
                              ", Side : " & frmOrgan_.GetSideName & _
                              ", Procedure : " & frmOrgan_.GetProcedure
                             
    Call SetColor
    RaiseEvent StatusChanged(status)
End Sub


Private Sub SetColor()
Dim status As String
    status = frmOrgan_.GetStatus
    If Left(status, 1) = "C" Then
        cmdOpenForm.BackColor = &HC0FFC0
    ElseIf Left(status, 1) = "?" Then
         cmdOpenForm.BackColor = &H80FF&
        ' cmdOpenForm.BackColor = &H80FF &H000080FF& &H0080C0FF&
    Else
        cmdOpenForm.BackColor = &H8000000F
    End If
End Sub

Public Property Let Connection(con As Connection)
    Set frmOrgan_.con = con
    Set frmOrgan_.sdg_log.con = con
End Property

Public Property Let SessionId(dSessionId As Double)
    frmOrgan_.sdg_log.Session = dSessionId
End Property


Public Property Let SdgId(strSdgId As String)
    frmOrgan_.strSdgId = strSdgId
    frmOrgan_.strSampleId = ""
End Property

Public Property Let SampleId(strSampleId As String)
    frmOrgan_.strSampleId = strSampleId
    frmOrgan_.strSdgId = ""
End Property

Public Property Let OperatorName(strOperatorName As String)
    frmOrgan_.strOperatorName = strOperatorName
End Property

Public Property Let Enabled(IsEnabled As Boolean)
    cmdOpenForm.Enabled = IsEnabled
End Property

Public Property Let FontName(strFontName As String)
    cmdOpenForm.Font.Name = strFontName
End Property

Public Property Let FontSize(nFontSize As Integer)
    cmdOpenForm.Font.Size = nFontSize
End Property

Public Property Let Font(fnt As Font)
    Set cmdOpenForm.Font = fnt
End Property

Public Property Let Width(dWidth As Double)
    
    UserControl.Width = dWidth
    'cmdOpenForm.Width = dWidth
    
End Property

Public Property Let Height(dHeight As Double)
    
    UserControl.Height = dHeight
    'cmdOpenForm.Height = dHeight
    
End Property
Public Property Get SampleId() As String
    SampleId = frmOrgan_.GetSampleId
End Property

Public Property Get OrganCode() As String
    OrganCode = frmOrgan_.GetOrganCode
End Property


Public Property Get Organ() As String
    Organ = frmOrgan_.GetOrganName
End Property

Public Property Get Side() As String
    Side = frmOrgan_.GetSideName
End Property

Public Property Get Topography() As String
    Topography = frmOrgan_.GetTopographyName
End Property

Public Property Get TopographyCode() As String
    TopographyCode = frmOrgan_.GetTopographyCode
End Property


Public Property Get Remark() As String
    Remark = frmOrgan_.GetRemark
End Property

Public Property Get ProcedureCode() As String
    ProcedureCode = frmOrgan_.GetProcedureCode

End Property

Public Property Get ProcedureName() As String
    ProcedureName = frmOrgan_.GetProcedure
End Property

Public Property Get BillingCode() As String
    BillingCode = frmOrgan_.GetBillingCode
End Property

Public Property Get status() As String
    status = frmOrgan_.GetStatus
End Property

Public Property Get SnomedT() As String
   'ASHI 5/8/20 ' SnomedT = frmOrgan_.GetSnomedT
End Property

Private Sub UserControl_Resize()
    cmdOpenForm.Width = UserControl.Width
    cmdOpenForm.Height = UserControl.Height
End Sub

Public Function Organs() As Dictionary
    Set Organs = frmOrgan_.GetOrgans
End Function
