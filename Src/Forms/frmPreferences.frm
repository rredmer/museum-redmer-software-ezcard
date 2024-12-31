VERSION 5.00
Begin VB.Form frmPreferences 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRemovable 
      BackColor       =   &H00C0E0FF&
      Caption         =   "PhotoCard or other removable storage in use for batches."
      Height          =   240
      Left            =   75
      TabIndex        =   15
      Top             =   690
      Width           =   4515
   End
   Begin VB.TextBox txtDegrees 
      Height          =   315
      Left            =   1185
      TabIndex        =   2
      Text            =   "90"
      Top             =   375
      Width           =   435
   End
   Begin VB.ComboBox cboReports 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   5250
   End
   Begin VB.CheckBox chkMoveImages 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Move Images"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   975
      Width           =   1290
   End
   Begin VB.Frame fraTo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Move Images To"
      Height          =   3375
      Left            =   3345
      TabIndex        =   6
      Top             =   1245
      Width           =   3075
      Begin VB.DriveListBox drvTo 
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   2925
      End
      Begin VB.DirListBox dirTo 
         Height          =   1215
         Left            =   60
         TabIndex        =   12
         Top             =   570
         Width           =   2925
      End
      Begin VB.FileListBox filTo 
         Height          =   1455
         Left            =   60
         TabIndex        =   11
         Top             =   1800
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdMain 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   705
      Index           =   1
      Left            =   735
      Picture         =   "frmPreferences.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4695
      Width           =   675
   End
   Begin VB.CommandButton cmdMain 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      Height          =   705
      Index           =   0
      Left            =   60
      Picture         =   "frmPreferences.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4695
      Width           =   675
   End
   Begin VB.Frame fraFrom 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Original Images Are Stored In"
      Height          =   3375
      Left            =   255
      TabIndex        =   5
      Top             =   1245
      Width           =   3075
      Begin VB.FileListBox filFrom 
         Height          =   1455
         Left            =   60
         TabIndex        =   10
         Top             =   1800
         Width           =   2925
      End
      Begin VB.DirListBox dirFrom 
         Height          =   1215
         Left            =   60
         TabIndex        =   9
         Top             =   570
         Width           =   2925
      End
      Begin VB.DriveListBox drvFrom 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   2925
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rotate Images"
      Height          =   240
      Left            =   75
      TabIndex        =   14
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Card Report"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   1110
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CON_KEY_APPTAG = "RSC_EZCARD"
Private Const CON_KEY_BRANCH = "Preferences"
Private Const CON_KEY_SOURCE = "SOURCE_PATH"
Private Const CON_KEY_TARGET = "TARGET_PATH"
Private Const CON_KEY_REPORT = "ID_REPORT"
Private Const CON_KEY_MOVE = "MOVE_IMAGES"
Private Const CON_KEY_ROTATE = "ROTATION_ANGLE"
Private Const CON_KEY_REMOVABLE = "REMOVABLE_MEDIA"

Private Sub Form_Load()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim s As String
    Dim i As Integer
    Dim oAccess As Access.Application
    Dim obj As AccessObject, dbs As Object
    Set oAccess = New Access.Application
    oAccess.OpenCurrentDatabase App.Path & "\EZCARDS.MDB", False
    Set dbs = oAccess.CurrentProject
    For Each obj In dbs.AllReports
        cboReports.AddItem obj.Name
    Next obj
    oAccess.Quit acQuitSaveNone
    s = Me.SourcePath
    s = Me.TargetPath
    s = Me.ReportName
    i = Me.MoveImages
    i = Me.RotationAngle
    i = Me.RemovableMedia
    Set oAccess = Nothing
End Sub

Private Sub Form_Paint()
    If Not EZ_DEBUG Then On Error Resume Next
    If chkMoveImages.Value = 1 Then
        fraTo.Enabled = True
        fraTo.Visible = True
    Else
        fraTo.Enabled = False
        fraTo.Visible = False
    End If
End Sub

Private Sub chkMoveImages_Click()
    If Not EZ_DEBUG Then On Error Resume Next
    Me.Refresh
End Sub

Private Sub cmdMain_Click(Index As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    Dim iAnswer As Integer
    Select Case Index
        Case 0:             'Exit
            Me.Hide
        Case 1:             'Save
            SaveSetting CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_SOURCE, dirFrom.Path
            SaveSetting CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_TARGET, dirTo.Path
            SaveSetting CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_REPORT, Str$(cboReports.ListIndex)
            SaveSetting CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_MOVE, Str$(chkMoveImages.Value)
            SaveSetting CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_ROTATE, txtDegrees.Text
            SaveSetting CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_REMOVABLE, Str$(chkRemovable.Value)
            Me.Hide
    End Select
End Sub

Private Sub dirFrom_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    filFrom.Path = dirFrom.Path
    Me.Refresh
End Sub

Private Sub drvFrom_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    dirFrom.Path = drvFrom.Drive
    Me.Refresh
End Sub

Private Sub dirTo_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    filTo.Path = dirTo.Path
    Me.Refresh
End Sub

Private Sub drvTo_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    dirTo.Path = drvTo.Drive
    Me.Refresh
End Sub

Public Property Get SourcePath() As String
    If Not EZ_DEBUG Then On Error Resume Next
    Dim sFromPath As String
    sFromPath = GetSetting(CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_SOURCE)
    dirFrom.Path = sFromPath
    drvFrom.Drive = IIf(Mid$(sFromPath, 2, 1) = ":", Left$(sFromPath, 2), "C:")
    SourcePath = dirFrom.Path
End Property

Public Property Get TargetPath() As String
    If Not EZ_DEBUG Then On Error Resume Next
    Dim sToPath As String
    sToPath = GetSetting(CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_TARGET)
    dirTo.Path = sToPath
    drvTo.Drive = IIf(Mid$(sToPath, 2, 1) = ":", Left$(sToPath, 2), "C:")
    TargetPath = dirTo.Path
End Property

Public Property Get ReportName() As String
    If Not EZ_DEBUG Then On Error Resume Next
    Dim iReport As Integer
    iReport = Val(GetSetting(CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_REPORT))
    cboReports.ListIndex = iReport
    ReportName = cboReports.Text
End Property

Public Property Get MoveImages() As Integer
    If Not EZ_DEBUG Then On Error Resume Next
    Dim iMove As Integer
    iMove = Val(GetSetting(CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_MOVE))
    chkMoveImages.Value = iMove
    MoveImages = chkMoveImages.Value
End Property

Public Property Get RotationAngle() As Integer
    If Not EZ_DEBUG Then On Error Resume Next
    txtDegrees.Text = GetSetting(CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_ROTATE)
    RotationAngle = Val(txtDegrees.Text)
End Property

Public Property Get RemovableMedia() As Integer
    If Not EZ_DEBUG Then On Error Resume Next
    Dim iRemovable As Integer
    iRemovable = Val(GetSetting(CON_KEY_APPTAG, CON_KEY_BRANCH, CON_KEY_REMOVABLE))
    chkRemovable.Value = iRemovable
    RemovableMedia = chkRemovable.Value
End Property


