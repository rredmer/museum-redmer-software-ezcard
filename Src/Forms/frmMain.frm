VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00100003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltocx10N.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RSC EZ-CARD Version 1.10 R 001"
   ClientHeight    =   6555
   ClientLeft      =   2670
   ClientTop       =   1995
   ClientWidth     =   8895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkReprint 
      Caption         =   "Reprint Batch"
      Height          =   345
      Left            =   2175
      TabIndex        =   9
      Top             =   5685
      Width           =   1395
   End
   Begin VB.CommandButton cmdMain 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   705
      Index           =   2
      Left            =   1410
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5730
      Width           =   675
   End
   Begin VB.CommandButton cmdMain 
      Appearance      =   0  'Flat
      Caption         =   "&Batch"
      Height          =   705
      Index           =   1
      Left            =   735
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5730
      Width           =   675
   End
   Begin VB.CommandButton cmdMain 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      Height          =   705
      Index           =   0
      Left            =   60
      Picture         =   "frmMain.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5730
      Width           =   675
   End
   Begin VB.Frame fraStudent 
      Height          =   5715
      Left            =   15
      TabIndex        =   0
      Top             =   -90
      Width           =   8850
      Begin LEADLib.LEAD leadMain 
         Height          =   4545
         Left            =   5445
         TabIndex        =   10
         Top             =   1095
         Width           =   3345
         _Version        =   65537
         _ExtentX        =   5900
         _ExtentY        =   8017
         _StockProps     =   229
         BorderStyle     =   1
         ScaleHeight     =   301
         ScaleWidth      =   221
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
         PanWinTitle     =   "PanWindow"
         CLeadCtrl       =   0
      End
      Begin VB.TextBox txtSubjectID 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         TabIndex        =   4
         Top             =   660
         Width           =   1665
      End
      Begin MSDataGridLib.DataGrid grdMain 
         Height          =   4545
         Left            =   75
         TabIndex        =   3
         Top             =   1095
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exposures in Current Batch"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "sequence"
            Caption         =   "#"
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
         BeginProperty Column01 
            DataField       =   "subject"
            Caption         =   "subject"
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
         BeginProperty Column02 
            DataField       =   "first_name"
            Caption         =   "first_name"
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
         BeginProperty Column03 
            DataField       =   "last_name"
            Caption         =   "last_name"
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
               ColumnWidth     =   374.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2234.835
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtCompactFlashNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label lblSubjectID 
         Caption         =   "Subject ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblPictureCardNo 
         Caption         =   "Photo Batch#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuFile 
         Caption         =   "&New Batch"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Preferences"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsCurrent As ADODB.Recordset
Private iSequence As Integer
Private Const CON_QUERY = "SELECT SEQUENCE, SUBJECT, FIRST_NAME, LAST_NAME,* FROM ENDCUST"
Private Const CON_QUERY_EMPTY = "SELECT SEQUENCE, SUBJECT, FIRST_NAME, LAST_NAME,* FROM ENDCUST WHERE (false)"
Private Const CON_QUERY_FULL = "SELECT * FROM ENDCUST"

Private Sub Form_Load()
    If Not EZ_DEBUG Then On Error Resume Next
    Set rsCurrent = New ADODB.Recordset
    rsCurrent.Open CON_QUERY_EMPTY, gConn, adOpenDynamic, adLockOptimistic
    Set grdMain.DataSource = rsCurrent
    grdMain.Refresh
    txtCompactFlashNo.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    cmdMain_Click 0
End Sub

Private Sub grdMain_GotFocus()
    If Not EZ_DEBUG Then On Error Resume Next
    isBatchLoaded
End Sub

Private Sub mnuFile_Click(Index As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    Select Case Index
        Case 1
            cmdMain_Click 1
        Case 2
            cmdMain_Click 2
        Case 3
            frmPreferences.Show vbModal
        Case 5
            cmdMain_Click 0
    End Select
End Sub

Private Sub cmdMain_Click(Index As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    Dim iAnswer As Integer
    Select Case Index
        Case 0:
            If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo, "Quit Program") = vbYes Then
                rsCurrent.Close
                Set rsCurrent = Nothing
                gConn.Close
                Set gConn = Nothing
                End
            End If
        Case 1:
            txtCompactFlashNo.Enabled = True
            txtCompactFlashNo.SetFocus
        Case 2:
            PrintIDcards
    End Select
End Sub

Private Sub txtCompactFlashNo_GotFocus()
    If Not EZ_DEBUG Then On Error Resume Next
    If rsCurrent.State = adStateOpen Then
        rsCurrent.UpdateBatch adAffectAllChapters
        rsCurrent.Close
    End If
    rsCurrent.Open CON_QUERY_EMPTY, gConn, adOpenDynamic, adLockOptimistic
    If rsCurrent.RecordCount > 0 Then
        rsCurrent.MoveFirst
        iSequence = rsCurrent("SEQUENCE") + 1
    End If
    Set grdMain.DataSource = rsCurrent
    grdMain.Refresh
    cmdMain(1).Enabled = False
    iSequence = 1
    txtCompactFlashNo.Text = ""
    Me.Refresh
End Sub

Private Sub txtCompactFlashNo_LostFocus()
    If Not EZ_DEBUG Then On Error Resume Next
    If Len(Trim(txtCompactFlashNo.Text)) > 0 Then
        rsCurrent.Close
        rsCurrent.Open CON_QUERY + " WHERE (COMPACTFLASHNO='" & txtCompactFlashNo.Text & "') ORDER BY SEQUENCE DESC", gConn, adOpenDynamic, adLockOptimistic
        If rsCurrent.RecordCount > 0 Then
            rsCurrent.MoveFirst
            iSequence = rsCurrent("SEQUENCE") + 1
        End If
        rsCurrent.Requery
        grdMain.Refresh
        Me.Refresh
        txtCompactFlashNo.Enabled = False
        cmdMain(1).Enabled = True
        txtSubjectID.Enabled = True
        txtSubjectID.SetFocus
    End If
End Sub

Private Function isBatchLoaded() As Boolean
    If Not EZ_DEBUG Then On Error Resume Next
    If Len(Trim(txtCompactFlashNo.Text)) = 0 Then
        MsgBox "Please enter a batch number.", vbApplicationModal + vbInformation + vbOKOnly, "Missing batch number"
        txtCompactFlashNo.SetFocus
        isBatchLoaded = False
    Else
        isBatchLoaded = True
    End If
End Function

Private Sub txtSubjectID_GotFocus()
    If Not EZ_DEBUG Then On Error Resume Next
    isBatchLoaded
End Sub

Private Sub txtSubjectID_LostFocus()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim rsSubject As ADODB.Recordset
    
    '--- Validate the Subject id
    If Len(Trim$(txtSubjectID.Text)) = 0 Then
        Exit Sub
    End If
    
    Set rsSubject = New ADODB.Recordset
    rsSubject.Open CON_QUERY_FULL + " WHERE (SUBJECT='" & txtSubjectID.Text & "')", gConn, adOpenKeyset, adLockOptimistic
    If rsSubject.RecordCount > 0 Then
        'Display a message if subject is already in this batch, else change batch number.
        If rsSubject("COMPACTFLASHNO").Value = txtCompactFlashNo.Text Then
            MsgBox "Subject picture already taken on this PhotoCard", vbExclamation + vbOKOnly + vbApplicationModal, "Error."
        Else
            rsSubject("SEQUENCE").Value = iSequence
            rsSubject("COMPACTFLASHNO").Value = txtCompactFlashNo.Text
            rsSubject.Update
            iSequence = iSequence + 1
        End If
    Else
        MsgBox "Subject not found.  You may add this subject to the list below.", vbApplicationModal + vbInformation + vbOKOnly, "Not found."
    End If
    rsSubject.Close
    Set rsSubject = Nothing
    txtSubjectID.Text = ""
    txtSubjectID.SetFocus
End Sub

Private Sub grdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    LoadImage Trim$(frmPreferences.TargetPath) & "\" & Trim$(rsCurrent("SUBJECT").Value & "") & ".BMP"
End Sub

Private Sub grdMain_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    rsCurrent("CompactFlashNo").Value = txtCompactFlashNo.Text              'Update the batch indicator!
End Sub

Private Sub PrintIDcards()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim oAccess As Access.Application
    Dim iAnswer As Integer
  
    If Not isBatchLoaded Then Exit Sub
    iAnswer = MsgBox("Ready to print id cards for Photo Card #" & txtCompactFlashNo.Text & "?", vbQuestion + vbApplicationModal + vbYesNo, "Print Cards")
    If iAnswer = vbYes Then
        If (chkReprint.Value <> 1 And frmPreferences.MoveImages = 1) Then MoveImages
        Else: If MsgBox("Rotate Images?", vbQuestion + vbApplicationModal + vbYesNo, "Re-Print Cards") = vbYes Then RotateImages
        Set oAccess = New Access.Application
        oAccess.OpenCurrentDatabase App.Path & "\EZCARDS.MDB", False
        oAccess.DoCmd.OpenReport frmPreferences.ReportName, acViewPreview, , "[CompactFlashNo] = '" & Trim$(txtCompactFlashNo.Text) & "'"
        oAccess.DoCmd.PrintOut acSelection
        oAccess.DoCmd.Quit acQuitSaveAll
        Set oAccess = Nothing
    End If
    Exit Sub
End Sub

Private Sub MoveImages()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim rsBatch As ADODB.Recordset
    Dim iRecCnt As Integer                                  'Number of records in current set
    Dim sImgName As String
    Dim sOldName As String
    Dim sNewName As String
    
    MsgBox "Please remove Photo Card from the camera and place it into the reader.", vbApplicationModal + vbExclamation + vbOKOnly, "Insert Photo Card"

    '--- Loop through the records for the current report and retrieve the images
    Set rsBatch = New ADODB.Recordset
    gConn.Execute "DELETE * FROM FILEBATCH"
    rsBatch.Open "SELECT * FROM FILEBATCH ORDER BY FILENAME DESC", gConn, adOpenDynamic, adLockOptimistic
    iRecCnt = rsCurrent.RecordCount
    sImgName = Dir(frmPreferences.SourcePath & "\*.JPG", vbNormal)         'Retrieve the first entry.
    Do While sImgName <> ""                                 'Start the loop.
        If sImgName <> "." And sImgName <> ".." Then        'Ignore the current directory and the encompassing directory.
            rsBatch.AddNew
            rsBatch("FILENAME").Value = sImgName
            rsBatch.Update
        End If
        sImgName = Dir                                      'Get next entry.
    Loop
    rsBatch.MoveFirst
    rsBatch.Requery
    rsCurrent.MoveFirst
    Do While Not rsCurrent.EOF
        sOldName = Trim$(frmPreferences.SourcePath) & "\" & Trim$(rsBatch("FILENAME").Value)
        If LoadImage(sOldName) Then
            sNewName = Trim$(frmPreferences.TargetPath) & "\" & Trim$(rsCurrent("SUBJECT").Value) & ".BMP"
            leadMain.Save sNewName, FILE_BMP, 24, QFACTOR_PQ1, False
            Kill sOldName
        End If
        rsCurrent.MoveNext
        rsBatch.MoveNext
    Loop
    rsBatch.Close
    Set rsBatch = Nothing
End Sub

Private Sub RotateImages()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim sOldName As String
    
    '--- Rotate the images
    rsCurrent.MoveFirst
    Do While Not rsCurrent.EOF
        sOldName = Trim$(frmPreferences.TargetPath) & "\" & Trim$(rsCurrent("SUBJECT").Value) & ".BMP"
        LoadImage sOldName
        leadMain.Save sOldName, FILE_BMP, 24, QFACTOR_PQ1, False
        rsCurrent.MoveNext
    Loop
End Sub

Private Function LoadImage(sName As String) As Boolean
    If Not EZ_DEBUG Then On Error Resume Next
    If Dir(sName, vbNormal) <> "" Then
        leadMain.Load sName, 24, 0, 1
        If Val(frmPreferences.RotationAngle) <> 0 Then leadMain.Rotate Val(frmPreferences.RotationAngle) * 100, True, 0
        LoadImage = True
    Else
        LoadImage = False
    End If
End Function
