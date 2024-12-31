Attribute VB_Name = "modMain"
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-CARDS(r)
'
' Component....: modMain.bas
'
' Procedure....: (Declarations)
'
' Description..: Public and global variables
'
' Author.......: Ronald D. Redmer
'
' History......: 08-14-00 RDR Updated & made generic w/ enhanced preferences.
'                08-15-99 RDR Designed and Programmed
'
' (c) 1997-2000 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit                                             'Require explicit variable declaration
Public Const EZ_CAPTION As String = "EZ-CARD"
Public Const EZ_DEBUG As Boolean = False
Public Const EZ_MSG_TECH_SUPPORT As String = "Please contact technical support."
Public gConn As ADODB.Connection
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: modMain.bas
'
' Procedure....: Main
'
' Description..: Application Main Procedure: Show splash screen, open database,
'                and call the main form.
'
'----------------------------------------------------------------------------
Sub Main()                                          'Application main procedure declaration
    If Not EZ_DEBUG Then On Error GoTo ErrorHandler 'Local error handler
    Load frmPreferences
    Set gConn = New ADODB.Connection
    gConn.CursorLocation = adUseClient
    gConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path & "\EZCARDS.MDB"
   
    Load frmMain                                    'Load the main form
    If EZ_DEBUG Then frmMain.Caption = frmMain.Caption & " **** DEBUG MODE **** "
    frmMain.Show                                    'Show the main application form
    Exit Sub                                        'Exit this routine
ErrorHandler:
    MsgBox "Error occured initializing application." & vbCr & EZ_MSG_TECH_SUPPORT, vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    End
End Sub





