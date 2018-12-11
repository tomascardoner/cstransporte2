VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.5#0"; "CRViewer.dll"
Begin VB.Form frmReportViewer 
   Caption         =   "Reportes"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReportViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   6855
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _cx             =   11668
      _cy             =   9128
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   11274
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CRAXDRTReport As CRAXDRT.Report

Public Sub PrinterSetup()
    CRAXDRTReport.PrinterSetup frmMDI.hwnd
End Sub

Private Sub Form_Load()
    CSM_Forms.ResizeAndPosition frmMDI, Me
End Sub

Private Sub Form_Resize()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    CRViewer.Top = CONTROL_SPACE
    CRViewer.Left = CONTROL_SPACE
    CRViewer.Height = ScaleHeight - (CONTROL_SPACE * 2)
    CRViewer.Width = ScaleWidth - (CONTROL_SPACE * 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CRAXDRTReport = Nothing
    Set frmReportViewer = Nothing
End Sub
