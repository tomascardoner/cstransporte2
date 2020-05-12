VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Splash.frx":08CA
   ScaleHeight     =   3735
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCardonerSistemas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   1185
      Left            =   1980
      Picture         =   "Splash.frx":3132E
      ScaleHeight     =   77
      ScaleMode       =   0  'User
      ScaleWidth      =   250
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3780
   End
   Begin VB.Image imgCompany 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   660
      Top             =   1140
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.Label lblLicense 
      BackStyle       =   0  'Transparent
      Caption         =   "Permitido su uso a:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   3360
      Width           =   5640
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1500
      TabIndex        =   1
      Top             =   720
      Width           =   4260
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   5520
   End
   Begin VB.Shape shpBorder 
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   5985
   End
   Begin VB.Image imgApp 
      Height          =   720
      Left            =   180
      Picture         =   "Splash.frx":362EC
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1680
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Caption = App.Title
    lblTitle.Caption = App.Title
    
    ' Version info
    Dim ExeFileDate As Date
    ExeFileDate = FileDateTime(App.Path & "\" & App.EXEName & ".exe")
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision & " (" & Format(ExeFileDate, "yyyyMMdd") & ")"

    lblCopyright.Caption = App.LegalCopyright
End Sub
