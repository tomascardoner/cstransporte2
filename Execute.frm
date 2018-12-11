VERSION 5.00
Begin VB.Form frmExecute 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecutar"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   Icon            =   "Execute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Examinar..."
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtExecute 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   900
      Width           =   4035
   End
   Begin VB.Label lblOpen 
      Caption         =   "&Abrir:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblLegend 
      Caption         =   "Escriba el nombre del programa, carpeta, documento o recurso de internet que desea que Windows abra."
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4035
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "Execute.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    Dim FileName As String
    
    FileName = CSM_CommonDialog.FileOpen(Me.hWnd, "Seleccione el Programa a Ejecutar...", "Archivos ejecutables (*.exe; *.pif)|*.exe;*.pif|Todos los archivos (*.*)|*.*", "")
    If FileName <> "" Then
        txtExecute.Text = FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmExecute
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrorHandler
    
    If Trim(txtExecute.Text) <> "" Then
        CSM_Instance.Execute frmMDI.hWnd, txtExecute.Text
    End If
    Unload frmExecute
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Execute.OK", "Error al abrir el objeto especificado." & vbCr & vbCr & txtExecute.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExecute = Nothing
End Sub

Private Sub txtExecute_GotFocus()
    CSM_Control_TextBox.SelAllText txtExecute
End Sub
