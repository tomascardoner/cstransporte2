VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConfirm 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1080
      Width           =   2910
   End
   Begin VB.TextBox txtNew 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2910
   End
   Begin VB.TextBox txtOld 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2910
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   780
      TabIndex        =   6
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2340
      TabIndex        =   7
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblConfirm 
      Caption         =   "C&onfirma:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   1140
      Width           =   885
   End
   Begin VB.Label lblNew 
      Caption         =   "N&ueva:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   885
   End
   Begin VB.Label lblOld 
      Caption         =   "A&nterior:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIntentos As Long
Private mUsuario As Usuario

Public Sub LoadDataAndShow()
    Set mUsuario = New Usuario
    
    mUsuario.IDUsuario = pUsuario.IDUsuario
    If Not mUsuario.Load() Then
        Exit Sub
    End If
    
    Load Me
    
    Me.Show vbModal, frmMDI
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtOld.Text) = "" Then
        MsgBox "Debe ingresar la Contraseña anterior.", vbInformation, App.Title
        txtOld.SetFocus
        Exit Sub
    End If
    If txtOld.Text <> mUsuario.Password Then
        mIntentos = mIntentos + 1
        MsgBox "La Contraseña anterior es incorrecta.", vbExclamation, App.Title
        txtOld.SetFocus
        txtOld_GotFocus
        If mIntentos = 3 Then
            MsgBox "Ha realizado 3 intentos." & vbCr & "Se cerrará la ventana de Cambio de Contraseña.", vbExclamation, App.Title
            Unload Me
        End If
        Exit Sub
    End If
    If Trim(txtNew.Text) = "" Then
        MsgBox "Debe ingresar la nueva Contraseña.", vbInformation, App.Title
        txtNew.SetFocus
        Exit Sub
    End If
    If Len(txtNew.Text) < 4 Then
        MsgBox "La Contraseña debe contener al menos 4 letras o números.", vbInformation, App.Title
        txtNew.SetFocus
        txtNew_GotFocus
        Exit Sub
    End If
    If Trim(txtConfirm.Text) = "" Then
        MsgBox "Debe confirmar la nueva Contraseña.", vbInformation, App.Title
        txtConfirm.SetFocus
        Exit Sub
    End If
    If txtNew.Text <> txtConfirm.Text Then
        MsgBox "La nueva Contraseña no coincide con la confirmación.", vbInformation, App.Title
        txtConfirm.SetFocus
        txtConfirm_GotFocus
        Exit Sub
    End If
    
    WriteLogEvent "El Usuario ha cambiado su contraseña:", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    WriteLogEvent "     Usuario: " & pUsuario.IDUsuario, vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    mUsuario.Password = txtNew.Text
    
    If Not mUsuario.Update Then
        Exit Sub
    End If
    
    MsgBox "La Contraseña ha sido cambiada.", vbInformation, App.Title
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mUsuario = Nothing
    Set frmChangePassword = Nothing
End Sub

Private Sub txtConfirm_GotFocus()
    CSM_Control_TextBox.SelAllText txtConfirm
End Sub

Private Sub txtNew_GotFocus()
    CSM_Control_TextBox.SelAllText txtNew
End Sub

Private Sub txtOld_GotFocus()
    CSM_Control_TextBox.SelAllText txtOld
End Sub
