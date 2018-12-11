VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
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
      TabIndex        =   4
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
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
   Begin VB.TextBox txtIDUsuario 
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
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   1
      Top             =   120
      Width           =   2910
   End
   Begin VB.Label lblPassword 
      Caption         =   "C&ontraseña:"
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
   Begin VB.Label lblIDUsuario 
      Caption         =   "&Usuario:"
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
      Width           =   600
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIntentos As Long

Private Sub cmdCancel_Click()
    pUsuario.IDUsuario = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim UsuarioGrupo As UsuarioGrupo
    
    If Trim(txtIDUsuario.Text) = "" Then
        MsgBox "Debe ingresar el Usuario.", vbInformation, App.Title
        txtIDUsuario.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "Debe ingresar la Contraseña.", vbInformation, App.Title
        txtPassword.SetFocus
        Exit Sub
    End If
    
    pUsuario.LoginName = LCase(txtIDUsuario.Text)
    pUsuario.NoMatchRaiseError = False
    pUsuario.Requery
    If Not pUsuario.LoadByLoginName() Then
        pUsuario.NoMatchRaiseError = True
        Exit Sub
    End If
    pUsuario.NoMatchRaiseError = True
    If pUsuario.NoMatch Then
        mIntentos = mIntentos + 1
        WriteLogEvent "User Login Failed: User Unknown - LoginName: " & pUsuario.LoginName, vbLogEventTypeWarning
        MsgBox "El Usuario ingresado no existe.", vbExclamation, App.Title
        txtIDUsuario.SetFocus
        txtIDUsuario_GotFocus
        If mIntentos = 3 Then
            MsgBox "Ha realizado 3 intentos de ingreso incorrectos." & vbCr & "Se cerrará el Sistema.", vbExclamation, App.Title
            pUsuario.IDUsuario = 0
            pUsuario.LoginName = ""
            Unload Me
        End If
        Exit Sub
    End If
    
    If Not pUsuario.Activo Then
        mIntentos = mIntentos + 1
        WriteLogEvent "User Login Failed: User Not Active - LoginName: " & pUsuario.LoginName, vbLogEventTypeWarning
        MsgBox "El Usuario está desactivado.", vbExclamation, App.Title
        txtIDUsuario.SetFocus
        txtIDUsuario_GotFocus
        If mIntentos = 3 Then
            MsgBox "Ha realizado 3 intentos de ingreso incorrectos." & vbCr & "Se cerrará el Sistema.", vbExclamation, App.Title
            pUsuario.IDUsuario = 0
            pUsuario.LoginName = ""
            Unload Me
        End If
        Exit Sub
    End If
    
    If txtPassword.Text <> pUsuario.Password Then
        mIntentos = mIntentos + 1
        WriteLogEvent "User Login Failed: Wrong Password - LoginName: " & pUsuario.LoginName, vbLogEventTypeWarning
        MsgBox "La Contraseña ingresada es incorrecta.", vbExclamation, App.Title
        txtPassword.SetFocus
        txtPassword_GotFocus
        If mIntentos = 3 Then
            MsgBox "Ha realizado 3 intentos de ingreso incorrectos." & vbCr & "Se cerrará el Sistema.", vbExclamation, App.Title
            pUsuario.IDUsuario = 0
            pUsuario.LoginName = ""
            Unload Me
        End If
        Exit Sub
    End If
    
    Set UsuarioGrupo = New UsuarioGrupo
    UsuarioGrupo.IDUsuarioGrupo = pUsuario.IDUsuarioGrupo
    If Not UsuarioGrupo.Load() Then
        Set UsuarioGrupo = Nothing
        Exit Sub
    End If
    If Not UsuarioGrupo.Activo Then
        mIntentos = mIntentos + 1
        WriteLogEvent "User Login Failed: User Group Not Active - LogiName: " & pUsuario.LoginName, vbLogEventTypeWarning
        MsgBox "El Grupo de Usuarios está desactivado.", vbExclamation, App.Title
        txtIDUsuario.SetFocus
        txtIDUsuario_GotFocus
        Set UsuarioGrupo = Nothing
        If mIntentos = 3 Then
            MsgBox "Ha realizado 3 intentos de ingreso incorrectos." & vbCr & "Se cerrará el Sistema.", vbExclamation, App.Title
            pUsuario.IDUsuario = 0
            pUsuario.LoginName = ""
            Unload Me
        End If
        Exit Sub
    End If
    Set UsuarioGrupo = Nothing
    
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub txtIDUsuario_GotFocus()
    CSM_Control_TextBox.SelAllText txtIDUsuario
End Sub

Private Sub txtPassword_GotFocus()
    CSM_Control_TextBox.SelAllText txtPassword
End Sub
