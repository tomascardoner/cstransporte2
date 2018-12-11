VERSION 5.00
Begin VB.Form frmMensaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                      ...:::| Centro de Mensajes |:::..."
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mensaje.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9525
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1980
      Picture         =   "Mensaje.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   4380
      Width           =   555
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   435
      Left            =   7980
      TabIndex        =   3
      Top             =   4380
      Width           =   1395
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "&Siguiente >>"
      Height          =   435
      Left            =   5220
      TabIndex        =   2
      Top             =   4380
      Width           =   1395
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<< &Anterior"
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   4380
      Width           =   1395
   End
   Begin VB.TextBox txtMensaje 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label lblMensajeNumero 
      Alignment       =   2  'Center
      Caption         =   "Mensaje N° 1 de 99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   4
      Top             =   4500
      Width           =   2235
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrecData As ADODB.Recordset
Private mCloseButton As Boolean
Private maMessageRead() As Boolean

Private Sub cmdAnterior_Click()
    If pTrapErrors Then
        On Error GoTo errorMessage
    End If
    
    mrecData.MovePrevious
    
    Call ShowMessage
    Exit Sub

errorMessage:
    CSM_Error.ShowErrorMessage "Forms.Mensaje.Anterior", "Error al leer el Mensaje Anterior."
End Sub

Private Sub cmdSiguiente_Click()
    If pTrapErrors Then
        On Error GoTo errorMessage
    End If
    
    mrecData.MoveNext
    
    Call ShowMessage
    Exit Sub

errorMessage:
    CSM_Error.ShowErrorMessage "Forms.Mensaje.Siguiente", "Error al leer el Mensaje Siguiente."
End Sub

Private Sub cmdCerrar_Click()
    Unload frmMensaje
End Sub

Private Sub Form_Load()
    Call CSM_Forms.CenterToParent(frmMDI, Me)
    
    Call ShowMessage
End Sub

Public Function CheckMessages() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo errorMessage
    End If

    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Mensaje_GetList"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDUsuario", adVarChar, adParamInput, 20, pUsuario.IDUsuario)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockReadOnly
    Set cmdData = Nothing
    
    CheckMessages = (mrecData.RecordCount > 0)
    mCloseButton = (mrecData.RecordCount = 1)
    If (mrecData.RecordCount > 0) Then
        ReDim maMessageRead(1 To mrecData.RecordCount) As Boolean
    End If
    Exit Function
    
errorMessage:
    CSM_Error.ShowErrorMessage "Forms.Mensaje.CheckMessages", "Error al verificar los mensajes"
End Function

Private Function ShowMessage() As Boolean
    Dim Mensaje_Usuario As Mensaje_Usuario
    
    txtMensaje.Text = mrecData("Mensaje").Value

    cmdAnterior.Enabled = mrecData.AbsolutePosition > 1
    lblMensajeNumero.Caption = "Mensaje N° " & mrecData.AbsolutePosition & " de " & mrecData.RecordCount
    cmdSiguiente.Enabled = mrecData.AbsolutePosition < mrecData.RecordCount
    
    If mrecData.AbsolutePosition = mrecData.RecordCount Then
        mCloseButton = True
    End If
    cmdCerrar.Enabled = mCloseButton
    
    If Not maMessageRead(mrecData.AbsolutePosition) Then
        Set Mensaje_Usuario = New Mensaje_Usuario
        Mensaje_Usuario.IDMensaje = mrecData("IDMensaje").Value
        Mensaje_Usuario.IDUsuario = pUsuario.IDUsuario
        Mensaje_Usuario.NoMatchRaiseError = False
        If Mensaje_Usuario.Load() Then
            Mensaje_Usuario.LeidoVeces = Mensaje_Usuario.LeidoVeces + 1
            Mensaje_Usuario.Update
        End If
        Set Mensaje_Usuario = Nothing
        
        maMessageRead(mrecData.AbsolutePosition) = True
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    If mrecData.State = adStateOpen Then
        mrecData.Close
    End If
    Set mrecData = Nothing
    Set frmMensaje = Nothing
End Sub
