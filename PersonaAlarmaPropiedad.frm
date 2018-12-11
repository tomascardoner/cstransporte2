VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersonaAlarmaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "PersonaAlarmaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3660
      Picture         =   "PersonaAlarmaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdPersona 
      Height          =   315
      Left            =   3720
      Picture         =   "PersonaAlarmaPropiedad.frx":0B74
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar..."
      Top             =   960
      Width           =   315
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   2715
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Ultimo"
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   555
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
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
      Left            =   120
      TabIndex        =   12
      Top             =   3900
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtPreaviso 
      Alignment       =   1  'Right Justify
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
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2220
      Width           =   1035
   End
   Begin VB.CommandButton cmdPersonaAlarmaGrupo 
      Caption         =   "..."
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
      Left            =   4380
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Grupos"
      Top             =   1380
      Width           =   255
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2820
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   17
      Top             =   780
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   14
      Top             =   4500
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
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
      Height          =   375
      Left            =   2100
      TabIndex        =   13
      Top             =   4500
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboPersonaAlarmaGrupo 
      Height          =   330
      Left            =   1020
      TabIndex        =   5
      Top             =   1380
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1020
      TabIndex        =   7
      Top             =   1800
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   103284737
      CurrentDate     =   36950
   End
   Begin VB.Label lblPreavisoUnidad 
      AutoSize        =   -1  'True
      Caption         =   "días antes."
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
      Left            =   2220
      TabIndex        =   18
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lblPreaviso 
      AutoSize        =   -1  'True
      Caption         =   "Aviso:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Width           =   495
   End
   Begin VB.Label lblPersonaAlarmaGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "&Persona:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   645
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   465
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos de la Alarma de Personas"
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
      Left            =   780
      TabIndex        =   16
      Top             =   300
      Width           =   2325
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaAlarmaPropiedad.frx":10FE
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPersonaAlarmaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPersonaAlarma As PersonaAlarma
Private mNew As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mPersonaAlarma
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef PersonaAlarma As PersonaAlarma)
    Dim Persona As Persona
    
    Set mPersonaAlarma = PersonaAlarma
    mNew = (mPersonaAlarma.IDPersona = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mPersonaAlarma
        If mNew Then
            txtPersona.Tag = 0
            txtPersona.Text = ""
        Else
            txtPersona.Tag = .IDPersona
            Set Persona = New Persona
            Persona.IDPersona = .IDPersona
            If Not Persona.Load() Then
                Set Persona = Nothing
                Unload Me
                Exit Sub
            End If
            txtPersona.Text = Persona.ApellidoNombre
            Set Persona = Nothing
        End If
    
        If Not CSM_Control_DataCombo.FillFromSQL(datcboPersonaAlarmaGrupo, "SELECT IDPersonaAlarmaGrupo, Nombre FROM PersonaAlarmaGrupo WHERE Activo = 1 OR IDPersonaAlarmaGrupo = " & .IDPersonaAlarmaGrupo & " ORDER BY Nombre", "IDPersonaAlarmaGrupo", "Nombre", "Grupos de Alarmas de Personas", cscpItemOrfirst, .IDPersonaAlarmaGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        If mNew Then
            dtpFecha.Value = Date
        Else
            dtpFecha.Value = .Fecha
        End If
        txtPreaviso.Text = .Preaviso
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub FillComboBoxPersonaAlarmaGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboPersonaAlarmaGrupo.BoundText)
    Set recData = datcboPersonaAlarmaGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboPersonaAlarmaGrupo.BoundText = KeySave
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Val(txtPersona.Tag) = 0 Then
        MsgBox "Debe seleccionar la Persona.", vbInformation, App.Title
        cmdPersona.SetFocus
        Exit Sub
    End If
    
    If Val(datcboPersonaAlarmaGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboPersonaAlarmaGrupo.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtPreaviso.Text) Then
        MsgBox "El Aviso debe ser un valor numérico.", vbInformation, App.Title
        txtPreaviso.SetFocus
        Exit Sub
    End If
    If CLng(txtPreaviso.Text) < 0 Then
        MsgBox "El Aviso debe ser mayor o igual a cero.", vbInformation, App.Title
        txtPreaviso.SetFocus
        Exit Sub
    End If
    
    With mPersonaAlarma
        .IDPersona = Val(txtPersona.Tag)
        .IDPersonaAlarmaGrupo = Val(datcboPersonaAlarmaGrupo.BoundText)
        .Fecha = dtpFecha.Value
        .Preaviso = CLng(txtPreaviso.Text)
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update Then
                Exit Sub
            End If
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersona.Tag), UCase(Left(txtPersona.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdPersonaAlarmaGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA_ALARMA_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmPersonaAlarmaGrupo.Show
        On Error Resume Next
        Set frmPersonaAlarmaGrupo.lvwData.SelectedItem = frmPersonaAlarmaGrupo.lvwData.ListItems(KEY_STRINGER & datcboPersonaAlarmaGrupo.BoundText)
        frmPersonaAlarmaGrupo.lvwData.SelectedItem.EnsureVisible
        If frmPersonaAlarmaGrupo.WindowState = vbMinimized Then
            frmPersonaAlarmaGrupo.WindowState = vbNormal
        End If
        frmPersonaAlarmaGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        PersonaSelected Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex)), "PP"
    End If
    cmdPersona.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPersonaAlarma = Nothing
    Set frmPersonaAlarmaPropiedad = Nothing
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub txtPreaviso_GotFocus()
    CSM_Control_TextBox.SelAllText txtPreaviso
End Sub

Private Sub txtPreaviso_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPreaviso_LostFocus()
    txtPreaviso.Text = Val(txtPreaviso.Text)
End Sub

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    If IDPersona = 0 Then
        Exit Sub
    End If
    
    txtPersona.Tag = IDPersona
    
    txtPersona.Text = frmMDI.cboPersona.Text
End Sub

