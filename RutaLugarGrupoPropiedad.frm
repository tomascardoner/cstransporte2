VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRutaLugarGrupoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RutaLugarGrupoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   5670
   Begin VB.CommandButton cmdLugar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   1980
      Width           =   255
   End
   Begin VB.CommandButton cmdLugarGrupo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   1500
      Width           =   255
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   1020
      Width           =   255
   End
   Begin VB.Frame fraLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   5415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboLugarGrupo 
      Height          =   330
      Left            =   1920
      TabIndex        =   4
      Top             =   1500
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
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   1020
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
   Begin MSDataListLib.DataCombo datcboLugar 
      Height          =   330
      Left            =   1920
      TabIndex        =   7
      Top             =   1980
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
   Begin VB.Label lblLugar 
      AutoSize        =   -1  'True
      Caption         =   "Lugar predeterminado:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1635
   End
   Begin VB.Label lblLugarGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos de la Ruta-Grupo de Lugar"
      Height          =   210
      Left            =   780
      TabIndex        =   11
      Top             =   300
      Width           =   3570
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmRutaLugarGrupoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRutaLugarGrupo As RutaLugarGrupo

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef RutaLugarGrupo As RutaLugarGrupo)
    Set mRutaLugarGrupo = RutaLugarGrupo
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mRutaLugarGrupo
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpItemOrNone, .IDRuta) Then
            Unload Me
            Exit Sub
        End If
        datcboLugarGrupo.BoundText = .IDLugarGrupo
        datcboLugar.BoundText = .IDLugarPredeterminado
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = datcboRuta.BoundText
    Set recData = datcboRuta.RowSource
    recData.Requery
    Set recData = Nothing
    datcboRuta.BoundText = KeySave
End Sub

Private Sub datcboRuta_Change()
    FillComboBoxLugarGrupo
End Sub

Private Sub cmdRuta_Click()
    If pCPermiso.GotPermission(PERMISO_RUTA) Then
        Screen.MousePointer = vbHourglass
        frmRuta.Show
        On Error Resume Next
        Set frmRuta.lvwData.SelectedItem = frmRuta.lvwData.ListItems(KEY_STRINGER & datcboRuta.BoundText)
        frmRuta.lvwData.SelectedItem.EnsureVisible
        If frmRuta.WindowState = vbMinimized Then
            frmRuta.WindowState = vbNormal
        End If
        frmRuta.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub FillComboBoxLugarGrupo()
    If datcboRuta.BoundText <> "" Then
        Call CSM_Control_DataCombo.FillFromSQL(datcboLugarGrupo, "SELECT DISTINCT LugarGrupo.IDLugarGrupo, LugarGrupo.Nombre FROM LugarGrupo INNER JOIN RutaDetalle ON LugarGrupo.IDLugarGrupo = RutaDetalle.IDLugarGrupo WHERE RutaDetalle.IDRuta = '" & datcboRuta.BoundText & "' AND LugarGrupo.IDLugarGrupo <> " & pParametro.LugarGrupo_ID_Otro & " AND LugarGrupo.Activo = 1 ORDER BY LugarGrupo.Nombre", "IDLugarGrupo", "Nombre", "Grupos de Lugares", cscpNone)
    End If
End Sub

Private Sub datcboLugarGrupo_Change()
    FillComboBoxLugar
End Sub

Private Sub cmdLugarGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmLugarGrupo.Show
        On Error Resume Next
        Set frmLugarGrupo.lvwData.SelectedItem = frmLugarGrupo.lvwData.ListItems(KEY_STRINGER & datcboLugarGrupo.BoundText)
        frmLugarGrupo.lvwData.SelectedItem.EnsureVisible
        If frmLugarGrupo.WindowState = vbMinimized Then
            frmLugarGrupo.WindowState = vbNormal
        End If
        frmLugarGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub FillComboBoxLugar()
    If datcboRuta.BoundText <> "" And Val(datcboLugarGrupo.BoundText) > 0 Then
        Call CSM_Control_DataCombo.FillFromSQL(datcboLugar, "SELECT Lugar.IDLugar, Lugar.Nombre FROM Lugar INNER JOIN RutaDetalle ON Lugar.IDLugar = RutaDetalle.IDLugar WHERE RutaDetalle.IDRuta = '" & datcboRuta.BoundText & "' AND RutaDetalle.IDLugarGrupo = " & Val(datcboLugarGrupo.BoundText) & " AND Lugar.IDLugar <> " & pParametro.Lugar_ID_Otro & " AND Lugar.Activo = 1 ORDER BY RutaDetalle.Indice, Lugar.Nombre", "IDLugar", "Nombre", "Lugares", cscpNone)
    End If
End Sub

Private Sub cmdLugar_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboLugar.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdOK_Click()
    If datcboRuta.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRuta.SetFocus
        Exit Sub
    End If
    If Val(datcboLugarGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboLugarGrupo.SetFocus
        Exit Sub
    End If
    If Val(datcboLugar.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Lugar.", vbInformation, App.Title
        datcboLugar.SetFocus
        Exit Sub
    End If
        
    With mRutaLugarGrupo
        .IDRuta = datcboRuta.BoundText
        .IDLugarGrupo = Val(datcboLugarGrupo.BoundText)
        .IDLugarPredeterminado = Val(datcboLugar.BoundText)
        If Not .Update() Then
            Exit Sub
        End If
    End With
        
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRutaLugarGrupo = Nothing
    Set frmRutaLugarGrupoPropiedad = Nothing
End Sub
