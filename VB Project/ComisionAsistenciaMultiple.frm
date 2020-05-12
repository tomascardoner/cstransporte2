VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComisionAsistenciaMultiple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Comisiones"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ComisionAsistenciaMultiple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   5310
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1455
   End
   Begin VB.CommandButton cmdCuentaCorrienteCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   2580
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtImporteContado 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   9
      Top             =   780
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3420
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja 
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   2580
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
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   960
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblCuentaCorrienteCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label lblComisionCount 
      AutoSize        =   -1  'True
      Caption         =   "Ha seleccionado # comisiones para Pagar."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4185
   End
   Begin VB.Label lblImporteContado 
      AutoSize        =   -1  'True
      Caption         =   "Confirmar Importe:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   2220
      Width           =   1320
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos de los Pagos a las Comisiones"
      Height          =   210
      Left            =   780
      TabIndex        =   8
      Top             =   300
      Width           =   3840
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ComisionAsistenciaMultiple.frx":054A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmComisionAsistenciaMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCFechaHora As Collection
Private mCIDRuta As Collection
Private mCIndice As Collection
Private mImporteTotal As Currency

Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef CFechaHora As Collection, ByRef CIDRuta As Collection, ByRef CIndice As Collection, ByVal ImporteTotal As Currency)
    Set mCFechaHora = CFechaHora
    Set mCIDRuta = CIDRuta
    Set mCIndice = CIndice
    mImporteTotal = ImporteTotal
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    lblComisionCount.Caption = "Ha seleccionado " & mCFechaHora.Count & " comisiones para Pagar"
    
    txtImporte.Text = Format(mImporteTotal, "Currency")
        
    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
        If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpFirstIfUnique) Then
            Unload Me
            Exit Sub
        End If
    Else
        If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT CuentaCorrienteCaja.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre FROM CuentaCorrienteCaja LEFT JOIN Persona ON CuentaCorrienteCaja.IDPersona = Persona.IDPersona WHERE CuentaCorrienteCaja.Activo = 1 AND (CuentaCorrienteCaja.MostrarSiempre = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & ") ORDER BY CuentaCorrienteCaja.Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpFirstIfUnique) Then
            Unload Me
            Exit Sub
        End If
    End If
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboCuentaCorrienteCaja.BoundText)
    Set recData = datcboCuentaCorrienteCaja.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCuentaCorrienteCaja.BoundText = KeySave
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ItemIndex As Integer
    
    Dim ViajeDetalle As ViajeDetalle
    
    If Not IsNumeric(txtImporteContado.Text) Then
        MsgBox "El Pago ingresado es incorrecto.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteContado.Text) <= 0 Then
        MsgBox "El Importe debe ser mayor a cero.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteContado.Text) <> mImporteTotal Then
        MsgBox "El Importe a Confirmar debe ser igual al Importe Total.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If Val(datcboCuentaCorrienteCaja.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja.", vbInformation, App.Title
        On Error Resume Next
        datcboCuentaCorrienteCaja.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Se aplicará el Pago de " & txtImporte.Text & " a las " & mCFechaHora.Count & " Comisiones seleccionadas." & vbCr & vbCr & "¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    Set ViajeDetalle = New ViajeDetalle
    ViajeDetalle.RefreshListSkip = True
    For ItemIndex = 1 To mCFechaHora.Count
        With ViajeDetalle
            .FechaHora = CDate(mCFechaHora(ItemIndex))
            .IDRuta = CStr(mCIDRuta(ItemIndex))
            .Indice = CLng(mCIndice(ItemIndex))
            If .Load() Then
                .ImporteContado = .Importe
                .IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
                If Not .Realizar() Then
                    Unload Me
                End If
            End If
        End With
    Next ItemIndex
    Set ViajeDetalle = Nothing
    
    If CSM_Forms.IsLoaded("frmComision") Then
        frmComision.FillListView Now, "", 0
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mCFechaHora = Nothing
    Set mCIDRuta = Nothing
    Set mCIndice = Nothing
End Sub

Private Sub txtImporte_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporte
End Sub

Private Sub txtImporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporte.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporte.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporte_LostFocus()
    If Not IsNumeric(txtImporte.Text) Then
        txtImporte.Text = Val(txtImporte.Text)
    End If
    txtImporte.Text = Format(CCur(txtImporte.Text), "Currency")
End Sub

Private Sub txtImporteContado_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporteContado
End Sub

Private Sub txtImporteContado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporteContado_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporteContado.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporteContado.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporteContado_LostFocus()
    If Not IsNumeric(txtImporteContado.Text) Then
        txtImporteContado.Text = Val(txtImporteContado.Text)
    End If
    txtImporteContado.Text = Format(CCur(txtImporteContado.Text), "Currency")

    lblCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
    datcboCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
End Sub

