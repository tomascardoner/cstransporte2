VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListaPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDListaPrecio As Long
Private mNombre As String
Private mLeyenda As String
Private mDescripcion As String
Private mPrepagoEs As Boolean
Private mPrepagoVencimiento As String
Private mPrepagoReservasCantidad As Integer
Private mIDCuentaCorrienteGrupo_Credito As Long
Private mIDCuentaCorrienteGrupo_Debito As Long
Private mNotas As String
Private mActivo As Integer
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'INTERNAL VARIABLES

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDListaPrecio() As Long
    IDListaPrecio = mIDListaPrecio
End Property

Public Property Let IDListaPrecio(ByVal Value As Long)
    If Value <> mIDListaPrecio Then
        mIsDirty = True
    End If
    mIDListaPrecio = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(ByVal Value As String)
    If Value <> mNombre Then
        mIsDirty = True
    End If
    mNombre = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Leyenda() As String
    Leyenda = mLeyenda
End Property

Public Property Let Leyenda(ByVal Value As String)
    If Value <> mLeyenda Then
        mIsDirty = True
    End If
    mLeyenda = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal Value As String)
    If Value <> mDescripcion Then
        mIsDirty = True
    End If
    mDescripcion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PrepagoEs() As Boolean
    PrepagoEs = mPrepagoEs
End Property

Public Property Let PrepagoEs(ByVal Value As Boolean)
    If Value <> mPrepagoEs Then
        mIsDirty = True
    End If
    mPrepagoEs = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PrepagoVencimiento() As String
    PrepagoVencimiento = mPrepagoVencimiento
End Property

Public Property Get PrepagoVencimiento_Nombre() As String
    Select Case mPrepagoVencimiento
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6_NOMBRE
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1
            PrepagoVencimiento_Nombre = LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1_NOMBRE
        Case Else
            PrepagoVencimiento_Nombre = ""
    End Select
End Property

Public Property Get PrepagoVencimiento_ListIndex() As Integer
    Select Case mPrepagoVencimiento
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1
            PrepagoVencimiento_ListIndex = 0
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2
            PrepagoVencimiento_ListIndex = 1
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15
            PrepagoVencimiento_ListIndex = 2
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30
            PrepagoVencimiento_ListIndex = 3
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45
            PrepagoVencimiento_ListIndex = 4
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1
            PrepagoVencimiento_ListIndex = 5
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2
            PrepagoVencimiento_ListIndex = 6
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3
            PrepagoVencimiento_ListIndex = 7
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6
            PrepagoVencimiento_ListIndex = 8
        Case LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1
            PrepagoVencimiento_ListIndex = 9
        Case Else
            PrepagoVencimiento_ListIndex = -1
    End Select
End Property

Public Property Let PrepagoVencimiento(ByVal Value As String)
    If Value <> mPrepagoVencimiento Then
        mIsDirty = True
    End If
    mPrepagoVencimiento = Value
End Property

Public Property Let PrepagoVencimiento_ListIndex(ByVal Value As Integer)
    Select Case Value
        Case 0
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1
        Case 1
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2
        Case 2
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15
        Case 3
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30
        Case 4
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45
        Case 5
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1
        Case 6
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2
        Case 7
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3
        Case 8
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6
        Case 9
            mPrepagoVencimiento = LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1
        Case Else
            mPrepagoVencimiento = ""
    End Select
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PrepagoReservasCantidad() As Integer
    PrepagoReservasCantidad = mPrepagoReservasCantidad
End Property

Public Property Let PrepagoReservasCantidad(ByVal Value As Integer)
    If Value <> mPrepagoReservasCantidad Then
        mIsDirty = True
    End If
    mPrepagoReservasCantidad = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteGrupo_Credito() As Long
    IDCuentaCorrienteGrupo_Credito = mIDCuentaCorrienteGrupo_Credito
End Property

Public Property Let IDCuentaCorrienteGrupo_Credito(ByVal Value As Long)
    If Value <> mIDCuentaCorrienteGrupo_Credito Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteGrupo_Credito = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteGrupo_Debito() As Long
    IDCuentaCorrienteGrupo_Debito = mIDCuentaCorrienteGrupo_Debito
End Property

Public Property Let IDCuentaCorrienteGrupo_Debito(ByVal Value As Long)
    If Value <> mIDCuentaCorrienteGrupo_Debito Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteGrupo_Debito = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Notas() As String
    Notas = mNotas
End Property

Public Property Let Notas(ByVal Value As String)
    If Value <> mNotas Then
        mIsDirty = True
    End If
    mNotas = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Activo() As Integer
    Activo = mActivo
End Property

Public Property Let Activo(ByVal Value As Integer)
    If Value <> mActivo Then
        mIsDirty = True
    End If
    mActivo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraCreacion() As Date
    FechaHoraCreacion = mFechaHoraCreacion
End Property

Public Property Let FechaHoraCreacion(ByVal Value As Date)
    If Value <> mFechaHoraCreacion Then
        mIsDirty = True
    End If
    mFechaHoraCreacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioCreacion() As String
    IDUsuarioCreacion = mIDUsuarioCreacion
End Property

Public Property Let IDUsuarioCreacion(ByVal Value As String)
    If Value <> mIDUsuarioCreacion Then
        mIsDirty = True
    End If
    mIDUsuarioCreacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraModificacion() As Date
    FechaHoraModificacion = mFechaHoraModificacion
End Property

Public Property Let FechaHoraModificacion(ByVal Value As Date)
    If Value <> mFechaHoraModificacion Then
        mIsDirty = True
    End If
    mFechaHoraModificacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioModificacion() As String
    IDUsuarioModificacion = mIDUsuarioModificacion
End Property

Public Property Let IDUsuarioModificacion(ByVal Value As String)
    If Value <> mIDUsuarioModificacion Then
        mIsDirty = True
    End If
    mIDUsuarioModificacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get IsCopy() As Boolean
    IsCopy = mIsCopy
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Sub MakeDirty()
    mIsDirty = True
End Sub

Public Property Get NoMatch() As Boolean
    NoMatch = mNoMatch
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get NoMatchRaiseError() As Boolean
    NoMatchRaiseError = mNoMatchRaiseError
End Property

Public Property Let NoMatchRaiseError(ByVal Value As Boolean)
    mNoMatchRaiseError = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal Value As Boolean)
    mRefreshListSkip = Value
End Property

Public Sub RefreshList()
    RefreshList_Module.RefreshList_RefreshListaPrecio IDListaPrecio
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function Load() As Boolean
    Dim cmdSP As ADODB.command
    Dim recData As ADODB.Recordset

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_ListaPrecio_Get"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)
    End With

    Set recData = New ADODB.Recordset
    recData.Open cmdSP, , adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    Screen.MousePointer = vbHourglass

    Load = GetDataFromRecordset(recData)

    recData.Close
    Set recData = Nothing
    Set cmdSP = Nothing
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecio.Load", "Error al obtener los datos de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio
End Function

Private Function GetDataFromRecordset(ByRef recData As ADODB.Recordset) As Boolean

    If (Not mNoMatchRaiseError) And recData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False

        mIDListaPrecio = CLng(recData("IDListaPrecio").Value)
        mNombre = CStr(recData("Nombre").Value)
        mLeyenda = CStr(recData("Leyenda").Value & "")
        mDescripcion = CStr(recData("Descripcion").Value & "")
        mPrepagoEs = CBool(recData("PrepagoEs").Value)
        mPrepagoVencimiento = Trim(CStr(recData("PrepagoVencimiento").Value & ""))
        mPrepagoReservasCantidad = Val(recData("PrepagoReservasCantidad").Value & "")
        mIDCuentaCorrienteGrupo_Credito = Val(recData("IDCuentaCorrienteGrupo_Credito").Value & "")
        mIDCuentaCorrienteGrupo_Debito = Val(recData("IDCuentaCorrienteGrupo_Debito").Value & "")
        mNotas = CStr(recData("Notas").Value & "")
        mActivo = CBool(recData("Activo").Value)
        mFechaHoraCreacion = CDate(recData("FechaHoraCreacion").Value)
        mIDUsuarioCreacion = recData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = recData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = recData("IDUsuarioModificacion").Value
    End If

    Screen.MousePointer = vbDefault
    GetDataFromRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecio.GetDataFromRecordset", "Error al obtener los datos de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDListaPrecio = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim cmdSP As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        If mIsNew Then
            .CommandText = "usp_ListaPrecio_Add"
        Else
            .CommandText = "usp_ListaPrecio_Update"
        End If
        .CommandType = adCmdStoredProc

        If mIsNew Then
            .Parameters.Append .CreateParameter("IDListaPrecio", adInteger, adParamOutput)
        Else
            .Parameters.Append .CreateParameter("IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)
        End If
        .Parameters.Append .CreateParameter("Nombre", adVarChar, adParamInput, 50, mNombre)
        .Parameters.Append .CreateParameter("Leyenda", adVarChar, adParamInput, 50, IIf(Trim(mLeyenda) = "", Null, mLeyenda))
        .Parameters.Append .CreateParameter("Descripcion", adVarChar, adParamInput, 8000, IIf(Trim(mDescripcion) = "", Null, mDescripcion))
        .Parameters.Append .CreateParameter("PrepagoEs", adBoolean, adParamInput, , mPrepagoEs)
        .Parameters.Append .CreateParameter("PrepagoVencimiento", adChar, adParamInput, 3, IIf(Trim(mPrepagoVencimiento) = "", Null, mPrepagoVencimiento))
        .Parameters.Append .CreateParameter("PrepagoReservasCantidad", adSmallInt, adParamInput, , IIf(mPrepagoReservasCantidad = 0, Null, mPrepagoReservasCantidad))
        .Parameters.Append .CreateParameter("IDCuentaCorrienteGrupo_Credito", adInteger, adParamInput, , IIf(mIDCuentaCorrienteGrupo_Credito = 0, Null, mIDCuentaCorrienteGrupo_Credito))
        .Parameters.Append .CreateParameter("IDCuentaCorrienteGrupo_Debito", adInteger, adParamInput, , IIf(mIDCuentaCorrienteGrupo_Debito = 0, Null, mIDCuentaCorrienteGrupo_Debito))
        .Parameters.Append .CreateParameter("Notas", adVarChar, adParamInput, 8000, IIf(Trim(mNotas) = "", Null, mNotas))
        .Parameters.Append .CreateParameter("Activo", adBoolean, adParamInput, , mActivo)
        .Parameters.Append .CreateParameter("IDUsuario", adSmallInt, adParamInput, , pUsuario.IDUsuario)
        
        .Execute

        If mIsNew Then
            mIDListaPrecio = CLng(.Parameters("IDListaPrecio").Value)
        End If
    End With

    mIsCopy = False
    mIsDirty = False

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Update = True
    Exit Function

ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_COMMAND_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Lista de Precios con el mismo Nombre." & vbCr & "Si la buscó y no la encontró, verifique si está activa y si tiene permisos para verla.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.ListaPrecio.Update", "Error al actualizar la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio
        End If
    Else
        ShowErrorMessage "Classes.ListaPrecio.Update", "Error al actualizar la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio
    End If
End Function

Public Function Delete() As Boolean
    Dim cmdSP As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_ListaPrecio_Delete"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)

        .Execute
    End With

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecio.Delete", "Error al eliminar la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio
End Function

Private Sub InitializeValues()
    mNombre = ""
    mLeyenda = ""
    mDescripcion = ""
    mPrepagoEs = False
    mPrepagoVencimiento = ""
    mPrepagoReservasCantidad = 0
    mIDCuentaCorrienteGrupo_Credito = 0
    mIDCuentaCorrienteGrupo_Debito = 0
    mNotas = ""
    mActivo = True
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    Call InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
End Sub

Private Sub Class_Terminate()
    'TERMINATE CODE
End Sub
