VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDContacto As Long
Public Titulo As String
Public Apellido As String
Public Nombre As String
Public Compania As String
Public TituloLaboral As String
Public IDTelefono1Tipo As Byte
Public Telefono1TipoOtro As String
Public Telefono1Area As String
Public Telefono1Numero As String
Public IDTelefono2Tipo As Byte
Public Telefono2TipoOtro As String
Public Telefono2Area As String
Public Telefono2Numero As String
Public IDTelefono3Tipo As Byte
Public Telefono3TipoOtro As String
Public Telefono3Area As String
Public Telefono3Numero As String
Public IDTelefono4Tipo As Byte
Public Telefono4TipoOtro As String
Public Telefono4Area As String
Public Telefono4Numero As String
Public IDTelefono5Tipo As Byte
Public Telefono5TipoOtro As String
Public Telefono5Area As String
Public Telefono5Numero As String
Public IDContactoGrupo As Long
Public DomicilioLaboralCalle1 As String
Public DomicilioLaboralNumero As String
Public DomicilioLaboralPiso As String
Public DomicilioLaboralDepartamento As String
Public DomicilioLaboralCalle2 As String
Public DomicilioLaboralCalle3 As String
Public DomicilioLaboralCodigoPostal As String
Public DomicilioLaboralIDProvincia As String
Public DomicilioLaboralIDLocalidad As Long
Public DomicilioLaboralMailing As Boolean
Public DomicilioParticularCalle1 As String
Public DomicilioParticularNumero As String
Public DomicilioParticularPiso As String
Public DomicilioParticularDepartamento As String
Public DomicilioParticularCalle2 As String
Public DomicilioParticularCalle3 As String
Public DomicilioParticularCodigoPostal As String
Public DomicilioParticularIDProvincia As String
Public DomicilioParticularIDLocalidad As Long
Public DomicilioParticularMailing As Boolean
Public DomicilioOtroNombre As String
Public DomicilioOtroCalle1 As String
Public DomicilioOtroNumero As String
Public DomicilioOtroPiso As String
Public DomicilioOtroDepartamento As String
Public DomicilioOtroCalle2 As String
Public DomicilioOtroCalle3 As String
Public DomicilioOtroCodigoPostal As String
Public DomicilioOtroIDProvincia As String
Public DomicilioOtroIDLocalidad As Long
Public DomicilioOtroMailing As Boolean
Public Email1 As String
Public Email1Nombre As String
Public Email2 As String
Public Email2Nombre As String
Public Email3 As String
Public Email3Nombre As String
Public PaginaWeb As String
Public SobreNombre As String
Public FechaNacimiento As Date
Public Asistente As String
Public Notas As String
Public Activo As Boolean
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public NoMatchRaiseError As Boolean
Public NoMatch As Boolean

Public RefreshList As Boolean

Private mrecData As ADODB.Recordset

Public Property Get DisplayName() As String
    If Trim(Apellido) <> "" Then
        If Trim(Nombre) <> "" Then
            DisplayName = Apellido & ", " & Nombre
        Else
            DisplayName = Apellido
        End If
    ElseIf Trim(Nombre) <> "" Then
        DisplayName = Nombre
    ElseIf Trim(Compania) = "" Then
        DisplayName = Compania
    End If
End Property

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Contacto_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDContacto", adInteger, adParamInput, , IDContacto)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Contacto.OpenRecordset", "Error al abrir la tabla de Contactos."
End Function

Public Function Load() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    If Not OpenRecordset() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If (Not NoMatchRaiseError) And mrecData.EOF Then
        NoMatch = True
        
        InitializeValues
    Else
        NoMatch = False
        
        Titulo = mrecData("Titulo").Value & ""
        Apellido = mrecData("Apellido").Value & ""
        Nombre = mrecData("Nombre").Value & ""
        Compania = mrecData("Compania").Value & ""
        TituloLaboral = mrecData("TituloLaboral").Value & ""
        IDTelefono1Tipo = Val(mrecData("IDTelefono1Tipo").Value & "")
        Telefono1TipoOtro = mrecData("Telefono1TipoOtro").Value & ""
        Telefono1Area = mrecData("Telefono1Area").Value & ""
        Telefono1Numero = mrecData("Telefono1Numero").Value & ""
        IDTelefono2Tipo = Val(mrecData("IDTelefono2Tipo").Value & "")
        Telefono2TipoOtro = mrecData("Telefono2TipoOtro").Value & ""
        Telefono2Area = mrecData("Telefono2Area").Value & ""
        Telefono2Numero = mrecData("Telefono2Numero").Value & ""
        IDTelefono3Tipo = Val(mrecData("IDTelefono3Tipo").Value & "")
        Telefono3TipoOtro = mrecData("Telefono3TipoOtro").Value & ""
        Telefono3Area = mrecData("Telefono3Area").Value & ""
        Telefono3Numero = mrecData("Telefono3Numero").Value & ""
        IDTelefono4Tipo = Val(mrecData("IDTelefono4Tipo").Value & "")
        Telefono4TipoOtro = mrecData("Telefono4TipoOtro").Value & ""
        Telefono4Area = mrecData("Telefono4Area").Value & ""
        Telefono4Numero = mrecData("Telefono4Numero").Value & ""
        IDTelefono5Tipo = Val(mrecData("IDTelefono5Tipo").Value & "")
        Telefono5TipoOtro = mrecData("Telefono5TipoOtro").Value & ""
        Telefono5Area = mrecData("Telefono5Area").Value & ""
        Telefono5Numero = mrecData("Telefono5Numero").Value & ""
        IDContactoGrupo = mrecData("IDContactoGrupo").Value
        DomicilioLaboralCalle1 = mrecData("DomicilioLaboralCalle1").Value & ""
        DomicilioLaboralNumero = mrecData("DomicilioLaboralNumero").Value & ""
        DomicilioLaboralPiso = mrecData("DomicilioLaboralPiso").Value & ""
        DomicilioLaboralDepartamento = mrecData("DomicilioLaboralDepartamento").Value & ""
        DomicilioLaboralCalle2 = mrecData("DomicilioLaboralCalle2").Value & ""
        DomicilioLaboralCalle3 = mrecData("DomicilioLaboralCalle3").Value & ""
        DomicilioLaboralCodigoPostal = mrecData("DomicilioLaboralCodigoPostal").Value & ""
        DomicilioLaboralIDProvincia = mrecData("DomicilioLaboralIDProvincia").Value & ""
        DomicilioLaboralIDLocalidad = Val(mrecData("DomicilioLaboralIDLocalidad").Value & "")
        DomicilioLaboralMailing = mrecData("DomicilioLaboralMailing").Value
        DomicilioParticularCalle1 = mrecData("DomicilioParticularCalle1").Value & ""
        DomicilioParticularNumero = mrecData("DomicilioParticularNumero").Value & ""
        DomicilioParticularPiso = mrecData("DomicilioParticularPiso").Value & ""
        DomicilioParticularDepartamento = mrecData("DomicilioParticularDepartamento").Value & ""
        DomicilioParticularCalle2 = mrecData("DomicilioParticularCalle2").Value & ""
        DomicilioParticularCalle3 = mrecData("DomicilioParticularCalle3").Value & ""
        DomicilioParticularCodigoPostal = mrecData("DomicilioParticularCodigoPostal").Value & ""
        DomicilioParticularIDProvincia = mrecData("DomicilioParticularIDProvincia").Value & ""
        DomicilioParticularIDLocalidad = Val(mrecData("DomicilioParticularIDLocalidad").Value & "")
        DomicilioParticularMailing = mrecData("DomicilioParticularMailing").Value
        DomicilioOtroNombre = mrecData("DomicilioOtroNombre").Value & ""
        DomicilioOtroCalle1 = mrecData("DomicilioOtroCalle1").Value & ""
        DomicilioOtroNumero = mrecData("DomicilioOtroNumero").Value & ""
        DomicilioOtroPiso = mrecData("DomicilioOtroPiso").Value & ""
        DomicilioOtroDepartamento = mrecData("DomicilioOtroDepartamento").Value & ""
        DomicilioOtroCalle2 = mrecData("DomicilioOtroCalle2").Value & ""
        DomicilioOtroCalle3 = mrecData("DomicilioOtroCalle3").Value & ""
        DomicilioOtroCodigoPostal = mrecData("DomicilioOtroCodigoPostal").Value & ""
        DomicilioOtroIDProvincia = mrecData("DomicilioOtroIDProvincia").Value & ""
        DomicilioOtroIDLocalidad = Val(mrecData("DomicilioOtroIDLocalidad").Value & "")
        DomicilioOtroMailing = mrecData("DomicilioOtroMailing").Value
        Email1 = mrecData("Email1").Value & ""
        Email1Nombre = mrecData("Email1Nombre").Value & ""
        Email2 = mrecData("Email2").Value & ""
        Email2Nombre = mrecData("Email2Nombre").Value & ""
        Email3 = mrecData("Email3").Value & ""
        Email3Nombre = mrecData("Email3Nombre").Value & ""
        PaginaWeb = mrecData("PaginaWeb").Value & ""
        SobreNombre = mrecData("SobreNombre").Value & ""
        FechaNacimiento = IIf(IsNull(mrecData("FechaNacimiento").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaNacimiento").Value)
        Asistente = mrecData("Asistente").Value & ""
        Notas = mrecData("Notas").Value & ""
        Activo = mrecData("Activo").Value
        FechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Contacto.Load", "Error al obtener los datos del Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
End Function

Public Function AddNew() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Contacto_IDMax"
    cmdData.CommandType = adCmdStoredProc
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    IDContacto = Val(recData("IDContactoMax").Value & "") + 1
    recData.Close
    Set recData = Nothing
    
    mrecData.AddNew
    mrecData("IDContacto").Value = IDContacto
    mrecData("Titulo").Value = IIf(Trim(Titulo) = "", Null, Titulo)
    mrecData("Apellido").Value = IIf(Trim(Apellido) = "", Null, Apellido)
    mrecData("Nombre").Value = IIf(Trim(Nombre) = "", Null, Nombre)
    mrecData("Compania").Value = IIf(Trim(Compania) = "", Null, Compania)
    mrecData("TituloLaboral").Value = IIf(Trim(TituloLaboral) = "", Null, TituloLaboral)
    mrecData("IDTelefono1Tipo").Value = IIf(IDTelefono1Tipo = 0, Null, IDTelefono1Tipo)
    mrecData("Telefono1TipoOtro").Value = IIf(Trim(Telefono1TipoOtro) = "", Null, Telefono1TipoOtro)
    mrecData("Telefono1Area").Value = IIf(Trim(Telefono1Area) = "", Null, Telefono1Area)
    mrecData("Telefono1Numero").Value = IIf(Trim(Telefono1Numero) = "", Null, Telefono1Numero)
    mrecData("IDTelefono2Tipo").Value = IIf(IDTelefono2Tipo = 0, Null, IDTelefono2Tipo)
    mrecData("Telefono2TipoOtro").Value = IIf(Trim(Telefono2TipoOtro) = "", Null, Telefono2TipoOtro)
    mrecData("Telefono2Area").Value = IIf(Trim(Telefono2Area) = "", Null, Telefono2Area)
    mrecData("Telefono2Numero").Value = IIf(Trim(Telefono2Numero) = "", Null, Telefono2Numero)
    mrecData("IDTelefono3Tipo").Value = IIf(IDTelefono3Tipo = 0, Null, IDTelefono3Tipo)
    mrecData("Telefono3TipoOtro").Value = IIf(Trim(Telefono3TipoOtro) = "", Null, Telefono3TipoOtro)
    mrecData("Telefono3Area").Value = IIf(Trim(Telefono3Area) = "", Null, Telefono3Area)
    mrecData("Telefono3Numero").Value = IIf(Trim(Telefono3Numero) = "", Null, Telefono3Numero)
    mrecData("IDTelefono4Tipo").Value = IIf(IDTelefono4Tipo = 0, Null, IDTelefono4Tipo)
    mrecData("Telefono4TipoOtro").Value = IIf(Trim(Telefono4TipoOtro) = "", Null, Telefono4TipoOtro)
    mrecData("Telefono4Area").Value = IIf(Trim(Telefono4Area) = "", Null, Telefono4Area)
    mrecData("Telefono4Numero").Value = IIf(Trim(Telefono4Numero) = "", Null, Telefono4Numero)
    mrecData("IDTelefono5Tipo").Value = IIf(IDTelefono5Tipo = 0, Null, IDTelefono5Tipo)
    mrecData("Telefono5TipoOtro").Value = IIf(Trim(Telefono5TipoOtro) = "", Null, Telefono5TipoOtro)
    mrecData("Telefono5Area").Value = IIf(Trim(Telefono5Area) = "", Null, Telefono5Area)
    mrecData("Telefono5Numero").Value = IIf(Trim(Telefono5Numero) = "", Null, Telefono5Numero)
    mrecData("IDContactoGrupo").Value = IDContactoGrupo
    mrecData("DomicilioLaboralCalle1").Value = IIf(Trim(DomicilioLaboralCalle1) = "", Null, DomicilioLaboralCalle1)
    mrecData("DomicilioLaboralNumero").Value = IIf(Trim(DomicilioLaboralNumero) = "", Null, DomicilioLaboralNumero)
    mrecData("DomicilioLaboralPiso").Value = IIf(Trim(DomicilioLaboralPiso) = "", Null, DomicilioLaboralPiso)
    mrecData("DomicilioLaboralDepartamento").Value = IIf(Trim(DomicilioLaboralDepartamento) = "", Null, DomicilioLaboralDepartamento)
    mrecData("DomicilioLaboralCalle2").Value = IIf(Trim(DomicilioLaboralCalle2) = "", Null, DomicilioLaboralCalle2)
    mrecData("DomicilioLaboralCalle3").Value = IIf(Trim(DomicilioLaboralCalle3) = "", Null, DomicilioLaboralCalle3)
    mrecData("DomicilioLaboralCodigoPostal").Value = IIf(Trim(DomicilioLaboralCodigoPostal) = "", Null, DomicilioLaboralCodigoPostal)
    mrecData("DomicilioLaboralIDProvincia").Value = IIf(Trim(DomicilioLaboralIDProvincia) = "", Null, DomicilioLaboralIDProvincia)
    mrecData("DomicilioLaboralIDLocalidad").Value = IIf(DomicilioLaboralIDLocalidad = 0, Null, DomicilioLaboralIDLocalidad)
    mrecData("DomicilioLaboralMailing").Value = DomicilioLaboralMailing
    mrecData("DomicilioParticularCalle1").Value = IIf(Trim(DomicilioParticularCalle1) = "", Null, DomicilioParticularCalle1)
    mrecData("DomicilioParticularNumero").Value = IIf(Trim(DomicilioParticularNumero) = "", Null, DomicilioParticularNumero)
    mrecData("DomicilioParticularPiso").Value = IIf(Trim(DomicilioParticularPiso) = "", Null, DomicilioParticularPiso)
    mrecData("DomicilioParticularDepartamento").Value = IIf(Trim(DomicilioParticularDepartamento) = "", Null, DomicilioParticularDepartamento)
    mrecData("DomicilioParticularCalle2").Value = IIf(Trim(DomicilioParticularCalle2) = "", Null, DomicilioParticularCalle2)
    mrecData("DomicilioParticularCalle3").Value = IIf(Trim(DomicilioParticularCalle3) = "", Null, DomicilioParticularCalle3)
    mrecData("DomicilioParticularCodigoPostal").Value = IIf(Trim(DomicilioParticularCodigoPostal) = "", Null, DomicilioParticularCodigoPostal)
    mrecData("DomicilioParticularIDProvincia").Value = IIf(Trim(DomicilioParticularIDProvincia) = "", Null, DomicilioParticularIDProvincia)
    mrecData("DomicilioParticularIDLocalidad").Value = IIf(DomicilioParticularIDLocalidad = 0, Null, DomicilioParticularIDLocalidad)
    mrecData("DomicilioParticularMailing").Value = DomicilioParticularMailing
    mrecData("DomicilioOtroNombre").Value = IIf(Trim(DomicilioOtroNombre) = "", Null, DomicilioOtroNombre)
    mrecData("DomicilioOtroCalle1").Value = IIf(Trim(DomicilioOtroCalle1) = "", Null, DomicilioOtroCalle1)
    mrecData("DomicilioOtroNumero").Value = IIf(Trim(DomicilioOtroNumero) = "", Null, DomicilioOtroNumero)
    mrecData("DomicilioOtroPiso").Value = IIf(Trim(DomicilioOtroPiso) = "", Null, DomicilioOtroPiso)
    mrecData("DomicilioOtroDepartamento").Value = IIf(Trim(DomicilioOtroDepartamento) = "", Null, DomicilioOtroDepartamento)
    mrecData("DomicilioOtroCalle2").Value = IIf(Trim(DomicilioOtroCalle2) = "", Null, DomicilioOtroCalle2)
    mrecData("DomicilioOtroCalle3").Value = IIf(Trim(DomicilioOtroCalle3) = "", Null, DomicilioOtroCalle3)
    mrecData("DomicilioOtroCodigoPostal").Value = IIf(Trim(DomicilioOtroCodigoPostal) = "", Null, DomicilioOtroCodigoPostal)
    mrecData("DomicilioOtroIDProvincia").Value = IIf(Trim(DomicilioOtroIDProvincia) = "", Null, DomicilioOtroIDProvincia)
    mrecData("DomicilioOtroIDLocalidad").Value = IIf(DomicilioOtroIDLocalidad = 0, Null, DomicilioOtroIDLocalidad)
    mrecData("DomicilioOtroMailing").Value = DomicilioOtroMailing
    mrecData("Email1").Value = IIf(Trim(Email1) = "", Null, Email1)
    mrecData("Email1Nombre").Value = IIf(Trim(Email1Nombre) = "", Null, Email1Nombre)
    mrecData("Email2").Value = IIf(Trim(Email2) = "", Null, Email2)
    mrecData("Email2Nombre").Value = IIf(Trim(Email2Nombre) = "", Null, Email2Nombre)
    mrecData("Email3").Value = IIf(Trim(Email3) = "", Null, Email3)
    mrecData("Email3Nombre").Value = IIf(Trim(Email3Nombre) = "", Null, Email3Nombre)
    mrecData("PaginaWeb").Value = IIf(Trim(PaginaWeb) = "", Null, PaginaWeb)
    mrecData("SobreNombre").Value = IIf(Trim(SobreNombre) = "", Null, SobreNombre)
    mrecData("FechaNacimiento").Value = IIf(FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE, Null, FechaNacimiento)
    mrecData("Asistente").Value = IIf(Trim(Asistente) = "", Null, Asistente)
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshContacto IDContacto
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Contacto con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Contacto.AddNew", "Error al Agregar el Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
        End If
    Else
        ShowErrorMessage "Classes.Contacto.AddNew", "Error al Agregar el Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function Update() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData("Titulo").Value = IIf(Trim(Titulo) = "", Null, Titulo)
    mrecData("Apellido").Value = IIf(Trim(Apellido) = "", Null, Apellido)
    mrecData("Nombre").Value = IIf(Trim(Nombre) = "", Null, Nombre)
    mrecData("Compania").Value = IIf(Trim(Compania) = "", Null, Compania)
    mrecData("TituloLaboral").Value = IIf(Trim(TituloLaboral) = "", Null, TituloLaboral)
    mrecData("IDTelefono1Tipo").Value = IIf(IDTelefono1Tipo = 0, Null, IDTelefono1Tipo)
    mrecData("Telefono1TipoOtro").Value = IIf(Trim(Telefono1TipoOtro) = "", Null, Telefono1TipoOtro)
    mrecData("Telefono1Area").Value = IIf(Trim(Telefono1Area) = "", Null, Telefono1Area)
    mrecData("Telefono1Numero").Value = IIf(Trim(Telefono1Numero) = "", Null, Telefono1Numero)
    mrecData("IDTelefono2Tipo").Value = IIf(IDTelefono2Tipo = 0, Null, IDTelefono2Tipo)
    mrecData("Telefono2TipoOtro").Value = IIf(Trim(Telefono2TipoOtro) = "", Null, Telefono2TipoOtro)
    mrecData("Telefono2Area").Value = IIf(Trim(Telefono2Area) = "", Null, Telefono2Area)
    mrecData("Telefono2Numero").Value = IIf(Trim(Telefono2Numero) = "", Null, Telefono2Numero)
    mrecData("IDTelefono3Tipo").Value = IIf(IDTelefono3Tipo = 0, Null, IDTelefono3Tipo)
    mrecData("Telefono3TipoOtro").Value = IIf(Trim(Telefono3TipoOtro) = "", Null, Telefono3TipoOtro)
    mrecData("Telefono3Area").Value = IIf(Trim(Telefono3Area) = "", Null, Telefono3Area)
    mrecData("Telefono3Numero").Value = IIf(Trim(Telefono3Numero) = "", Null, Telefono3Numero)
    mrecData("IDTelefono4Tipo").Value = IIf(IDTelefono4Tipo = 0, Null, IDTelefono4Tipo)
    mrecData("Telefono4TipoOtro").Value = IIf(Trim(Telefono4TipoOtro) = "", Null, Telefono4TipoOtro)
    mrecData("Telefono4Area").Value = IIf(Trim(Telefono4Area) = "", Null, Telefono4Area)
    mrecData("Telefono4Numero").Value = IIf(Trim(Telefono4Numero) = "", Null, Telefono4Numero)
    mrecData("IDTelefono5Tipo").Value = IIf(IDTelefono5Tipo = 0, Null, IDTelefono5Tipo)
    mrecData("Telefono5TipoOtro").Value = IIf(Trim(Telefono5TipoOtro) = "", Null, Telefono5TipoOtro)
    mrecData("Telefono5Area").Value = IIf(Trim(Telefono5Area) = "", Null, Telefono5Area)
    mrecData("Telefono5Numero").Value = IIf(Trim(Telefono5Numero) = "", Null, Telefono5Numero)
    mrecData("IDContactoGrupo").Value = IDContactoGrupo
    mrecData("DomicilioLaboralCalle1").Value = IIf(Trim(DomicilioLaboralCalle1) = "", Null, DomicilioLaboralCalle1)
    mrecData("DomicilioLaboralNumero").Value = IIf(Trim(DomicilioLaboralNumero) = "", Null, DomicilioLaboralNumero)
    mrecData("DomicilioLaboralPiso").Value = IIf(Trim(DomicilioLaboralPiso) = "", Null, DomicilioLaboralPiso)
    mrecData("DomicilioLaboralDepartamento").Value = IIf(Trim(DomicilioLaboralDepartamento) = "", Null, DomicilioLaboralDepartamento)
    mrecData("DomicilioLaboralCalle2").Value = IIf(Trim(DomicilioLaboralCalle2) = "", Null, DomicilioLaboralCalle2)
    mrecData("DomicilioLaboralCalle3").Value = IIf(Trim(DomicilioLaboralCalle3) = "", Null, DomicilioLaboralCalle3)
    mrecData("DomicilioLaboralCodigoPostal").Value = IIf(Trim(DomicilioLaboralCodigoPostal) = "", Null, DomicilioLaboralCodigoPostal)
    mrecData("DomicilioLaboralIDLocalidad").Value = IIf(DomicilioLaboralIDLocalidad = 0, Null, DomicilioLaboralIDLocalidad)
    mrecData("DomicilioLaboralIDProvincia").Value = IIf(Trim(DomicilioLaboralIDProvincia) = "", Null, DomicilioLaboralIDProvincia)
    mrecData("DomicilioLaboralMailing").Value = DomicilioLaboralMailing
    mrecData("DomicilioParticularCalle1").Value = IIf(Trim(DomicilioParticularCalle1) = "", Null, DomicilioParticularCalle1)
    mrecData("DomicilioParticularNumero").Value = IIf(Trim(DomicilioParticularNumero) = "", Null, DomicilioParticularNumero)
    mrecData("DomicilioParticularPiso").Value = IIf(Trim(DomicilioParticularPiso) = "", Null, DomicilioParticularPiso)
    mrecData("DomicilioParticularDepartamento").Value = IIf(Trim(DomicilioParticularDepartamento) = "", Null, DomicilioParticularDepartamento)
    mrecData("DomicilioParticularCalle2").Value = IIf(Trim(DomicilioParticularCalle2) = "", Null, DomicilioParticularCalle2)
    mrecData("DomicilioParticularCalle3").Value = IIf(Trim(DomicilioParticularCalle3) = "", Null, DomicilioParticularCalle3)
    mrecData("DomicilioParticularCodigoPostal").Value = IIf(Trim(DomicilioParticularCodigoPostal) = "", Null, DomicilioParticularCodigoPostal)
    mrecData("DomicilioParticularIDLocalidad").Value = IIf(DomicilioParticularIDLocalidad = 0, Null, DomicilioParticularIDLocalidad)
    mrecData("DomicilioParticularIDProvincia").Value = IIf(Trim(DomicilioParticularIDProvincia) = "", Null, DomicilioParticularIDProvincia)
    mrecData("DomicilioParticularMailing").Value = DomicilioParticularMailing
    mrecData("DomicilioOtroNombre").Value = IIf(Trim(DomicilioOtroNombre) = "", Null, DomicilioOtroNombre)
    mrecData("DomicilioOtroCalle1").Value = IIf(Trim(DomicilioOtroCalle1) = "", Null, DomicilioOtroCalle1)
    mrecData("DomicilioOtroNumero").Value = IIf(Trim(DomicilioOtroNumero) = "", Null, DomicilioOtroNumero)
    mrecData("DomicilioOtroPiso").Value = IIf(Trim(DomicilioOtroPiso) = "", Null, DomicilioOtroPiso)
    mrecData("DomicilioOtroDepartamento").Value = IIf(Trim(DomicilioOtroDepartamento) = "", Null, DomicilioOtroDepartamento)
    mrecData("DomicilioOtroCalle2").Value = IIf(Trim(DomicilioOtroCalle2) = "", Null, DomicilioOtroCalle2)
    mrecData("DomicilioOtroCalle3").Value = IIf(Trim(DomicilioOtroCalle3) = "", Null, DomicilioOtroCalle3)
    mrecData("DomicilioOtroCodigoPostal").Value = IIf(Trim(DomicilioOtroCodigoPostal) = "", Null, DomicilioOtroCodigoPostal)
    mrecData("DomicilioOtroIDProvincia").Value = IIf(Trim(DomicilioOtroIDProvincia) = "", Null, DomicilioOtroIDProvincia)
    mrecData("DomicilioOtroIDLocalidad").Value = IIf(DomicilioOtroIDLocalidad = 0, Null, DomicilioOtroIDLocalidad)
    mrecData("DomicilioOtroMailing").Value = DomicilioOtroMailing
    mrecData("Email1").Value = IIf(Trim(Email1) = "", Null, Email1)
    mrecData("Email1Nombre").Value = IIf(Trim(Email1Nombre) = "", Null, Email1Nombre)
    mrecData("Email2").Value = IIf(Trim(Email2) = "", Null, Email2)
    mrecData("Email2Nombre").Value = IIf(Trim(Email2Nombre) = "", Null, Email2Nombre)
    mrecData("Email3").Value = IIf(Trim(Email3) = "", Null, Email3)
    mrecData("Email3Nombre").Value = IIf(Trim(Email3Nombre) = "", Null, Email3Nombre)
    mrecData("PaginaWeb").Value = IIf(Trim(PaginaWeb) = "", Null, PaginaWeb)
    mrecData("SobreNombre").Value = IIf(Trim(SobreNombre) = "", Null, SobreNombre)
    mrecData("FechaNacimiento").Value = IIf(FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE, Null, FechaNacimiento)
    mrecData("Asistente").Value = IIf(Trim(Asistente) = "", Null, Asistente)
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshContacto IDContacto
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Contacto con el mismo Apellido, Nombre o Compañía.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Contacto.Update", "Error al Actualizar el Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
        End If
    Else
        ShowErrorMessage "Classes.Contacto.Update", "Error al Actualizar el Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    If RefreshList Then
        RefreshList_RefreshContacto 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el Contacto debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Contacto.Delete", "Error al Eliminar el Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
        End If
    Else
        ShowErrorMessage "Classes.Contacto.Delete", "Error al Eliminar el Contacto." & vbCr & vbCr & "IDContacto: " & IDContacto
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    Titulo = ""
    Apellido = ""
    Nombre = ""
    Compania = ""
    TituloLaboral = ""
    IDTelefono1Tipo = 0
    Telefono1TipoOtro = ""
    Telefono1Area = ""
    Telefono1Numero = ""
    IDTelefono2Tipo = 0
    Telefono2TipoOtro = ""
    Telefono2Area = ""
    Telefono2Numero = ""
    IDTelefono3Tipo = 0
    Telefono3TipoOtro = ""
    Telefono3Area = ""
    Telefono3Numero = ""
    IDTelefono4Tipo = 0
    Telefono4TipoOtro = ""
    Telefono4Area = ""
    Telefono4Numero = ""
    IDTelefono5Tipo = 0
    Telefono5TipoOtro = ""
    Telefono5Area = ""
    Telefono5Numero = ""
    IDContactoGrupo = 0
    DomicilioLaboralCalle1 = ""
    DomicilioLaboralNumero = ""
    DomicilioLaboralPiso = ""
    DomicilioLaboralDepartamento = ""
    DomicilioLaboralCalle2 = ""
    DomicilioLaboralCalle3 = ""
    DomicilioLaboralCodigoPostal = ""
    DomicilioLaboralIDProvincia = ""
    DomicilioLaboralIDLocalidad = 0
    DomicilioLaboralMailing = False
    DomicilioParticularCalle1 = ""
    DomicilioParticularNumero = ""
    DomicilioParticularPiso = ""
    DomicilioParticularDepartamento = ""
    DomicilioParticularCalle2 = ""
    DomicilioParticularCalle3 = ""
    DomicilioParticularCodigoPostal = ""
    DomicilioParticularIDProvincia = ""
    DomicilioParticularIDLocalidad = 0
    DomicilioParticularMailing = False
    DomicilioOtroNombre = ""
    DomicilioOtroCalle1 = ""
    DomicilioOtroNumero = ""
    DomicilioOtroPiso = ""
    DomicilioOtroDepartamento = ""
    DomicilioOtroCalle2 = ""
    DomicilioOtroCalle3 = ""
    DomicilioOtroCodigoPostal = ""
    DomicilioOtroIDProvincia = ""
    DomicilioOtroIDLocalidad = 0
    DomicilioOtroMailing = False
    Email1 = ""
    Email1Nombre = ""
    Email2 = ""
    Email2Nombre = ""
    Email3 = ""
    Email3Nombre = ""
    PaginaWeb = ""
    SobreNombre = ""
    FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE
    Asistente = ""
    Notas = ""
    Activo = True
    FechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioCreacion = 0
    FechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioModificacion = 0
End Sub

Private Sub Class_Initialize()
    InitializeValues
    NoMatchRaiseError = True
    RefreshList = True
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            If Not (mrecData.BOF Or mrecData.EOF) Then
                If mrecData.EditMode <> adEditNone Then
                    mrecData.CancelUpdate
                End If
            End If
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub

