VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSemaforoGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCSemaforoGeneral As Collection

Public Property Get Item(IndexKey As Variant) As Semaforo
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = mCSemaforoGeneral(IndexKey)
End Property

Public Property Get Count() As Long
    Count = mCSemaforoGeneral.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCSemaforoGeneral.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCSemaforoGeneral = New Collection
    
    Add MODULE_USUARIO
    Add MODULE_USUARIO_GRUPO
    Add MODULE_USUARIO_GRUPO_PERMISO
    Add MODULE_LUGAR
    Add MODULE_LUGAR_GRUPO
    Add MODULE_RUTA
    Add MODULE_RUTA_DETALLE
    Add MODULE_RUTA_DETALLE_HORARIO
    Add MODULE_RUTA_LUGARGRUPO
    Add MODULE_VEHICULO
    Add MODULE_VEHICULO_UTILIZACION
    Add MODULE_LISTA_PRECIO
    Add MODULE_LISTA_PRECIO_DETALLE
    Add MODULE_HORARIO
    Add MODULE_PERSONA
    Add MODULE_PERSONA_HORARIO
    Add MODULE_PERSONA_RUTA
    Add MODULE_PERSONA_PREPAGO
    Add MODULE_PERSONA_RESPUESTA
    Add MODULE_CONDUCTOR_RUTA
    Add MODULE_VIAJE
    Add MODULE_VIAJE_DETALLE
    Add MODULE_CUENTA_CORRIENTE
    Add MODULE_CUENTA_CORRIENTE_CAJA
    Add MODULE_CUENTA_CORRIENTE_GRUPO
    Add MODULE_MEDIOPAGO
    Add MODULE_FERIADO
    Add MODULE_VEHICULO_MANTENIMIENTO
    Add MODULE_VEHICULO_MANTENIMIENTO_GRUPO
    Add MODULE_VEHICULO_MANTENIMIENTO_ACCION
    Add MODULE_DOCUMENTO_TIPO
    Add MODULE_TELEFONO_TIPO
    Add MODULE_PERSONA_ALARMA
    Add MODULE_PERSONA_ALARMA_GRUPO
    Add MODULE_ALARMA
    Add MODULE_CONTACTO
    Add MODULE_CONTACTO_GRUPO
    Add MODULE_REGISTROLLAMADA
    Add MODULE_FRANCO
    Add MODULE_MENSAJE
    Add MODULE_PERSONAL
    Add MODULE_MESSENGER
End Sub

Private Sub Add(ByVal IDSemaforo As String)
    Dim Semaforo As Semaforo
    
    Set Semaforo = New Semaforo
    With Semaforo
        .IDSemaforo = IDSemaforo
        .ValorTimer = -1
    End With
    mCSemaforoGeneral.Add Semaforo, Semaforo.IDSemaforo
    
    Set Semaforo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mCSemaforoGeneral = Nothing
End Sub
