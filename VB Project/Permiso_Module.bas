Attribute VB_Name = "Permiso_Module"
Option Explicit

Public Const PERMISO_OPCIONES_SYSTEM = "Opciones_Sistema"
Public Const PERMISO_OPCIONES_WORKSTATION = "Opciones_EstacionTrabajo"

Public Const PERMISO_USUARIO = "Usuario"
Public Const PERMISO_USUARIO_ADD = "Usuario_Agregar"
Public Const PERMISO_USUARIO_MODIFY = "Usuario_Modificar"
Public Const PERMISO_USUARIO_DELETE = "Usuario_Eliminar"

Public Const PERMISO_USUARIO_GRUPO = "Usuario_Grupo"
Public Const PERMISO_USUARIO_GRUPO_ADD = "Usuario_Grupo_Agregar"
Public Const PERMISO_USUARIO_GRUPO_MODIFY = "Usuario_Grupo_Modificar"
Public Const PERMISO_USUARIO_GRUPO_DELETE = "Usuario_Grupo_Eliminar"

Public Const PERMISO_USUARIO_GRUPO_PERMISSION = "Usuario_Grupo_Permiso"
Public Const PERMISO_USUARIO_GRUPO_PERMISSION_MODIFY = "Usuario_Grupo_Permiso_Modificar"

Public Const PERMISO_LUGAR = "Lugar"
Public Const PERMISO_LUGAR_ADD = "Lugar_Agregar"
Public Const PERMISO_LUGAR_MODIFY = "Lugar_Modificar"
Public Const PERMISO_LUGAR_DELETE = "Lugar_Eliminar"

Public Const PERMISO_LUGAR_GRUPO = "Lugar_Grupo"
Public Const PERMISO_LUGAR_GRUPO_ADD = "Lugar_Grupo_Agregar"
Public Const PERMISO_LUGAR_GRUPO_MODIFY = "Lugar_Grupo_Modificar"
Public Const PERMISO_LUGAR_GRUPO_DELETE = "Lugar_Grupo_Eliminar"

Public Const PERMISO_RUTA = "Ruta"
Public Const PERMISO_RUTA_ADD = "Ruta_Agregar"
Public Const PERMISO_RUTA_MODIFY = "Ruta_Modificar"
Public Const PERMISO_RUTA_DELETE = "Ruta_Eliminar"
Public Const PERMISO_RUTA_RUTA = "Ruta_Ruta_"

Public Const PERMISO_RUTA_DETALLE = "RutaDetalle"
Public Const PERMISO_RUTA_DETALLE_ADD = "RutaDetalle_Agregar"
Public Const PERMISO_RUTA_DETALLE_MODIFY = "RutaDetalle_Modificar"
Public Const PERMISO_RUTA_DETALLE_DELETE = "RutaDetalle_Eliminar"

Public Const PERMISO_RUTA_DETALLE_HORARIO = "RutaDetalle_Horario"
Public Const PERMISO_RUTA_DETALLE_HORARIO_ADD = "RutaDetalle_Horario_Agregar"
Public Const PERMISO_RUTA_DETALLE_HORARIO_MODIFY = "RutaDetalle_Horario_Modificar"
Public Const PERMISO_RUTA_DETALLE_HORARIO_DELETE = "RutaDetalle_Horario_Eliminar"

Public Const PERMISO_RUTALUGARGRUPO = "RutaLugarGrupo"
Public Const PERMISO_RUTALUGARGRUPO_ADD = "RutaLugarGrupo_Agregar"
Public Const PERMISO_RUTALUGARGRUPO_MODIFY = "RutaLugarGrupo_Modificar"
Public Const PERMISO_RUTALUGARGRUPO_DELETE = "RutaLugarGrupo_Eliminar"

Public Const PERMISO_LISTA_PRECIO = "Lista_Precio"
Public Const PERMISO_LISTA_PRECIO_ADD = "Lista_Precio_Agregar"
Public Const PERMISO_LISTA_PRECIO_MODIFY = "Lista_Precio_Modificar"
Public Const PERMISO_LISTA_PRECIO_DELETE = "Lista_Precio_Eliminar"
Public Const PERMISO_LISTA_PRECIO_LISTA_PRECIO = "Lista_Precio_Lista_Precio_"

Public Const PERMISO_LISTA_PRECIO_DETALLE = "Lista_Precio_Detalle"
Public Const PERMISO_LISTA_PRECIO_DETALLE_ADD = "Lista_Precio_Detalle_Agregar"
Public Const PERMISO_LISTA_PRECIO_DETALLE_MODIFY = "Lista_Precio_Detalle_Modificar"
Public Const PERMISO_LISTA_PRECIO_DETALLE_DELETE = "Lista_Precio_Detalle_Eliminar"

Public Const PERMISO_VEHICULO = "Vehiculo"
Public Const PERMISO_VEHICULO_ADD = "Vehiculo_Agregar"
Public Const PERMISO_VEHICULO_MODIFY = "Vehiculo_Modificar"
Public Const PERMISO_VEHICULO_DELETE = "Vehiculo_Eliminar"
Public Const PERMISO_VEHICULO_UTILIZACION = "Vehiculo_Utilizacion"

Public Const PERMISO_HORARIO = "Horario"
Public Const PERMISO_HORARIO_ADD = "Horario_Agregar"
Public Const PERMISO_HORARIO_MODIFY = "Horario_Modificar"
Public Const PERMISO_HORARIO_DELETE = "Horario_Eliminar"

Public Const PERMISO_PERSONA = "Persona"
Public Const PERMISO_PERSONA_ADD = "Persona_Agregar"
Public Const PERMISO_PERSONA_ADD_ALLTYPE = "Persona_Agregar_TodosTipos"
Public Const PERMISO_PERSONA_MODIFY = "Persona_Modificar"
Public Const PERMISO_PERSONA_DELETE = "Persona_Eliminar"
Public Const PERMISO_PERSONA_HABILITACION_VIAJAR_ESTABLECER = "Persona_Habilitacion_Viajar_Establecer"
Public Const PERMISO_PERSONA_HABILITACION_VIAJAR_IGNORAR = "Persona_Habilitacion_Viajar_Ignorar"

Public Const PERMISO_PERSONA_HORARIO = "Persona_Horario"
Public Const PERMISO_PERSONA_HORARIO_ADD = "Persona_Horario_Agregar"
Public Const PERMISO_PERSONA_HORARIO_MODIFY = "Persona_Horario_Modificar"
Public Const PERMISO_PERSONA_HORARIO_DELETE = "Persona_Horario_Eliminar"

Public Const PERMISO_PERSONA_RUTA = "Persona_Ruta"
Public Const PERMISO_PERSONA_RUTA_ADD = "Persona_Ruta_Agregar"
Public Const PERMISO_PERSONA_RUTA_MODIFY = "Persona_Ruta_Modificar"
Public Const PERMISO_PERSONA_RUTA_DELETE = "Persona_Ruta_Eliminar"

Public Const PERMISO_PERSONA_PREPAGO = "Persona_Prepago"
Public Const PERMISO_PERSONA_PREPAGO_ADD = "Persona_Prepago_Agregar"
Public Const PERMISO_PERSONA_PREPAGO_MODIFY = "Persona_Prepago_Modificar"
Public Const PERMISO_PERSONA_PREPAGO_DELETE = "Persona_Prepago_Eliminar"

Public Const PERMISO_PERSONA_INFO = "Persona_Informacion"

Public Const PERMISO_PERSONA_RESPUESTA = "Persona_Respuesta"
Public Const PERMISO_PERSONA_RESPUESTA_ADD = "Persona_Respuesta_Agregar"
Public Const PERMISO_PERSONA_RESPUESTA_MODIFY = "Persona_Respuesta_Modificar"
Public Const PERMISO_PERSONA_RESPUESTA_DELETE = "Persona_Respuesta_Eliminar"
Public Const PERMISO_PERSONA_RESPUESTA_ACTIVATE_ALL = "Persona_Respuesta_Activar_Todo"

Public Const PERMISO_CONDUCTOR_RUTA = "Conductor_Ruta"
Public Const PERMISO_CONDUCTOR_RUTA_ADD = "Conductor_Ruta_Agregar"
Public Const PERMISO_CONDUCTOR_RUTA_MODIFY = "Conductor_Ruta_Modificar"
Public Const PERMISO_CONDUCTOR_RUTA_DELETE = "Conductor_Ruta_Eliminar"

Public Const PERMISO_VIAJE = "Viaje"
Public Const PERMISO_VIAJE_HISTORICO = "Viaje_Historico"
Public Const PERMISO_VIAJE_GENERATE = "Viaje_Generar"
Public Const PERMISO_VIAJE_ADD = "Viaje_Agregar"
Public Const PERMISO_VIAJE_MODIFY = "Viaje_Modificar"
Public Const PERMISO_VIAJE_MODIFY_HORA = "Viaje_Modificar_Hora"
Public Const PERMISO_VIAJE_MODIFY_FECHA = "Viaje_Modificar_Fecha"
Public Const PERMISO_VIAJE_MODIFY_RUTA = "Viaje_Modificar_Ruta"
Public Const PERMISO_VIAJE_DELETE = "Viaje_Eliminar"
Public Const PERMISO_VIAJE_CHANGE_STATUS = "Viaje_Cambiar_Estado"
Public Const PERMISO_VIAJE_CHANGE_STATUS_CANCEL = "Viaje_Cambiar_Estado_Cancelar"
Public Const PERMISO_VIAJE_CHANGE_STATUS_SPECIAL = "Viaje_Cambiar_Estado_Todos"
Public Const PERMISO_VIAJE_SEND_EMAIL = "Viaje_Enviar_Email"

Public Const PERMISO_VIAJE_CONDUCTOR = "Viaje_Conductor"

Public Const PERMISO_VIAJE_ASISTENCIA = "Viaje_Asistencia"

Public Const PERMISO_VIAJE_DETALLE = "Viaje_Detalle"
Public Const PERMISO_VIAJE_DETALLE_HISTORICO = "Viaje_Detalle_Historico"
Public Const PERMISO_VIAJE_DETALLE_ADD = "Viaje_Detalle_Agregar"
Public Const PERMISO_VIAJE_DETALLE_ADD_FECHAHORARUTA = "Viaje_Detalle_Agregar_FechaHoraRuta"
Public Const PERMISO_VIAJE_DETALLE_MODIFY = "Viaje_Detalle_Modificar"
Public Const PERMISO_VIAJE_DETALLE_MODIFY_FECHAHORARUTA = "Viaje_Detalle_Modificar_FechaHoraRuta"
Public Const PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO = "Viaje_Detalle_Modificar_Importe_Otro_Usuario"
Public Const PERMISO_VIAJE_DETALLE_COMISION_MODIFY_IMPORTE_PAGOCONTADO_FINALIZADO = "Viaje_Detalle_Modificar_Importe_PagoContado_Finalizado"
Public Const PERMISO_VIAJE_DETALLE_COMISION_RENDIR = "Viaje_Detalle_Comision_Rendir"
Public Const PERMISO_VIAJE_DETALLE_PASAJERO_IMPORTE_ALLOWMODIFY = "Viaje_Detalle_Pasajero_Importe_AllowModify"
Public Const PERMISO_VIAJE_DETALLE_COMISION_IMPORTE_ALLOWMODIFY = "Viaje_Detalle_Comision_Importe_AllowModify"
Public Const PERMISO_VIAJE_DETALLE_PAQUETE_PASAJERO_IMPORTE_ALLOWMODIFY = "Viaje_Detalle_Paquete_Pasajero_Importe_AllowModify"
Public Const PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO = "Viaje_Detalle_Modificar_Caja_Otro_Usuario"
Public Const PERMISO_VIAJE_DETALLE_INASISTENCIA_NODEBITAR = "Viaje_Detalle_Inasistencia_NoDebitar"
Public Const PERMISO_VIAJE_DETALLE_DELETE = "Viaje_Detalle_Eliminar"
Public Const PERMISO_VIAJE_DETALLE_PRINT = "Viaje_Detalle_Imprimir"
Public Const PERMISO_VIAJE_DETALLE_CHANGE_STATUS = "Viaje_Detalle_Cambiar_Estado"
Public Const PERMISO_VIAJE_DETALLE_WEB_CHANGE_STATUS = "Viaje_Detalle_Web_Cambiar_Estado"
Public Const PERMISO_VIAJE_DETALLE_CHANGE_STATUS_AFTER_LIMIT = "Viaje_Detalle_Cambiar_Estado_Despues_Limite"
Public Const PERMISO_VIAJE_DETALLE_TRANSFER = "Viaje_Detalle_Transferencia"
Public Const PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWOUTOFRANGE = "Viaje_Detalle_Transferencia_Permite_FueraRango"
Public Const PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWDIFFERENTROUTE = "Viaje_Detalle_Transferencia_Permite_RutaDiferente"
Public Const PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY = "Viaje_Detalle_Finalizado_Modificar"

Public Const PERMISO_DOCUMENTOFISCAL = "DocumentoFiscal"
Public Const PERMISO_DOCUMENTOFISCAL_ADD = "DocumentoFiscal_Agregar"
Public Const PERMISO_DOCUMENTOFISCAL_MODIFY = "DocumentoFiscal_Modificar"
Public Const PERMISO_DOCUMENTOFISCAL_DELETE = "DocumentoFiscal_Eliminar"
Public Const PERMISO_DOCUMENTOFISCAL_PRINT = "DocumentoFiscal_Imprimir"

Public Const PERMISO_REPORTE = "Reporte"
Public Const PERMISO_REPORTE_REPORTE = "Reporte_"

Public Const PERMISO_CUENTA_CORRIENTE = "Cuenta_Corriente"
Public Const PERMISO_CUENTA_CORRIENTE_HISTORICO = "Cuenta_Corriente_Historico"
Public Const PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT = "Cuenta_Corriente_Conductor_Select"
Public Const PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT = "Cuenta_Corriente_Administrativo_Select"
Public Const PERMISO_CUENTA_CORRIENTE_ADD_ANTERIOR = "Cuenta_Corriente_Agregar_Anterior"
Public Const PERMISO_CUENTA_CORRIENTE_ADD_ACTUAL = "Cuenta_Corriente_Agregar_Actual"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_ANTERIOR = "Cuenta_Corriente_Modificar_Anterior"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_ACTUAL = "Cuenta_Corriente_Modificar_Actual"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_TRANSFER_ANTERIOR = "Cuenta_Corriente_Modificar_Transferencia_Anterior"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_TRANSFER_ACTUAL = "Cuenta_Corriente_Modificar_Transferencia_Actual"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_CAJA_OFICINA = "Cuenta_Corriente_Modificar_Caja_Oficina"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_NOTES_ANTERIOR = "Cuenta_Corriente_Modificar_Notas_Anterior"
Public Const PERMISO_CUENTA_CORRIENTE_MODIFY_NOTES_ACTUAL = "Cuenta_Corriente_Modificar_Notas_Actual"
Public Const PERMISO_CUENTA_CORRIENTE_DELETE = "Cuenta_Corriente_Eliminar"
Public Const PERMISO_CUENTA_CORRIENTE_DELETE_DEBITO_CREDITO = "Cuenta_Corriente_Eliminar_Debito_Credito"

Public Const PERMISO_CUENTA_CORRIENTE_GRUPO = "Cuenta_Corriente_Grupo"
Public Const PERMISO_CUENTA_CORRIENTE_GRUPO_ADD = "Cuenta_Corriente_Grupo_Agregar"
Public Const PERMISO_CUENTA_CORRIENTE_GRUPO_MODIFY = "Cuenta_Corriente_Grupo_Modificar"
Public Const PERMISO_CUENTA_CORRIENTE_GRUPO_DELETE = "Cuenta_Corriente_Grupo_Eliminar"
Public Const PERMISO_CUENTA_CORRIENTE_GRUPO_HIDDEN_SHOW = "Cuenta_Corriente_Grupo_Oculto_Mostrar"
Public Const PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO = "CuentaCorrienteGrupo_"

Public Const PERMISO_CUENTA_CORRIENTE_CAJA = "Cuenta_Corriente_Caja"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_ADD = "Cuenta_Corriente_Caja_Agregar"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_MODIFY = "Cuenta_Corriente_Caja_Modificar"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_DELETE = "Cuenta_Corriente_Caja_Eliminar"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER = "Cuenta_Corriente_Caja_Transferencia"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER_GRUPO_CAMBIAR = "Cuenta_Corriente_Caja_Transferencia_Grupo_Cambiar"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_SALDOOCULTO_VIEW = "Cuenta_Corriente_Caja_SaldoOculto_Ver"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL = "Cuenta_Corriente_Caja_Ver_Todos"
Public Const PERMISO_CUENTA_CORRIENTE_CAJA_CAJA = "CuentaCorrienteCaja_"

Public Const PERMISO_MEDIOPAGO = "MedioPago"
Public Const PERMISO_MEDIOPAGO_ADD = "MedioPago_Agregar"
Public Const PERMISO_MEDIOPAGO_MODIFY = "MedioPago_Modificar"
Public Const PERMISO_MEDIOPAGO_DELETE = "MedioPago_Eliminar"

Public Const PERMISO_FERIADO = "Feriado"
Public Const PERMISO_FERIADO_ADD = "Feriado_Agregar"
Public Const PERMISO_FERIADO_MODIFY = "Feriado_Modificar"
Public Const PERMISO_FERIADO_DELETE = "Feriado_Eliminar"

Public Const PERMISO_VEHICULO_MANTENIMIENTO = "Vehiculo_Mantenimiento"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_ADD = "Vehiculo_Mantenimiento_Agregar"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_MODIFY = "Vehiculo_Mantenimiento_Modificar"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_DELETE = "Vehiculo_Mantenimiento_Eliminar"

Public Const PERMISO_VEHICULO_MANTENIMIENTO_GRUPO = "Vehiculo_Mantenimiento_Grupo"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_GRUPO_ADD = "Vehiculo_Mantenimiento_Grupo_Agregar"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_GRUPO_MODIFY = "Vehiculo_Mantenimiento_Grupo_Modificar"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_GRUPO_DELETE = "Vehiculo_Mantenimiento_Grupo_Eliminar"

Public Const PERMISO_VEHICULO_MANTENIMIENTO_ACCION = "Vehiculo_Mantenimiento_Accion"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_ACCION_HISTORICO = "Vehiculo_Mantenimiento_Accion_Historico"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_ACCION_ADD = "Vehiculo_Mantenimiento_Accion_Agregar"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_ACCION_MODIFY = "Vehiculo_Mantenimiento_Accion_Modificar"
Public Const PERMISO_VEHICULO_MANTENIMIENTO_ACCION_DELETE = "Vehiculo_Mantenimiento_Accion_Eliminar"

Public Const PERMISO_DOCUMENTO_TIPO = "Documento_Tipo"
Public Const PERMISO_DOCUMENTO_TIPO_ADD = "Documento_Tipo_Agregar"
Public Const PERMISO_DOCUMENTO_TIPO_MODIFY = "Documento_Tipo_Modificar"
Public Const PERMISO_DOCUMENTO_TIPO_DELETE = "Documento_Tipo_Eliminar"

Public Const PERMISO_TELEFONO_TIPO = "Telefono_Tipo"
Public Const PERMISO_TELEFONO_TIPO_ADD = "Telefono_Tipo_Agregar"
Public Const PERMISO_TELEFONO_TIPO_MODIFY = "Telefono_Tipo_Modificar"
Public Const PERMISO_TELEFONO_TIPO_DELETE = "Telefono_Tipo_Eliminar"

Public Const PERMISO_PERSONA_ALARMA = "Persona_Alarma"
Public Const PERMISO_PERSONA_ALARMA_ADD = "Persona_Alarma_Agregar"
Public Const PERMISO_PERSONA_ALARMA_MODIFY = "Persona_Alarma_Modificar"
Public Const PERMISO_PERSONA_ALARMA_DELETE = "Persona_Alarma_Eliminar"

Public Const PERMISO_PERSONA_ALARMA_GRUPO = "Persona_Alarma_Grupo"
Public Const PERMISO_PERSONA_ALARMA_GRUPO_ADD = "Persona_Alarma_Grupo_Agregar"
Public Const PERMISO_PERSONA_ALARMA_GRUPO_MODIFY = "Persona_Alarma_Grupo_Modificar"
Public Const PERMISO_PERSONA_ALARMA_GRUPO_DELETE = "Persona_Alarma_Grupo_Eliminar"

Public Const PERMISO_ALARMA = "Alarma"
Public Const PERMISO_ALARMA_ADD = "Alarma_Agregar"
Public Const PERMISO_ALARMA_MODIFY = "Alarma_Modificar"
Public Const PERMISO_ALARMA_DELETE = "Alarma_Eliminar"

Public Const PERMISO_CONTACTO = "Contacto"
Public Const PERMISO_CONTACTO_ADD = "Contacto_Agregar"
Public Const PERMISO_CONTACTO_MODIFY = "Contacto_Modificar"
Public Const PERMISO_CONTACTO_DELETE = "Contacto_Eliminar"

Public Const PERMISO_CONTACTO_GRUPO = "Contacto_Grupo"
Public Const PERMISO_CONTACTO_GRUPO_ADD = "Contacto_Grupo_Agregar"
Public Const PERMISO_CONTACTO_GRUPO_MODIFY = "Contacto_Grupo_Modificar"
Public Const PERMISO_CONTACTO_GRUPO_DELETE = "Contacto_Grupo_Eliminar"

Public Const PERMISO_LISTAPASAJERO = "Lista_Pasajero"

Public Const PERMISO_REGISTROLLAMADA = "RegistroLlamada"

Public Const PERMISO_FRANCO = "Franco"
Public Const PERMISO_FRANCO_ADD = "Franco_Agregar"
Public Const PERMISO_FRANCO_MODIFY = "Franco_Modificar"
Public Const PERMISO_FRANCO_DELETE = "Franco_Eliminar"

Public Const PERMISO_APPEXT_MSWORD = "Permiso_AplicacionExterna_MicrosoftWord"
Public Const PERMISO_APPEXT_MSEXCEL = "Permiso_AplicacionExterna_MicrosoftExcel"

Public Const PERMISO_UTILIDAD = "Utilidad"
Public Const PERMISO_UTILIDAD_EDITOR_DIRECTO_PARAMETRO = "Utilidad_Editor_Directo_Parametro"
Public Const PERMISO_UTILIDAD_ACTUALIZAR_PRECIO_RESERVA = "Utilidad_Actualizar_Precio_Reserva"
Public Const PERMISO_UTILIDAD_ACTUALIZAR_SUELDO_VIAJE = "Utilidad_Actualizar_Sueldo_Viaje"

Public Const PERMISO_MENSAJE = "Mensaje"
Public Const PERMISO_MENSAJE_ADD = "Mensaje_Agregar"
Public Const PERMISO_MENSAJE_MODIFY = "Mensaje_Modificar"
Public Const PERMISO_MENSAJE_DELETE = "Mensaje_Eliminar"

Public Const PERMISO_MESSENGER = "Messenger"
