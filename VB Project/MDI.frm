VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm frmMDI 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Title"
   ClientHeight    =   5970
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20250
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMulti 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   2760
   End
   Begin MSCommLib.MSComm comTelephony 
      Left            =   4860
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5595
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25788
            MinWidth        =   1764
            Picture         =   "MDI.frx":08CA
            Key             =   "USUARIO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3545
            MinWidth        =   2
            Picture         =   "MDI.frx":0F04
            Text            =   "Centro de Mensajes"
            TextSave        =   "Centro de Mensajes"
            Key             =   "CENTRO_MENSAJE"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
            Key             =   "COMPANY_NAME"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   794
            MinWidth        =   794
            Text            =   "OFF"
            TextSave        =   "OFF"
            Key             =   "PERSONAL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   2
            TextSave        =   "CAPS"
            Key             =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   741
            MinWidth        =   2
            TextSave        =   "NUM"
            Key             =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   582
            MinWidth        =   2
            TextSave        =   "INS"
            Key             =   "INS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   2
            TextSave        =   "20/06/2020"
            Key             =   "DATE"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   2
            TextSave        =   "20:17"
            Key             =   "TIME"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList ilsToolbar 
      Left            =   4260
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":121E
            Key             =   "DATOS_INICIALES"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1AF8
            Key             =   "OTROS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":23D2
            Key             =   "PERSONA"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":2A0C
            Key             =   "VIAJE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3048
            Key             =   "TRANSFERENCIA_BOTH"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3682
            Key             =   "COMISION"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3CBC
            Key             =   "CUENTA_CORRIENTE"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":42FC
            Key             =   "TRANSFERENCIA"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4896
            Key             =   "REPORTE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4A70
            Key             =   "MESSENGER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1058
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tablas"
            Key             =   "DATOS_INICIALES"
            Object.ToolTipText     =   "Tablas de Datos Iniciales"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   35
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "USUARIO_GRUPO"
                  Text            =   "Grupos de Usuarios"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "USUARIO"
                  Text            =   "Usuarios"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LUGAR_GRUPO"
                  Text            =   "Grupos de Lugares"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LUGAR"
                  Text            =   "Lugares"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RUTA"
                  Text            =   "Rutas"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RUTA_DETALLE"
                  Text            =   "Detalle de Rutas"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RUTA_LUGARGRUPO"
                  Text            =   "Rutas-Grupos de Lugares"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CONDUCTOR_RUTA"
                  Text            =   "Rutas por Conductor"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LISTA_PRECIO"
                  Text            =   "Listas de Precios"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VEHICULO"
                  Text            =   "Vehículos"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VEHICULO_MANTENIMIENTO_GRUPO"
                  Text            =   "Grupos de Mantenimiento de Vehículos"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VEHICULO_MANTENIMIENTO"
                  Text            =   "Mantenimiento de Vehículos"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HORARIO"
                  Text            =   "Horarios"
               EndProperty
               BeginProperty ButtonMenu20 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu21 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CUENTACORRIENTE_CAJA"
                  Text            =   "Cajas de Cuenta Corriente"
               EndProperty
               BeginProperty ButtonMenu22 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CUENTACORRIENTE_GRUPO"
                  Text            =   "Grupos de Cuenta Corriente"
               EndProperty
               BeginProperty ButtonMenu23 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MEDIOPAGO"
                  Text            =   "Medios de Pago"
               EndProperty
               BeginProperty ButtonMenu24 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu25 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FERIADO"
                  Text            =   "Feriados"
               EndProperty
               BeginProperty ButtonMenu26 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu27 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DOCUMENTO_TIPO"
                  Text            =   "Tipos de Documento"
               EndProperty
               BeginProperty ButtonMenu28 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TELEFONO_TIPO"
                  Text            =   "Tipos de Teléfono"
               EndProperty
               BeginProperty ButtonMenu29 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu30 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PERSONA_ALARMA_GRUPO"
                  Text            =   "Grupos de Alarmas de Personas"
               EndProperty
               BeginProperty ButtonMenu31 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PERSONA_ALARMA"
                  Text            =   "Alarmas de Personas"
               EndProperty
               BeginProperty ButtonMenu32 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu33 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ALARMA"
                  Text            =   "Alarmas Generales"
               EndProperty
               BeginProperty ButtonMenu34 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu35 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CONTACTO_GRUPO"
                  Text            =   "Grupos de Contactos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEPARATOR_1"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros"
            Key             =   "OTROS"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   22
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VEHICULO_MANTENIMIENTO_ACCION"
                  Text            =   "Acciones de Mantenimiento de Vehículos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VEHICULO_MANTENIMIENTO_ACCION_HISTORICO"
                  Text            =   "Acciones de Mantenimiento de Vehículos (Histórico)"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CONTACTO"
                  Text            =   "Contactos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VIAJE_TRANSFERENCIA"
                  Text            =   "Transferencia de Reservas"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VIAJE_ASISTENCIA"
                  Text            =   "Dar Asistencia a Viajes"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VIAJE_CONDUCTOR"
                  Text            =   "Viajes por Conductor"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VEHICULO_UTILIZACION"
                  Text            =   "Utilización de Vehículos"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LISTAPASAJERO"
                  Text            =   "Generar Lista de Pasajeros"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "REGISTROLLAMADA"
                  Text            =   "Registro de Llamadas"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FRANCOS"
                  Text            =   "Francos"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MENSAJES"
                  Text            =   "Mensajes"
               EndProperty
               BeginProperty ButtonMenu20 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu21 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "APPEXT_MSWORD"
                  Text            =   "Microsoft Word"
               EndProperty
               BeginProperty ButtonMenu22 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "APPEXT_MSEXCEL"
                  Text            =   "Microsoft Excel"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEPARATOR_2"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Personas"
            Key             =   "PERSONA"
            Object.ToolTipText     =   "Clientes, Administrativos, Conductores"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEPARATOR_3"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Viajes"
            Key             =   "VIAJE"
            Object.ToolTipText     =   "Viajes"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACTUAL"
                  Text            =   "Actuales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HISTORICO"
                  Text            =   "Históricos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Comisión"
            Key             =   "COMISION"
            Object.ToolTipText     =   "Comisiones"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACTUAL"
                  Text            =   "Actuales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HISTORICO"
                  Text            =   "Históricas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEPARATOR_4"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cta. Cte."
            Key             =   "CUENTA_CORRIENTE"
            Object.ToolTipText     =   "Movimientos de Cuenta Corriente"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACTUAL"
                  Text            =   "Actuales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HISTORICO"
                  Text            =   "Históricas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transfer."
            Key             =   "CUENTA_CORRIENTE_TRANSFERENCIA"
            Object.ToolTipText     =   "Transferencias de Caja"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEPARATOR_5"
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            Key             =   "REPORTE"
            Object.ToolTipText     =   "Visor de Reportes"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEPARATOR_6"
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Messenger"
            Key             =   "MESSENGER"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "SEPARATOR_7"
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox picPersona 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   11040
         ScaleHeight     =   420
         ScaleWidth      =   3795
         TabIndex        =   4
         Top             =   60
         Width           =   3795
         Begin VB.CommandButton cmdPersona 
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
            Left            =   3540
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Personas"
            Top             =   60
            Width           =   255
         End
         Begin VB.ComboBox cboPersona 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   60
            Width           =   2895
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo:"
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
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   465
         End
      End
      Begin VB.PictureBox picCallerID 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   15060
         ScaleHeight     =   405
         ScaleWidth      =   3225
         TabIndex        =   1
         Top             =   60
         Width           =   3225
         Begin VB.TextBox txtCallerIDTipo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   60
            Width           =   975
         End
         Begin VB.TextBox txtCallerID 
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
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label lblCallerID 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo N°:"
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
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   675
         End
      End
   End
   Begin MSComctlLib.ImageList ilsFormToolbarHot 
      Left            =   4260
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":515A
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":56B6
            Key             =   "PROPERTIES"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":5C12
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":616E
            Key             =   "SELECT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":67AA
            Key             =   "DETAIL"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6DE6
            Key             =   "VIAJE_GENERATE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7422
            Key             =   "PERMISSION"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7A5C
            Key             =   "HORARIO"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8096
            Key             =   "RUTA"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8270
            Key             =   "INFO"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":88AA
            Key             =   "RESPUESTA"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8EE4
            Key             =   "ACTIVATE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":951E
            Key             =   "DEACTIVATE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9B58
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A192
            Key             =   "CHANGE_STATUS"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A7CC
            Key             =   "ASISTENCIA"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":AE06
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B440
            Key             =   "DOWN"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":BA7A
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":BFD4
            Key             =   "EMAIL"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":C60E
            Key             =   "PAGO_MULTIPLE"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":CC48
            Key             =   "RENDIR"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D282
            Key             =   "PREPAGO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsFormToolbar 
      Left            =   4860
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D8BC
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":DE18
            Key             =   "PROPERTIES"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":E374
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":E8D0
            Key             =   "SELECT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":EF0C
            Key             =   "DETAIL"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":F548
            Key             =   "VIAJE_GENERATE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":FB84
            Key             =   "PERMISSION"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":101BE
            Key             =   "HORARIO"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":107F8
            Key             =   "RUTA"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":109D2
            Key             =   "INFO"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1100C
            Key             =   "RESPUESTA"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":11646
            Key             =   "ACTIVATE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":11C80
            Key             =   "DEACTIVATE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":122BA
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":128F4
            Key             =   "CHANGE_STATUS"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":12F2E
            Key             =   "ASISTENCIA"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":13568
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":13BA2
            Key             =   "DOWN"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":141DC
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":14736
            Key             =   "EMAIL"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":14D70
            Key             =   "PAGO_MULTIPLE"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":153AA
            Key             =   "RENDIR"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":159E4
            Key             =   "PREPAGO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsFormPin 
      Left            =   6420
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1601E
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":165BA
            Key             =   "DOWN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsFormSortColumn 
      Left            =   5640
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":16B56
            Key             =   "ASC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":16C28
            Key             =   "DESC"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileChangePassword 
         Caption         =   "Cambiar Contraseña"
      End
      Begin VB.Menu mnuFileOptionSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Opciones"
         Begin VB.Menu mnuFileOptionSystem 
            Caption         =   "Sistema"
         End
         Begin VB.Menu mnuFileOptionWorkstation 
            Caption         =   "Estación de Trabajo"
         End
         Begin VB.Menu mnuFileOptionUser 
            Caption         =   "Usuario"
         End
      End
      Begin VB.Menu mnuFileCloseSessionSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseSession 
         Caption         =   "Cerrar Sesión del Usuario"
      End
      Begin VB.Menu mnuFileExitSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuUtilityParameter 
         Caption         =   "Editor directo de Parámetros"
      End
      Begin VB.Menu mnuUtilityCallerIDSimulate 
         Caption         =   "Simular Identificación de Llamada"
      End
      Begin VB.Menu mnuUtilityExecute 
         Caption         =   "Ejecutar..."
      End
      Begin VB.Menu mnuUtilityHistorySeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilityHistory 
         Caption         =   "Pasaje de Datos a Histórico"
      End
      Begin VB.Menu mnuUtilityUpdatePrecioSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilityUpdatePrecio 
         Caption         =   "Actualizar Precios de Reservas"
      End
      Begin VB.Menu mnuUtilityUpdateSueldoSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilityUpdateSueldo 
         Caption         =   "Actualizar Sueldos de Viaje"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowTileHorizontally 
         Caption         =   "Mosaico &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertically 
         Caption         =   "Mosaico &Vertical"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Organizar Iconos"
      End
      Begin VB.Menu mnuWindowSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowReset 
         Caption         =   "&Restaurar Tamaño"
      End
      Begin VB.Menu mnuWindowClose 
         Caption         =   "C&errar"
      End
      Begin VB.Menu mnuWindowCloseAll 
         Caption         =   "Cerrar Todas"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCTelefonoTipoNombre As Collection

Private WithEvents mTAPI As TAPI
Attribute mTAPI.VB_VarHelpID = -1
Private mCallInfo As ITCallInfo

Public Sub SetTAPIObject()
    Set mTAPI = Nothing
    Set mTAPI = pTelephony.TAPI
End Sub

Private Sub MDIForm_Load()
    Caption = App.Title
    
    Set mCTelefonoTipoNombre = New Collection
    
    mnuHelpAbout.Caption = "&Acerca de " & App.Title & "..."
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = ilsToolbar
    tlbMain.Buttons("DATOS_INICIALES").Image = "DATOS_INICIALES"
    tlbMain.Buttons("OTROS").Image = "OTROS"
    tlbMain.Buttons("PERSONA").Image = "PERSONA"
    tlbMain.Buttons("VIAJE").Image = "VIAJE"
    tlbMain.Buttons("COMISION").Image = "COMISION"
    tlbMain.Buttons("CUENTA_CORRIENTE").Image = "CUENTA_CORRIENTE"
    tlbMain.Buttons("CUENTA_CORRIENTE_TRANSFERENCIA").Image = "TRANSFERENCIA"
    tlbMain.Buttons("REPORTE").Image = "REPORTE"
    tlbMain.Buttons("MESSENGER").Image = "MESSENGER"
        
    stbMain.Panels("PERSONAL").Text = IIf(pPersonal, "ON", "OFF")
End Sub

Private Sub MDIForm_Resize()
    If picCallerID.Visible Then
        picCallerID.Left = tlbMain.Width - picCallerID.Width - 60
        picPersona.Left = picCallerID.Left - picPersona.Width - 180
    Else
        picPersona.Left = tlbMain.Width - picPersona.Width - 60
    End If
    
    CSM_Forms.ResizeAndPositionAll frmMDI
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbAppWindows And UnloadMode <> vbFormCode And pIsCompiled Then
        If MsgBox("¿Desea salir de la Aplicación?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Cancel = True
        Else
            Screen.MousePointer = vbHourglass
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set mCTelefonoTipoNombre = Nothing
    
    Set mCallInfo = Nothing
    Set mTAPI = Nothing
    
    TerminateApplication
End Sub

Private Sub mnuFileChangePassword_Click()
    frmChangePassword.LoadDataAndShow
End Sub

Private Sub mnuFileCloseSession_Click()
    If MsgBox("¿Desea Cerrar la Sesión del Usuario: '" & pUsuario.Nombre & "'?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        pUsuario.LogOut
        
        Load frmLogin
        frmLogin.Show vbModal, frmMDI
        If pUsuario.IDUsuario = 0 Then
            Unload Me
            Exit Sub
        End If
        If Not pUsuario.LogIn() Then
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    If MsgBox("¿Desea salir de la Aplicación?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Screen.MousePointer = vbHourglass
        Unload Me
    End If
End Sub

Private Sub mnuFileOptionSystem_Click()
    If pCPermiso.GotPermission(PERMISO_OPCIONES_SYSTEM) Then
        Screen.MousePointer = vbHourglass
        frmOpcionSystem.Show
        If frmOpcionSystem.WindowState = vbMinimized Then
            frmOpcionSystem.WindowState = vbNormal
        End If
        frmOpcionSystem.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuFileOptionWorkstation_Click()
    If pCPermiso.GotPermission(PERMISO_OPCIONES_WORKSTATION) Then
        Screen.MousePointer = vbHourglass
        frmOpcionWorkstation.Show
        If frmOpcionWorkstation.WindowState = vbMinimized Then
            frmOpcionWorkstation.WindowState = vbNormal
        End If
        frmOpcionWorkstation.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuFileOptionUser_Click()
    Screen.MousePointer = vbHourglass
    frmOpcionUser.Show
    If frmOpcionUser.WindowState = vbMinimized Then
        frmOpcionUser.WindowState = vbNormal
    End If
    frmOpcionUser.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuHelpAbout_Click()
    Screen.MousePointer = vbHourglass
    frmAbout.Show vbModal, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuUtilityCallerIDSimulate_Click()
    Dim NumeroTelefono As String
    
    NumeroTelefono = InputBox("Ingrese el Número de Teléfono:", "Simulación de Identifiación de LLamada")
    If NumeroTelefono <> "" Then
        txtCallerID.Text = NumeroTelefono
        
        CallerID_BuscarPersonas
        CallerID_BuscarVehiculos
    End If
End Sub

Private Sub mnuUtilityExecute_Click()
    frmExecute.Show vbModal, frmMDI
End Sub

Private Sub mnuUtilityHistory_Click()
    frmPasajeHistorico.Show
End Sub

Private Sub mnuUtilityParameter_Click()
    frmParametro.Show
End Sub

Private Sub mnuUtilityUpdatePrecio_Click()
    frmReservaActualizarPrecio.Show
End Sub

Private Sub mnuUtilityUpdateSueldo_Click()
    frmViajeActualizarSueldo.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnuWindowClose_Click()
    If Not ActiveForm Is Nothing Then
        Unload ActiveForm
    End If
End Sub

Private Sub mnuWindowCloseAll_Click()
    CSM_Forms.UnloadAll "frmMDI"
End Sub

Private Sub mnuWindowReset_Click()
    If Not ActiveForm Is Nothing Then
        CSM_Forms.ResizeAndPosition frmMDI, ActiveForm
    End If
End Sub

Private Sub mnuWindowTileHorizontally_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertically_Click()
    Arrange vbTileVertical
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PERSONA"
            If pCPermiso.GotPermission(PERMISO_PERSONA) Then
                Screen.MousePointer = vbHourglass
                frmPersona.Show
                If frmPersona.WindowState = vbMinimized Then
                    frmPersona.WindowState = vbNormal
                End If
                frmPersona.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "VIAJE"
            If pCPermiso.GotPermission(PERMISO_VIAJE) Then
                Screen.MousePointer = vbHourglass
                frmViaje.Show
                If frmViaje.WindowState = vbMinimized Then
                    frmViaje.WindowState = vbNormal
                End If
                frmViaje.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "COMISION"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE) Then
                Screen.MousePointer = vbHourglass
                frmComision.Show
                If frmComision.WindowState = vbMinimized Then
                    frmComision.WindowState = vbNormal
                End If
                frmComision.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "CUENTA_CORRIENTE"
            Call tlbMain_ButtonMenuClick(tlbMain.Buttons("CUENTA_CORRIENTE").ButtonMenus("ACTUAL"))
        Case "CUENTA_CORRIENTE_TRANSFERENCIA"
            If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER) Then
                Screen.MousePointer = vbHourglass
                frmCuentaCorrienteTransferencia.Show
                If frmCuentaCorrienteTransferencia.WindowState = vbMinimized Then
                    frmCuentaCorrienteTransferencia.WindowState = vbNormal
                End If
                frmCuentaCorrienteTransferencia.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "REPORTE"
            If pCPermiso.GotPermission(PERMISO_REPORTE) Then
                Screen.MousePointer = vbHourglass
                frmReporte.Show
                If frmReporte.WindowState = vbMinimized Then
                    frmReporte.WindowState = vbNormal
                End If
                frmReporte.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "MESSENGER"
            If pMessengerEnabled Then
                Screen.MousePointer = vbHourglass
                pMessengerBlinking = False
                tlbMain.Buttons("MESSENGER").Image = "MESSENGER"
                ' LAUNCH THE MESSENGER APPLICATION ON REMOTE DESKTOP
                Screen.MousePointer = vbDefault
            End If
    End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Static MSExcelApp As Object
    Static MSWordApp As Object
    
    Select Case ButtonMenu.Parent.Key & "_" & ButtonMenu.Key
        '//////////////////////////////////////////////////////////////////
        'D A T O S   I N I C I A L E S
        '//////////////////////////////////////////////////////////////////
        Case "DATOS_INICIALES_USUARIO_GRUPO"
            If pCPermiso.GotPermission(PERMISO_USUARIO_GRUPO) Then
                Screen.MousePointer = vbHourglass
                frmUsuarioGrupo.Show
                If frmUsuarioGrupo.WindowState = vbMinimized Then
                    frmUsuarioGrupo.WindowState = vbNormal
                End If
                frmUsuarioGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_USUARIO"
            If pCPermiso.GotPermission(PERMISO_USUARIO) Then
                Screen.MousePointer = vbHourglass
                frmUsuario.Show
                If frmUsuario.WindowState = vbMinimized Then
                    frmUsuario.WindowState = vbNormal
                End If
                frmUsuario.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_LUGAR_GRUPO"
            If pCPermiso.GotPermission(PERMISO_LUGAR_GRUPO) Then
                Screen.MousePointer = vbHourglass
                frmLugarGrupo.Show
                If frmLugarGrupo.WindowState = vbMinimized Then
                    frmLugarGrupo.WindowState = vbNormal
                End If
                frmLugarGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_LUGAR"
            If pCPermiso.GotPermission(PERMISO_LUGAR) Then
                Screen.MousePointer = vbHourglass
                frmLugar.Show
                If frmLugar.WindowState = vbMinimized Then
                    frmLugar.WindowState = vbNormal
                End If
                frmLugar.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_RUTA"
            If pCPermiso.GotPermission(PERMISO_RUTA) Then
                Screen.MousePointer = vbHourglass
                frmRuta.Show
                If frmRuta.WindowState = vbMinimized Then
                    frmRuta.WindowState = vbNormal
                End If
                frmRuta.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_RUTA_DETALLE"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE) Then
                Screen.MousePointer = vbHourglass
                frmRutaDetalle.Show
                If frmRutaDetalle.WindowState = vbMinimized Then
                    frmRutaDetalle.WindowState = vbNormal
                End If
                frmRutaDetalle.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_RUTA_LUGARGRUPO"
            If pCPermiso.GotPermission(PERMISO_RUTALUGARGRUPO) Then
                Screen.MousePointer = vbHourglass
                frmRutaLugarGrupo.Show
                If frmRutaLugarGrupo.WindowState = vbMinimized Then
                    frmRutaLugarGrupo.WindowState = vbNormal
                End If
                frmRutaLugarGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_CONDUCTOR_RUTA"
            If pCPermiso.GotPermission(PERMISO_CONDUCTOR_RUTA) Then
                Screen.MousePointer = vbHourglass
                frmConductorRuta.Show
                If frmConductorRuta.WindowState = vbMinimized Then
                    frmConductorRuta.WindowState = vbNormal
                End If
                frmConductorRuta.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_LISTA_PRECIO"
            If pCPermiso.GotPermission(PERMISO_LISTA_PRECIO) Then
                Screen.MousePointer = vbHourglass
                frmListaPrecio.Show
                If frmListaPrecio.WindowState = vbMinimized Then
                    frmListaPrecio.WindowState = vbNormal
                End If
                frmListaPrecio.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_VEHICULO"
            If pCPermiso.GotPermission(PERMISO_VEHICULO) Then
                Screen.MousePointer = vbHourglass
                frmVehiculo.Show
                If frmVehiculo.WindowState = vbMinimized Then
                    frmVehiculo.WindowState = vbNormal
                End If
                frmVehiculo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_HORARIO"
            If pCPermiso.GotPermission(PERMISO_HORARIO) Then
                Screen.MousePointer = vbHourglass
                frmHorario.Show
                If frmHorario.WindowState = vbMinimized Then
                    frmHorario.WindowState = vbNormal
                End If
                frmHorario.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_CUENTACORRIENTE_GRUPO"
            If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO) Then
                Screen.MousePointer = vbHourglass
                frmCuentaCorrienteGrupo.Show
                If frmCuentaCorrienteGrupo.WindowState = vbMinimized Then
                    frmCuentaCorrienteGrupo.WindowState = vbNormal
                End If
                frmCuentaCorrienteGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_CUENTACORRIENTE_CAJA"
            If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA) Then
                Screen.MousePointer = vbHourglass
                frmCuentaCorrienteCaja.Show
                If frmCuentaCorrienteCaja.WindowState = vbMinimized Then
                    frmCuentaCorrienteCaja.WindowState = vbNormal
                End If
                frmCuentaCorrienteCaja.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_MEDIOPAGO"
            If pCPermiso.GotPermission(PERMISO_MEDIOPAGO) Then
                Screen.MousePointer = vbHourglass
                frmMedioPago.Show
                If frmMedioPago.WindowState = vbMinimized Then
                    frmMedioPago.WindowState = vbNormal
                End If
                frmMedioPago.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_FERIADO"
            If pCPermiso.GotPermission(PERMISO_FERIADO) Then
                Screen.MousePointer = vbHourglass
                frmFeriado.Show
                If frmFeriado.WindowState = vbMinimized Then
                    frmFeriado.WindowState = vbNormal
                End If
                frmFeriado.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_VEHICULO_MANTENIMIENTO_GRUPO"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_GRUPO) Then
                Screen.MousePointer = vbHourglass
                frmVehiculoMantenimientoGrupo.Show
                If frmVehiculoMantenimientoGrupo.WindowState = vbMinimized Then
                    frmVehiculoMantenimientoGrupo.WindowState = vbNormal
                End If
                frmVehiculoMantenimientoGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_VEHICULO_MANTENIMIENTO"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO) Then
                Screen.MousePointer = vbHourglass
                frmVehiculoMantenimiento.Show
                If frmVehiculoMantenimiento.WindowState = vbMinimized Then
                    frmVehiculoMantenimiento.WindowState = vbNormal
                End If
                frmVehiculoMantenimiento.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_DOCUMENTO_TIPO"
            If pCPermiso.GotPermission(PERMISO_DOCUMENTO_TIPO) Then
                Screen.MousePointer = vbHourglass
                frmDocumentoTipo.Show
                If frmDocumentoTipo.WindowState = vbMinimized Then
                    frmDocumentoTipo.WindowState = vbNormal
                End If
                frmDocumentoTipo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_TELEFONO_TIPO"
            If pCPermiso.GotPermission(PERMISO_TELEFONO_TIPO) Then
                Screen.MousePointer = vbHourglass
                frmTelefonoTipo.Show
                If frmTelefonoTipo.WindowState = vbMinimized Then
                    frmTelefonoTipo.WindowState = vbNormal
                End If
                frmTelefonoTipo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_PERSONA_ALARMA"
            If pCPermiso.GotPermission(PERMISO_PERSONA_ALARMA) Then
                Screen.MousePointer = vbHourglass
                frmPersonaAlarma.Show
                If frmPersonaAlarma.WindowState = vbMinimized Then
                    frmPersonaAlarma.WindowState = vbNormal
                End If
                frmPersonaAlarma.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_PERSONA_ALARMA_GRUPO"
            If pCPermiso.GotPermission(PERMISO_PERSONA_ALARMA_GRUPO) Then
                Screen.MousePointer = vbHourglass
                frmPersonaAlarmaGrupo.Show
                If frmPersonaAlarmaGrupo.WindowState = vbMinimized Then
                    frmPersonaAlarmaGrupo.WindowState = vbNormal
                End If
                frmPersonaAlarmaGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_ALARMA"
            If pCPermiso.GotPermission(PERMISO_ALARMA) Then
                Screen.MousePointer = vbHourglass
                frmAlarma.Show
                If frmAlarma.WindowState = vbMinimized Then
                    frmAlarma.WindowState = vbNormal
                End If
                frmAlarma.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "DATOS_INICIALES_CONTACTO_GRUPO"
            If pCPermiso.GotPermission(PERMISO_CONTACTO_GRUPO) Then
                Screen.MousePointer = vbHourglass
                frmContactoGrupo.Show
                If frmContactoGrupo.WindowState = vbMinimized Then
                    frmContactoGrupo.WindowState = vbNormal
                End If
                frmContactoGrupo.SetFocus
                Screen.MousePointer = vbDefault
            End If
            
        '//////////////////////////////////////////////////////////////////
        'O T R O S
        '//////////////////////////////////////////////////////////////////
        Case "OTROS_VEHICULO_MANTENIMIENTO_ACCION"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_ACCION) Then
                Screen.MousePointer = vbHourglass
                frmVehiculoMantenimientoAccion.Show
                If frmVehiculoMantenimientoAccion.WindowState = vbMinimized Then
                    frmVehiculoMantenimientoAccion.WindowState = vbNormal
                End If
                frmVehiculoMantenimientoAccion.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_VEHICULO_MANTENIMIENTO_ACCION_HISTORICO"
'            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_ACCION_HISTORICO) Then
'                Screen.MousePointer = vbHourglass
'                frmVehiculoMantenimientoAccion.Show
'                If frmVehiculoMantenimientoAccion.WindowState = vbMinimized Then
'                    frmVehiculoMantenimientoAccion.WindowState = vbNormal
'                End If
'                frmVehiculoMantenimientoAccion.SetFocus
'                Screen.MousePointer = vbDefault
'            End If
        Case "OTROS_CONTACTO"
            If pCPermiso.GotPermission(PERMISO_CONTACTO) Then
                Screen.MousePointer = vbHourglass
                frmContacto.Show
                If frmContacto.WindowState = vbMinimized Then
                    frmContacto.WindowState = vbNormal
                End If
                frmContacto.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_VIAJE_TRANSFERENCIA"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_TRANSFER) Then
                Screen.MousePointer = vbHourglass
                frmViajeDetalleTransferencia.Show
                If frmViajeDetalleTransferencia.WindowState = vbMinimized Then
                    frmViajeDetalleTransferencia.WindowState = vbNormal
                End If
                frmViajeDetalleTransferencia.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_VIAJE_ASISTENCIA"
            If pCPermiso.GotPermission(PERMISO_VIAJE_ASISTENCIA) Then
                Screen.MousePointer = vbHourglass
                frmViajeAsistencia.Show
                If frmViajeAsistencia.WindowState = vbMinimized Then
                    frmViajeAsistencia.WindowState = vbNormal
                End If
                frmViajeAsistencia.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_VIAJE_CONDUCTOR"
            If pCPermiso.GotPermission(PERMISO_VIAJE_CONDUCTOR) Then
                Screen.MousePointer = vbHourglass
                frmViajeConductor.Show
                If frmViajeConductor.WindowState = vbMinimized Then
                    frmViajeConductor.WindowState = vbNormal
                End If
                frmViajeConductor.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_VEHICULO_UTILIZACION"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_UTILIZACION) Then
                Screen.MousePointer = vbHourglass
                frmVehiculoUtilizacion.LoadDataAndShow
                If frmVehiculoUtilizacion.WindowState = vbMinimized Then
                    frmVehiculoUtilizacion.WindowState = vbNormal
                End If
                frmVehiculoUtilizacion.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_LISTAPASAJERO"
            If pCPermiso.GotPermission(PERMISO_LISTAPASAJERO) Then
                frmListaPasajero.Show
            End If
        Case "OTROS_REGISTROLLAMADA"
            If pCPermiso.GotPermission(PERMISO_REGISTROLLAMADA) Then
                Screen.MousePointer = vbHourglass
                frmRegistroLlamada.Show
                If frmRegistroLlamada.WindowState = vbMinimized Then
                    frmRegistroLlamada.WindowState = vbNormal
                End If
                frmRegistroLlamada.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_FRANCOS"
            If pCPermiso.GotPermission(PERMISO_FRANCO) Then
                Screen.MousePointer = vbHourglass
                frmFranco.Show
                If frmFranco.WindowState = vbMinimized Then
                    frmFranco.WindowState = vbNormal
                End If
                frmFranco.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_MENSAJES"
            If pCPermiso.GotPermission(PERMISO_MENSAJE) Then
                Screen.MousePointer = vbHourglass
                frmMensajeLista.Show
                If frmMensajeLista.WindowState = vbMinimized Then
                    frmMensajeLista.WindowState = vbNormal
                End If
                frmMensajeLista.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "OTROS_APPEXT_MSWORD"    'MICROSOFT WORD
            If MSWordApp Is Nothing Then
                On Error Resume Next
                Err.Clear
                Set MSWordApp = CreateObject("Word.Application")
            End If
            If Err.Number <> 0 Then
                MsgBox "No se pudo iniciar Microsoft Word.", vbExclamation, App.Title
            Else
                MSWordApp.Visible = True
            End If
        
        Case "OTROS_APPEXT_MSEXCEL"   'MICROSOFT EXCEL
            If MSExcelApp Is Nothing Then
                On Error Resume Next
                Err.Clear
                Set MSExcelApp = CreateObject("Excel.Application")
            End If
            If Err.Number <> 0 Then
                MsgBox "No se pudo iniciar Microsoft Excel.", vbExclamation, App.Title
            Else
                MSExcelApp.Visible = True
            End If
    
        Case "VIAJE_ACTUAL"
            If pCPermiso.GotPermission(PERMISO_VIAJE) Then
                Screen.MousePointer = vbHourglass
                frmViaje.Show
                If frmViaje.WindowState = vbMinimized Then
                    frmViaje.WindowState = vbNormal
                End If
                frmViaje.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "VIAJE_HISTORICO"
'            If pCPermiso.GotPermission(PERMISO_VIAJE_HISTORICO) Then
'                Screen.MousePointer = vbHourglass
'                frmViaje.Show
'                If frmViaje.WindowState = vbMinimized Then
'                    frmViaje.WindowState = vbNormal
'                End If
'                frmViaje.SetFocus
'                Screen.MousePointer = vbDefault
'            End If
    
        Case "COMISION_ACTUAL"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE) Then
                Screen.MousePointer = vbHourglass
                frmComision.Show
                If frmComision.WindowState = vbMinimized Then
                    frmComision.WindowState = vbNormal
                End If
                frmComision.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "COMISION_HISTORICO"
'            If pCPermiso.GotPermission(PERMISO_VIAJE_HISTORICO) Then
'                Screen.MousePointer = vbHourglass
'                frmComision.Show
'                If frmComision.WindowState = vbMinimized Then
'                    frmComision.WindowState = vbNormal
'                End If
'                frmComision.SetFocus
'                Screen.MousePointer = vbDefault
'            End If
    
        Case "CUENTA_CORRIENTE_ACTUAL"
            If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE) Then
                Screen.MousePointer = vbHourglass
                frmCuentaCorriente.DatabaseName = pParametro.Database_Database
                frmCuentaCorriente.IsHistory = False
                frmCuentaCorriente.Caption = "Cuenta Corriente Actual"
                frmCuentaCorriente.Show
                If frmCuentaCorriente.WindowState = vbMinimized Then
                    frmCuentaCorriente.WindowState = vbNormal
                End If
                frmCuentaCorriente.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "CUENTA_CORRIENTE_HISTORICO"
            If pParametro.Database_DatabaseHistory = "" Then
                MsgBox "No se ha especificado la Base de Datos Histórica.", vbInformation, App.Title
                Exit Sub
            End If
            If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_HISTORICO) Then
                Screen.MousePointer = vbHourglass
                frmCuentaCorriente.DatabaseName = pParametro.Database_DatabaseHistory
                frmCuentaCorriente.IsHistory = True
                frmCuentaCorriente.Caption = "Cuenta Corriente Histórica"
                frmCuentaCorriente.Show
                If frmCuentaCorriente.WindowState = vbMinimized Then
                    frmCuentaCorriente.WindowState = vbNormal
                End If
                frmCuentaCorriente.SetFocus
                Screen.MousePointer = vbDefault
            End If
    End Select
End Sub

Private Sub cboPersona_Click()
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    If cboPersona.ListIndex > -1 Then
        If mCTelefonoTipoNombre.Count >= cboPersona.ListIndex + 1 Then
            txtCallerIDTipo.Text = mCTelefonoTipoNombre(cboPersona.ListIndex + 1)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    txtCallerIDTipo.Text = ""
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        If cboPersona.ListIndex > -1 Then
            Call frmPersona.FindAndShowItem(cboPersona.ItemData(cboPersona.ListIndex), UCase(Left(cboPersona.Text, 1)), Me.Name, "", "")
        End If
    End If
End Sub

Private Sub stbMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If pPersonal Then
        'VERIFICO SI ES EL BOTON DERECHO
        If Button = vbLeftButton Then
            'VERIFICO SI ESTAN PRESIONADAS CONTROL + ALT + SHIFT
            If Shift = (vbShiftMask Or vbCtrlMask Or vbAltMask) Then
                'VERIFICO SI EL CLICK ESTA DENTRO DE LAS COORDENADAS X DEL PANEL
                If x >= stbMain.Panels("PERSONAL").Left And x <= (stbMain.Panels("PERSONAL").Left + stbMain.Panels("PERSONAL").Width) Then
                    LogAccionAdd "", "Personal: Desactivado."
                    pPersonal = False
                    RefreshList_RefreshPersonal
                End If
            End If
        End If
    End If
End Sub

Private Sub stbMain_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Key
        Case "CENTRO_MENSAJE"
            'MENSAJES GENERALES
            If frmMensaje.CheckMessages() Then
                frmMensaje.Show vbModal, frmMDI
            Else
                Unload frmMensaje
                MsgBox "No hay mensajes pendientes.", vbInformation, App.Title
            End If
    End Select
End Sub

Private Sub stbMain_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Key
        Case "USUARIO"
            mnuFileCloseSession_Click
        Case "PERSONAL"
            If Not pPersonal Then
                LogAccionAdd "", "Personal: Activado."
                pPersonal = True
                RefreshList_RefreshPersonal
            End If
    End Select
End Sub

Private Sub tmrMulti_Timer()
    Static RefreshList_Fastest_Counter As Long
    Static RefreshList_Slowest_Counter As Long
    Static RefreshList_Personal_Counter As Long
    Static Viaje_EstadoVencido_Counter As Long
    Static SyncronizeServerDateTime_Counter As Long
    Static MessengerBlinkingOn As Boolean
    
    If pParametro.RefreshList_Enabled Then
        'FASTEST
        RefreshList_Fastest_Counter = RefreshList_Fastest_Counter + 1
        If RefreshList_Fastest_Counter >= pParametro.RefreshList_Fastest_CheckInterval_Seconds Then
            RefreshList_Fastest_Counter = 0
            Call RefreshList_Fastest_CheckForRefreshs
        End If
        'SLOWEST
        RefreshList_Slowest_Counter = RefreshList_Slowest_Counter + 1
        If RefreshList_Slowest_Counter >= pParametro.RefreshList_Slowest_CheckInterval_Seconds Then
            RefreshList_Slowest_Counter = 0
            Call RefreshList_Slowest_CheckForRefreshs
        End If
    Else
        If pParametro.Personal_Status_Global Then
            RefreshList_Personal_Counter = RefreshList_Personal_Counter + 1
            If RefreshList_Personal_Counter >= pParametro.Personal_Status_Global_CheckInterval_Seconds Then
                RefreshList_Personal_Counter = 0
                Call RefreshList_CheckForPersonal
            End If
        End If
    End If
    
    If pTelephony.TelephonyType = "COMM" Then
        If pTelephony.CallerIDSupported Then
            Call pTelephony.CallerID_Check
        End If
    End If
    
    If pTelephony.TelephonyType = "SERVER" Then
        pTelephony.SERVER_CheckForCall
    End If
    
    If pParametro.Viaje_EstadoVencido_Check Then
        Viaje_EstadoVencido_Counter = Viaje_EstadoVencido_Counter + 1
        If Viaje_EstadoVencido_Counter >= pParametro.Viaje_EstadoVencido_CheckIntervalSeconds Then
            Viaje_EstadoVencido_Counter = 0
            Call Viaje_EstadoVencido_List
        End If
    End If
    
    '//////////////////////////////////////////////////////////////////
    'SINCRONIZO LA HORA CON EL SERVER
'    If pParametro.DateTime_SyncWithServer_Primary And CSM_Session.GetComputerName() = "" Then
'    ElseIf pParametro.DateTime_SyncWithServer_Primary And CSM_Session.GetComputerName() Then
'    End If
    If pParametro.DateTime_SyncWithServer_Primary And pParametro.DateTime_SyncWithServer_Secondary Then
        SyncronizeServerDateTime_Counter = SyncronizeServerDateTime_Counter + 1
        If SyncronizeServerDateTime_Counter >= pParametro.DateTime_SyncWithServer_IntervalSeconds Then
            SyncronizeServerDateTime_Counter = 0
            Call CSM_Time.SyncronizeWithSQLServer
        End If
    End If

    If pMessengerEnabled And pMessengerBlinking Then
        If MessengerBlinkingOn Then
            tlbMain.Buttons("MESSENGER").Image = "MESSENGER"
        Else
            tlbMain.Buttons("MESSENGER").Image = 0
        End If
        MessengerBlinkingOn = Not MessengerBlinkingOn
    End If
End Sub

Private Sub mTAPI_Event(ByVal TapiEvent As TAPI3Lib.TAPI_EVENT, ByVal pEvent As Object)
    Dim CallNotificationEvent As ITCallNotificationEvent
    Dim CallInfoChangeEvent As ITCallInfoChangeEvent
    
    Select Case TapiEvent
    
        '///////////////////////////////////////////////////////////////
        Case TE_ADDRESS
            '### An Address object has changed.
'            Set AddressEvent = pEvent
'            Select Case AddressEvent.Event
'                Case AE_STATE
'                    lstEvent.AddItem "TE_ADDRESS -> AE_STATE: Incoming Call Notification."
'                Case AE_CAPSCHANGE
'                    lstEvent.AddItem "TE_ADDRESS -> AE_CAPSCHANGE: Address capabilities have changed."
'                Case AE_RINGING
'                    lstEvent.AddItem "TE_ADDRESS -> AE_RINGING: There is ringing on the address."
'                Case AE_CONFIGCHANGE
'                    lstEvent.AddItem "TE_ADDRESS -> AE_CONFIGCHANGE: Address configuration has changed."
'                Case AE_FORWARD
'                    lstEvent.AddItem "TE_ADDRESS -> AE_FORWARD: Forwarding has changed."
'                Case AE_NEWTERMINAL
'                    lstEvent.AddItem "TE_ADDRESS -> AE_NEWTERMINAL: New terminal."
'                Case AE_REMOVETERMINAL
'                    lstEvent.AddItem "TE_ADDRESS -> AE_REMOVETERMINAL: Terminal removed."
'            End Select
'            Set AddressEvent = Nothing
            
        '///////////////////////////////////////////////////////////////
        Case TE_CALLNOTIFICATION
            '### A new communications session has appeared on the address and the
            '### TAPI DLL has created a new call object. This could be a result
            '### from an incoming session, a session being handed off by another
            '### application, or a session being parked on the address.
            Set CallNotificationEvent = pEvent
'            Select Case CallNotificationEvent.Event
'                Case CNE_OWNER
'                    lstEvent.AddItem "TE_CALLNOTIFICATION -> CNE_OWNER: The current application owns the call on which the event occurred."
'                Case CNE_MONITOR
'                    lstEvent.AddItem "TE_CALLNOTIFICATION -> CNE_MONITOR: The current application is monitoring the call on which the event occurred."
'            End Select
            

            'STORES CALL INFO IN MODULE VARIABLE
            Set mCallInfo = CallNotificationEvent.Call
            Set CallNotificationEvent = Nothing
            
        '///////////////////////////////////////////////////////////////
        Case TE_CALLSTATE
            '### The Call state has changed.
'            Set CallStateEvent = pEvent
'            Select Case CallStateEvent.State
'                Case CS_IDLE
'                    lstEvent.AddItem "TE_CALLSTATE -> CS_IDLE: The call has been created, but Connect has not been called yet. A call can never transition into the idle state."
'                Case CS_INPROGRESS
'                    lstEvent.AddItem "TE_CALLSTATE -> CS_INPROGRESS: Connect has been called, and the service provider is working on making a connection. This state is valid only on outgoing calls. This message is optional, because a service provider may have a call transition directly to the connected state."
'                Case CS_CONNECTED
'                    lstEvent.AddItem "TE_CALLSTATE -> CS_CONNECTED: Call has been connected to the remote end and communication can take place."
'                Case CS_DISCONNECTED
'                    Select Case CallStateEvent.Cause
'                        Case CEC_NONE
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_NONE: No call event has occurred."
'                        Case CEC_DISCONNECT_NORMAL
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_NORMAL: The call was disconnected as part of the normal life cycle of the call."
'                        Case CEC_DISCONNECT_BUSY
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_BUSY: An outgoing call failed to connect because the remote end was busy."
'                        Case CEC_DISCONNECT_BADADDRESS
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_BADADDRESS: An outgoing call failed because the destination address was bad."
'                        Case CEC_DISCONNECT_NOANSWER
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_NOANSWER: An outgoing call failed because the remote end was not answered."
'                        Case CEC_DISCONNECT_CANCELLED
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_CANCELLED: An outgoing call failed because the caller disconnected."
'                        Case CEC_DISCONNECT_REJECTED
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_REJECTED: The outgoing call was rejected by the remote end."
'                        Case CEC_DISCONNECT_FAILED
'                            lstEvent.AddItem "TE_CALLSTATE -> CS_DISCONNECTED -> CEC_DISCONNECT_FAILED: The call failed to connect for some other reason."
'                    End Select
'                Case CS_OFFERING
'                    lstEvent.AddItem "TE_CALLSTATE -> CS_OFFERING: A new call has appeared, and is being offered to an application."
'                Case CS_HOLD
'                    lstEvent.AddItem "TE_CALLSTATE -> CS_HOLD: The call is in the hold state."
'                Case CS_QUEUED
'                    lstEvent.AddItem "TE_CALLSTATE -> CS_QUEUED: The call is queued."
'            End Select
'            Set CallStateEvent = Nothing
            
        '///////////////////////////////////////////////////////////////
        Case TE_CALLMEDIA
            '### The media associated with a call has changed.
'            Set CallMediaEvent = pEvent
'            Select Case CallMediaEvent.Event
'                Case CME_NEW_STREAM
'                    lstEvent.AddItem "TE_CALLMEDIA -> CME_NEW_STREAM: A new media stream has been created."
'                Case CME_STREAM_FAIL
'                    Select Case CallMediaEvent.Cause
'                        Case CMC_UNKNOWN
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_UNKNOWN: Call media is unknown."
'                        Case CMC_BAD_DEVICE
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_BAD_DEVICE: Device source or renderer is not functioning."
'                        Case CMC_CONNECT_FAIL
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_CONNECT_FAIL: Could not connect to media device."
'                        Case CMC_LOCAL_REQUEST
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_LOCAL_REQUEST: A local request has been received."
'                        Case CMC_REMOTE_REQUEST
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_REMOTE_REQUEST: A remote request has been received."
'                        Case CMC_MEDIA_TIMEOUT
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_MEDIA_TIMEOUT: The media device timed out."
'                        Case CMC_MEDIA_RECOVERED
'                            lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_FAIL -> CMC_MEDIA_RECOVERED: Media processing has resumed after an interruption."
'                    End Select
'                Case CME_TERMINAL_FAIL
'                    lstEvent.AddItem "TE_CALLMEDIA -> CME_TERMINAL_FAIL: A terminal has failed."
'                Case CME_STREAM_NOT_USED
'                    lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_NOT_USED: The media stream has not been used."
'                Case CME_STREAM_ACTIVE
'                    lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_ACTIVE: The media stream is active."
'
'                    'START RECORDING
'                    mDirectSound_Capture.StartCapture
'
'                Case CME_STREAM_INACTIVE
'                    lstEvent.AddItem "TE_CALLMEDIA -> CME_STREAM_INACTIVE: The media stream is not active."
'
'                    'END RECORDING
'                    mDirectSound_Capture.StopCapture
'            End Select
'            Set CallMediaEvent = Nothing
            
        '///////////////////////////////////////////////////////////////
        Case TE_CALLINFOCHANGE
            '### The call information has changed.
            Set CallInfoChangeEvent = pEvent
            txtCallerID.Text = mCallInfo.CallInfoString(CIS_CALLERIDNUMBER)
            Set CallInfoChangeEvent = Nothing
            CallerID_BuscarPersonas
            CallerID_BuscarVehiculos
    End Select
End Sub

Public Sub CallerID_BuscarPersonas()
    Dim cmdPersona As ADODB.command
    Dim recPersona As ADODB.Recordset
    
    Dim recTelefonoTipo As ADODB.Recordset
    
    Dim errorMessage As String
    Dim TelefonoIndex As Long
    Dim IDTelefonoTipo As Long
    
    Dim NumeroCompleto As String

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    NumeroCompleto = IIf(Left(txtCallerID.Text, 1) <> "0", pTelephony.LocationCityCode & txtCallerID.Text, txtCallerID.Text)
    
    errorMessage = "Error al obtener la Lista de Personas por el Número de Teléfono."
    
    Set cmdPersona = New ADODB.command
    Set cmdPersona.ActiveConnection = pDatabase.Connection
    cmdPersona.CommandText = "sp_Persona_CallerID_Search"
    cmdPersona.CommandType = adCmdStoredProc
    cmdPersona.Parameters.Append cmdPersona.CreateParameter("@TelefonoAreaLocal", adVarChar, adParamInput, 5, pTelephony.LocationCityCode)
    cmdPersona.Parameters.Append cmdPersona.CreateParameter("@TelefonoNumero", adVarChar, adParamInput, 21, NumeroCompleto)
    Set recPersona = New ADODB.Recordset
    recPersona.Open cmdPersona, , adOpenForwardOnly, adLockReadOnly
    Set cmdPersona = Nothing
    
    errorMessage = "Error al leer el Tipo de Teléfono."
    Set recTelefonoTipo = New ADODB.Recordset
    recTelefonoTipo.Source = "SELECT * FROM TelefonoTipo"
    recTelefonoTipo.Open , pDatabase.Connection, adOpenStatic, adLockReadOnly, adCmdText
    
    errorMessage = "Error al leer las Personas según el Número de Teléfono."
    cboPersona.Clear
    txtCallerIDTipo.Text = ""
    Set mCTelefonoTipoNombre = New Collection
    Do While Not recPersona.EOF
        cboPersona.AddItem recPersona("Persona").Value
        cboPersona.ItemData(cboPersona.NewIndex) = recPersona("IDPersona").Value
        
        'Busco el Tipo de Teléfono
        Select Case NumeroCompleto
            Case IIf(IsNull(recPersona("Telefono1Area").Value), pTelephony.LocationCityCode, recPersona("Telefono1Area").Value) & recPersona("Telefono1Numero").Value
                TelefonoIndex = 1
            Case IIf(IsNull(recPersona("Telefono2Area").Value), pTelephony.LocationCityCode, recPersona("Telefono2Area").Value) & recPersona("Telefono2Numero").Value
                TelefonoIndex = 2
            Case IIf(IsNull(recPersona("Telefono3Area").Value), pTelephony.LocationCityCode, recPersona("Telefono3Area").Value) & recPersona("Telefono3Numero").Value
                TelefonoIndex = 3
            Case IIf(IsNull(recPersona("Telefono4Area").Value), pTelephony.LocationCityCode, recPersona("Telefono4Area").Value) & recPersona("Telefono4Numero").Value
                TelefonoIndex = 4
            Case IIf(IsNull(recPersona("Telefono5Area").Value), pTelephony.LocationCityCode, recPersona("Telefono5Area").Value) & recPersona("Telefono5Numero").Value
                TelefonoIndex = 5
        End Select
        errorMessage = "Ha ocurrido un error al leer el Tipo de Teléfono."
        
        IDTelefonoTipo = Val(recPersona("IDTelefono" & TelefonoIndex & "Tipo").Value & "")
        Select Case IDTelefonoTipo
            Case 0
                mCTelefonoTipoNombre.Add ""
            Case pParametro.TelefonoTipo_ID_Otro
                mCTelefonoTipoNombre.Add recPersona("Telefono" & TelefonoIndex & "TipoOtro").Value & ""
            Case Else
                recTelefonoTipo.Filter = "IDTelefonoTipo = " & IDTelefonoTipo
                If recTelefonoTipo.EOF Then
                    Screen.MousePointer = vbDefault
                    MsgBox "No se ha encontrado el Tipo de Teléfono.", vbExclamation, App.Title
                    Screen.MousePointer = vbHourglass
                    mCTelefonoTipoNombre.Add ""
                Else
                    mCTelefonoTipoNombre.Add recTelefonoTipo("Nombre").Value & ""
                End If
        End Select
        
        errorMessage = "Error al leer las Personas según el Número de Teléfono."
        recPersona.MoveNext
    Loop
    recTelefonoTipo.Close
    Set recTelefonoTipo = Nothing
    recPersona.Close
    Set recPersona = Nothing
    If cboPersona.ListCount > 0 Then
        cboPersona.ListIndex = 0
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.MDI.BuscarPersonasPorCallerID", errorMessage
End Sub

Public Sub CallerID_BuscarVehiculos()
    Dim recVehiculo As ADODB.Recordset
    Dim NumeroCompleto As String

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    NumeroCompleto = IIf(Left(txtCallerID.Text, 1) <> "0", pTelephony.LocationCityCode & txtCallerID.Text, txtCallerID.Text)
    
    Set recVehiculo = New ADODB.Recordset
    recVehiculo.Source = "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE ISNULL(TelefonoArea, '" & pTelephony.LocationCityCode & "') + TelefonoNumero = '" & NumeroCompleto & "' ORDER BY Nombre"
    recVehiculo.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not recVehiculo.EOF Then
        If pCPermiso.GotPermission(PERMISO_VIAJE, False) Then
            If MsgBox("Está ingresando una Llamada desde uno de los Vehículos." & vbCr & vbCr & "Vehículo: " & recVehiculo("Nombre").Value & vbCr & vbCr & "¿Desea abrir el Viaje que está realizando este Vehículo?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                'ABRO EL VIAJE
                BuscarYAbrirViajeActualPorVehiculo recVehiculo("IDVehiculo").Value, recVehiculo("Nombre").Value
            End If
        End If
    End If
    
    recVehiculo.Close
    Set recVehiculo = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.MDI.BuscarVehiculosPorCallerID", "Error al leer las Personas según el Número de Teléfono."
End Sub
