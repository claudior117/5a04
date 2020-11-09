VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_bancos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO BANCOS"
   ClientHeight    =   8580
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12000
   FontTransparent =   0   'False
   Icon            =   "inicio_bancos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8580
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   735
      Left            =   4920
      TabIndex        =   23
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   4080
         Picture         =   "inicio_bancos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CAJA DIARIA"
      Height          =   1095
      Left            =   4680
      TabIndex        =   21
      Top             =   3120
      Width           =   2655
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   645
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1138
         ButtonWidth     =   2037
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Caja Diaria"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Composicion de los saldos de caja"
               ImageIndex      =   3
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MOVIMIENTOS BANCARIOS"
      Height          =   1095
      Left            =   4680
      TabIndex        =   19
      Top             =   1800
      Width           =   5175
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   645
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1138
         ButtonWidth     =   2037
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Administrador de Cheques de Terceros"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Generar chequeras propias"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Administrador Cheques Propios"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Object.ToolTipText     =   "Estado de Cuenta Bancaria"
               ImageIndex      =   2
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MOVIMIENTOS BANCARIOS"
      Height          =   1095
      Left            =   4680
      TabIndex        =   17
      Top             =   480
      Width           =   5175
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   645
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1138
         ButtonWidth     =   2037
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Depositos Bancarios"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Debitos y Creditos Bancarios"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Venta de Cheques"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Description     =   "Retiro de Efectivo por caja o Cajero"
               Object.ToolTipText     =   "Retiro de Efectivo por caja o Cajero"
               ImageIndex      =   4
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ARCHIVOS MAESTROS"
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   3975
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1535
         ButtonWidth     =   2963
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "    BANCOS       "
               Key             =   "B2"
               Object.ToolTipText     =   "Bancos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Conceptos Db. y Cr."
               Key             =   "B3"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6480
      TabIndex        =   13
      Top             =   5640
      Width           =   2055
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BANCOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   720
         Picture         =   "inicio_bancos.frx":0614
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   4575
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "CUIT:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Telefono:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Direccion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Razon Social:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9960
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "inicio_bancos.frx":0E96
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "inicio_bancos.frx":1718
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8325
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "29/12/2017"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "01:24 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":1F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":22B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   720
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":25CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":2CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":3481
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":3C12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":439D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":4A4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":51B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":5888
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   3480
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":5F89
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":66D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":6E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_bancos.frx":75F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Caption         =   "VERSION DE PRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_bancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu M_concp 
         Caption         =   "Conceptos para Db. y Cr."
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Consultas"
      Begin VB.Menu M_cartera 
         Caption         =   "Cartera Cheques de Tercero"
      End
      Begin VB.Menu M_carter2 
         Caption         =   "Cartera Cheques a una Fecha"
      End
      Begin VB.Menu M_infochterc 
         Caption         =   "Informe Cheques de Tercero"
      End
      Begin VB.Menu M_dbcr 
         Caption         =   "Informe Db. y Cr. Bancarios"
      End
      Begin VB.Menu M_movban 
         Caption         =   "Informe de Movmientos emitidos "
      End
      Begin VB.Menu M_chvend 
         Caption         =   "Saldo Cheques vendidos"
      End
      Begin VB.Menu M_chvebc 
         Caption         =   "Agenda bancaria -Cheques por vencer-"
      End
   End
   Begin VB.Menu m_utiles 
      Caption         =   "&Utiles"
      Begin VB.Menu M_conciava 
         Caption         =   "Conciliacion Avanzada"
      End
      Begin VB.Menu M_calactu 
         Caption         =   "Calculo Actualizacion de Ch. Dif. segun Cotizacion"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnsale_Click()
inicio.Show
Unload Me
End Sub




Private Sub Command1_Click()
Call nivel_acceso(4)
If para.id_grupo_modulo_actual >= 5 Then
  cyb_depositoS.Show
 Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
Call nivel_acceso(4)
If para.id_grupo_modulo_actual >= 2 Then
 cyb_carterach.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command3_Click()
Call nivel_acceso(4)
If para.id_grupo_modulo_actual >= 5 Then
 cyb_generachpropios.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command4_Click()
cyb_chpropios.Show
End Sub

Private Sub Command5_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Command6_Click()
cyb_movbanco.Show
End Sub

Private Sub Command7_Click()
cyb_VENTACH.Show
End Sub

Private Sub Form_Activate()
Call barraesag(Me)
Label7 = para.impresora_actual
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_Load()
Call titulos(Me)

Exit Sub

e1:
  MsgBox ("Error al Inicializar Parametros INICIO.LOAD")
  End

End Sub



Private Sub M_bancos_Click()
Call nivel_acceso(4)
If para.id_grupo_modulo_actual > 3 Then
 cyb_ABM_bancos.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_cuentas_Click()

End Sub

Private Sub M_fp_Click()
cyb_ABM_FP.Show
End Sub

Private Sub M_calactu_Click()
cyb_actucot_ch_terc.Show
End Sub

Private Sub M_carter2_Click()
cyb_carterach2.Show
End Sub

Private Sub M_cartera_Click()
  Call nivel_acceso(4)
  If para.id_grupo_modulo_actual >= 4 Then
      cyb_carterach.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub M_chvebc_Click()
cyb_venc_ch.Show
End Sub

Private Sub M_chvend_Click()
cyb_cuenta_ch_vend.Show
End Sub

Private Sub M_conciava_Click()
cyb_concilia2.Show
End Sub

Private Sub M_concp_Click()
Call nivel_acceso(4)
If para.id_grupo_modulo_actual > 3 Then
   cyb_abmconceptos.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_dbcr_Click()
cyb_informedbcr.Show
End Sub

Private Sub M_infochterc_Click()
Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 4 Then
      cyb_carterach.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub M_movban_Click()
cyb_informeMOV.Show
End Sub

Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B2"
   Call nivel_acceso(4)
   If para.id_grupo_modulo_actual > 3 Then
     cyb_ABM_bancos.Show
   Else
     Call sinpermisos
   End If
  Case Is = "B3"
   Call nivel_acceso(4)
   If para.id_grupo_modulo_actual > 3 Then
     cyb_abmconceptos.Show
   Else
     Call sinpermisos
   End If

End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_depositoS.Show
    Else
      Call sinpermisos
    End If

 Case Is = "B2"
    Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_movbanco.Show
    Else
      Call sinpermisos
    End If
    
    
 Case Is = "B3"
   Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_VENTACH.Show
    Else
      Call sinpermisos
    End If
 
Case Is = "B4"
   Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_retiroef.Show
    Else
      Call sinpermisos
    End If
 

End Select

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 4 Then
      cyb_carterach.Show
    Else
      Call sinpermisos
    End If

 Case Is = "B2"
    Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 4 Then
      cyb_generachpropios.Show
    Else
      Call sinpermisos
    End If
    
    
 Case Is = "B3"
   Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 4 Then
      cyb_chpropios.Show
    Else
      Call sinpermisos
    End If
 
 Case Is = "B4"
   Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 4 Then
      cyb_estadocuenta.Show
    Else
      Call sinpermisos
    End If
End Select


End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 4 Then
     cja_cajadiaria.Show
    Else
     Call sinpermisos
    End If
 Case Is = "B2"
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_cajadiaria.Show
    Else
      Call sinpermisos
    End If
End Select
End Sub
