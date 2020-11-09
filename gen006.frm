VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_consultaib 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CONSULTA PADRON INGRESOS BRUTOS"
   ClientHeight    =   5160
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5160
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consulta de Embargo"
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   5895
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estado de la Consulta"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   5895
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos en el Padron IB"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   5895
      Begin VB.TextBox T_RETIB 
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox T_PERCIB 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Tasa Ret. IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Tasa Perc. IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Tabla Clientes"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      Begin VB.TextBox t_tipo 
         Height          =   405
         Left            =   4320
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox t_percive 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox t_id 
         Height          =   405
         Left            =   4320
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox t_provincia 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox t_cuit 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox t_localidad 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox t_cli 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Caption         =   "Percive IB:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Provincia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Cuit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Localidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Razon Social"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4905
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "gen_consultaib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean





Sub limpia()
 t_cli = " "
 t_direccion = " "
 t_localidad = " "
 t_cuit = " "
 t_provincia = " "
 T_PERCIB = " "
 T_RETIB = " "
 Label5 = ""

 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If

End Sub

Sub carga()
If t_tipo = "C" Then 'clientes
 Set cl_cli = New Clientes
 'cl_cli.id = Val(t_id)
 cl_cli.carga (Val(t_id))
 If cl_cli.id > 0 Then
  t_cuit = cl_cli.CUIT
  t_cli = cl_cli.razonsocial
  t_localidad = cl_cli.localidad
  t_provincia = cl_cli.provincia
  t_percive = cl_cli.perciveib
 Else
   t_cuit = 0
 End If
 Label8 = "Percive IB"
 Set cl_cli = Nothing
Else
 Set cl_prov = New proveedores
  'cl_prov.idprov = Val(t_id)
  cl_prov.carga (Val(t_id))
  If cl_prov.idprov > 0 Then
   t_cuit = cl_prov.CUIT
   t_cli = cl_prov.razonsocial
   t_localidad = cl_prov.localidad
   t_provincia = cl_prov.provincia
   t_percive = cl_prov.retieneib
  Else
   t_cuit = 0
  End If
  Label8 = "Retiene IB"
  Set cl_prov = Nothing

End If
Set cl_padronib = New padron_ib
cl_padronib.cuit_texto = t_cuit
cl_padronib.buscar
T_PERCIB = Format$(cl_padronib.tasa_percib, "###0.00")
T_RETIB = Format$(cl_padronib.tasa_retib, "###0.00")
Select Case cl_padronib.estado_consulta
Case Is = "OK"
   t_cuit = cl_padronib.CUIT
   Label5 = "¡Consulta Satisfactoria!"
Case Is = "NO"
     Label5 = "¡ATENCION! El contribuyente no ha sido encontrado en el padron, si corresponde debera aplicarle la tasa de retencion o percepcion Fija "
Case Is = "ER"
     Label5 = "¡ERROR! Numero de cuit con un formato incorrecto"
End Select

If cl_padronib.estado_embargo = "OK" Then
  Label9 = "¡¡¡ATENCION!!! CUIT CON EMBARGO DE ARBA"
Else
  Label9 = "CUIT SIN EMBARGO DE ARBA"
End If

Set cl_padronib = Nothing
End Sub



