VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form abm_prov1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROVEEDORES"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   31
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   7695
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "c_cuenta"
         Top             =   6120
         Width           =   4815
      End
      Begin VB.ComboBox c_provincia 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   4815
      End
      Begin VB.TextBox t_transp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   16
         Top             =   5760
         Width           =   495
      End
      Begin VB.TextBox t_alicuotaretib 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   14
         Top             =   5040
         Width           =   1215
      End
      Begin VB.ComboBox c_retib 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox t_contacto 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   18
         Top             =   6480
         Width           =   6015
      End
      Begin VB.TextBox t_tecontacto 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   19
         Top             =   6840
         Width           =   3375
      End
      Begin VB.TextBox t_emailcontacto 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   149
         TabIndex        =   20
         Top             =   7200
         Width           =   6015
      End
      Begin VB.TextBox t_fecha_vto_exepcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         MaxLength       =   13
         TabIndex        =   12
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox t_numib 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   15
         Top             =   5400
         Width           =   2655
      End
      Begin VB.ComboBox c_ib 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4680
         Width           =   2655
      End
      Begin VB.TextBox t_inscgan 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   9
         Top             =   3960
         Width           =   495
      End
      Begin VB.ComboBox c_ret 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox t_cuit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   10
         Top             =   4320
         Width           =   2655
      End
      Begin VB.ComboBox c_iva 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox t_email 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   6
         Top             =   2880
         Width           =   6015
      End
      Begin VB.TextBox t_cp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox t_localidad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox t_direccion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox t_te 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   0
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ingresar Cuit sin guines"
         Height          =   255
         Left            =   5040
         TabIndex        =   52
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cuenta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   51
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Emp. Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   480
         TabIndex        =   50
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Alicuota(%) Ret. IB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   49
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cod.Ret.IB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   480
         TabIndex        =   48
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Te. Contacto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   47
         Top             =   6840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Email Contacto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   46
         Top             =   7200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   45
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Fecha Vto. Exepcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   5160
         TabIndex        =   44
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Numero IBBA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   43
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Insc. IBBA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   42
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Inscrip. Gananc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   480
         TabIndex        =   41
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cod. Ret. Gan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   40
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cod. Postal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   39
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cuit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   38
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   37
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Resp. ante IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   36
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   35
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Localidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Telefonos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "id. Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Razon Social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   22
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "Arch001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Arch001.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   23
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
      TabIndex        =   21
      Top             =   8265
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "24/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:20 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label13 
      Caption         =   "Los campos marcados con ROJO no estan siendo utilizados actualmente"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   7920
      Width           =   5535
   End
End
Attribute VB_Name = "abm_prov1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String


Private Sub btnacepta_Click()
If verifica Then
 Call graba
End If
End Sub
Function verifica() As Boolean
v = True

'verifico que   el cuit no este en el sistema
  If t_funcion = "A" And t_cuit <> "0" Then
   q = "select * from a1 where val([cuit]) = " & Val(t_cuit)
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
     J = MsgBox("El numero de cuit ya fue ingresado en el sistema en la cuenta [" & rs("denominacion") & "]. ¿Desea igualmente crea una nueva cuenta?", 4)
     If J <> 6 Then
        v = False
     End If
   End If
   Set rs = Nothing
  End If

verifica = v
End Function
Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   On Error GoTo ERRORGRABA
    
   Select Case t_funcion
      
   Case "A"
  QUERY = "INSERT INTO a1([DEnominacion], [direccion], [te], [localidad], [cp], [provincia], [email], [cuit], [cod_tipoiva], [id_codretgan], [inscripto_gan], [id_tipoib], [num_ib], [fecha_vto_exepcion_ib], [contacto], " & _
  "[te_contacto], [email_contacto], [id_codretib], [alicuota_retib], [transporte], [id_provincia], [id_cuenta_a1])"
  QUERY = QUERY & " VALUES ('" & t_descripcion & "', '" & t_direccion & "', '" & t_te & "', '" & t_localidad & "', '" & t_cp & "', '" & c_provincia & "', '" & t_email & "', '" & t_cuit & "', " & _
  c_iva.ItemData(c_iva.ListIndex) & ", " & c_ret.ItemData(c_ret.ListIndex) & ", '" & t_inscgan & "', " & c_ib.ItemData(c_ib.ListIndex) & ", '" & RTrim$(t_numib) & " ', '" & t_fecha_vto_exepcion & _
  "', '" & t_contacto & "', '" & t_tecontacto & "', '" & t_emailcontacto & "', " & c_retib.ItemData(c_retib.ListIndex) & ", " & Val(t_alicuotaretib) & ", '" & t_transp & "', " & c_provincia.ItemData(c_provincia.ListIndex) & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ")"
        
        cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
  Case "M"

QUERY = "update a1 set  [Denominacion]='" & t_descripcion & "' , [direccion]='" & t_direccion & "' , [te]='" & t_te & "' , [localidad]='" & t_localidad & "' , [cp]='" & t_cp & "' , [provincia]='" & c_provincia & _
"' , [email]='" & t_email & "' , [cuit]='" & t_cuit & "' , [cod_tipoiva]=" & c_iva.ItemData(c_iva.ListIndex) & " , [id_codretgan] = " & c_ret.ItemData(c_ret.ListIndex) & " , [inscripto_gan] = '" & t_inscgan & _
"' , [id_tipoib] = " & c_ib.ItemData(c_ib.ListIndex) & " , [num_ib] = '" & RTrim$(t_numib) & " ' , [fecha_vto_exepcion_ib] = '" & t_fecha_vto_exepcion & " ' , [contacto] = '" & t_contacto & _
"' , [te_contacto] = '" & t_tecontacto & "' , [email_contacto] = '" & t_emailcontacto & "' , [id_codretib] = " & c_retib.ItemData(c_retib.ListIndex) & " , [alicuota_retib] = " & Val(t_alicuotaretib) & _
" , [transporte] = '" & t_transp & "' , [id_provincia]=" & c_provincia.ItemData(c_provincia.ListIndex) & ", [id_cuenta_a1]=" & c_cuenta.ItemData(c_cuenta.ListIndex)
QUERY = QUERY & " where [id_proveedor]= " & Val(t_id)
      
       
      
      
      
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
   Case "B"
      QUERY = "DELETE FROM a1 WHERE [id_proveedor] = " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   End Select
   ABM_PROv.Show
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub



Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
End Sub

Private Sub c_provincia_GotFocus()
If t_cp = "2705" Then
 c_provincia.ListIndex = 1
End If

End Sub

Private Sub c_provincia_LostFocus()
If c_provincia.ListIndex < 0 Then
  c_provincia.ListIndex = 0
End If
End Sub

Private Sub c_retib_LostFocus()
If c_retib.ListIndex < 0 Then
   c_retib.ListIndex = 0
End If

End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 20)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_tipoiva(c_iva)
Call carga_impuestos(c_ret, 217) 'ret. ganancias
Call carga_tipoib(c_ib)
Call carga_impuestos(c_retib, 50) 'ret. ganancias
Call carga_provincias(c_provincia)
Call carga_cuentas_cont(c_cuenta, "C", "A")

End Sub

Private Sub t_alicuotaretib_LostFocus()
If Val(t_alicuotaretib) < 0 Then
   t_alicuotaretib = "0.00"
End If
End Sub

Private Sub t_contacto_LostFocus()
Call NULOS(t_contacto)
End Sub

Private Sub t_cp_LostFocus()
If t_localidad = "Rojas" And t_cp = "" Then
  t_cp = "2705"
Else
  Call NULOS(t_cp)
End If
End Sub

Private Sub t_cuit_LostFocus()
If t_cuit = "" Then
  t_cuit = "0"
Else
  Call NULOS(t_cuit)
End If
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub



Private Sub t_direccion_LostFocus()
Call NULOS(t_direccion)
End Sub

Private Sub t_email_LostFocus()
Call NULOS(t_email)
End Sub

Private Sub t_emailcontacto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If

End Sub

Private Sub t_emailcontacto_LostFocus()
Call NULOS(t_emailcontacto)
End Sub

Private Sub t_fecha_vto_exepcion_LostFocus()
  If Not IsNull(t_fecha_vto_exepcion) Then
    If Not IsDate(t_fecha_vto_exepcion) Then
       t_fecha_vto_exepcion = Format$(Date, "dd/mm/yyyy")
    End If
  Else
    t_fecha_vto_exepcion = Format$(Date, "dd/mm/yyyy")
  End If
End Sub

Private Sub t_inscgan_GotFocus()
StatusBar1.Panels.item(2) = "[S] Inscripto - [N] No Inscripto - [E] Exento"

End Sub

Private Sub t_inscgan_LostFocus()
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
t_inscgan = UCase(t_inscgan)
If t_inscgan <> "S" And t_inscgan <> "N" And t_inscgan <> "E" Then
  t_inscgan = "N"
End If
Call barraesag(Me)
End Sub

Private Sub t_localidad_LostFocus()
If t_localidad = "" Then
 t_localidad = "Rojas"
Else
 Call NULOS(t_localidad)
End If
End Sub

Private Sub t_numib_LostFocus()
Call NULOS(t_numib)
End Sub



Private Sub t_te_LostFocus()
Call NULOS(t_te)
End Sub

Private Sub t_tecontacto_LostFocus()
 Call NULOS(t_tecontacto)
End Sub

Private Sub t_transp_LostFocus()
t_transp = Format$(t_transp, ">@")
If t_transp <> "S" Then
  t_transp = "N"
End If
End Sub
