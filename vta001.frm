VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_abm_cli1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIENTES"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   27
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   28
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
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   7815
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8895
      Begin VB.TextBox t_direccionlocal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   15
         Top             =   6360
         Width           =   5895
      End
      Begin VB.ComboBox c_provincia 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   5880
         Width           =   3615
      End
      Begin VB.TextBox t_observaciones 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   45
         TabIndex        =   13
         Top             =   5520
         Width           =   5175
      End
      Begin VB.TextBox t_saldoi 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   12
         Top             =   5160
         Width           =   375
      End
      Begin VB.TextBox t_percib 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   11
         Top             =   4800
         Width           =   375
      End
      Begin VB.TextBox t_operadorgranos 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   10
         Top             =   4440
         Width           =   375
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox t_credito 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   14
         TabIndex        =   8
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox t_cuit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   16
         Top             =   6720
         Width           =   2295
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
         TabIndex        =   22
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
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Direccion Local"
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
         TabIndex        =   49
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Identificar si el cliente es proveedor"
         Height          =   255
         Left            =   5880
         TabIndex        =   48
         Top             =   5880
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe maximo aceptado en pesos ($)"
         Height          =   255
         Left            =   3840
         TabIndex        =   47
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   $"vta001.frx":0000
         Height          =   855
         Left            =   4560
         TabIndex        =   46
         Top             =   6720
         Width           =   3855
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id. Proveedor:"
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
         Height          =   495
         Index           =   10
         Left            =   480
         TabIndex        =   45
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Observaciones"
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
         TabIndex        =   44
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "[S] Si  -  [N] No"
         Height          =   255
         Left            =   2640
         TabIndex        =   43
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "[S] Si  -  [N] No"
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "[S] Si  -  [N] No"
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Saldo Incobrable"
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
         Left            =   480
         TabIndex        =   40
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Percive IB"
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
         TabIndex        =   39
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Insc. Registro Operador Granos"
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
         Height          =   495
         Index           =   6
         Left            =   480
         TabIndex        =   38
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Vendedor"
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
         TabIndex        =   37
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Limite Credito"
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   6720
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Localiadad"
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
         TabIndex        =   30
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
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "id. Cliente"
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Direccion Fiscal"
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
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   18
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "vta001.frx":00BA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta001.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   19
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
      TabIndex        =   17
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_abm_cli1"
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
If c_iva.ItemData(c_iva.ListIndex) <> 3 And c_iva.ItemData(c_iva.ListIndex) <> 8 Then
  If para.fiscal = 1 Then
   If verificacuit(t_cuit) = 0 Then
      MsgBox ("El Numero de CUIT es Incorrecto")
      v = False
   End If
  End If

  'verifico que   el cuit no este en el sistema
  If t_funcion = "A" Then
   q = "select * from vta_01 where val([cuit]) = " & Val(t_cuit)
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
Else
  If Val(t_cuit) < 0 Then
       MsgBox ("El DNI es Incorrecto")
       v = False
  End If
End If

verifica = v
 
End Function
Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   If c_prov.ListIndex > 0 Then
      iprov = c_prov.ItemData(c_prov.ListIndex)
   Else
      iprov = 0
   End If
   
   
   'On Error GoTo ERRORGRABA
    
   Select Case t_funcion
      
   Case "A"
      cn1.BeginTrans
      QUERY = "INSERT INTO vta_01([DEnominacion], [direccion], [te], [localidad], [cp], [provincia], [email], [cuit], [id_tipoiva], [limite_credito], [id_vendedor], [inscripto_operador_granos], [percive_ib], [saldo_incobrable], [observaciones], [id_proveedor], [id_prov], [direccion_local] )"
      QUERY = QUERY & " VALUES ('" & t_descripcion & "', '" & t_direccion & "', '" & t_te & "', '" & t_localidad & "', '" & t_cp & "', '" & c_provincia & "', '" & t_email & "', '" & t_cuit & "', " & c_iva.ItemData(c_iva.ListIndex) & ", " & Val(t_credito) & ", " & c_vend.ItemData(c_vend.ListIndex) & ", '" & t_operadorgranos & "', '" & T_PERCIB & "', '" & t_saldoi & "', '" & t_observaciones & "', " & iprov & ", " & c_provincia.ItemData(c_provincia.ListIndex) & ", '" & t_direccionlocal & "')"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      nc = rs.Fields("NewID").Value

      
      QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
      QUERY = QUERY & " VALUES ('Alta Cliente ', " & para.id_usuario & ", 'V', 0, '" & Now & "', '" & Left$(t_descripcion, 50) & "',1," & nc & ")"
     ' MsgBox (QUERY)
      cn1.Execute QUERY
      
      cn1.CommitTrans
   
   
   Case "M"
      cn1.BeginTrans
    QUERY = "update vta_01 set  [Denominacion]='" & t_descripcion & "' , [direccion]='" & t_direccion & "' , [te]='" & t_te & "' , [localidad]='" & t_localidad & "' , [cp]='" & t_cp & "' , [provincia]='" & c_provincia & _
    "' , [email]='" & t_email & "' , [cuit]='" & t_cuit & "' , [id_tipoiva]=" & c_iva.ItemData(c_iva.ListIndex) & " , [limite_credito] = " & Val(t_credito) & " , [id_vendedor]=" & c_vend.ItemData(c_vend.ListIndex) & _
    " , [inscripto_operador_granos]='" & t_operadorgranos & "' , [percive_ib]='" & T_PERCIB & "' , [saldo_incobrable]='" & t_saldoi & "' , [observaciones]='" & t_observaciones & "' , [id_proveedor]=" & iprov & _
    " , [id_prov]=" & c_provincia.ItemData(c_provincia.ListIndex) & " , [direccion_local]='" & t_direccionlocal & "'"
    QUERY = QUERY & " where [id_cliente]= " & Val(t_id)
      
    cn1.Execute QUERY
      
    QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
    QUERY = QUERY & " VALUES ('Modifica Cliente " & t_id & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', '" & Left$(t_descripcion, 50) & "', 1, " & Val(t_id) & ")"
      
    cn1.Execute QUERY
      
      
      
      cn1.CommitTrans
      
   Case "B"
      q = "select * from vta_02 where [id_cliente] = " & Val(t_id)
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
          MsgBox ("El cliente tiene movimientos cargados. No se puede Eliminar")
      Else
          QUERY = "DELETE FROM vta_01 WHERE [id_cliente] = " & Val(t_id)
          cn1.BeginTrans
          cn1.Execute QUERY
          
          
          QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
          QUERY = QUERY & " VALUES ('Borra Cliente " & t_id & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', '" & Left$(t_descripcion, 50) & "', 1, " & Val(t_id) & ")"
      
          cn1.Execute QUERY
              
          cn1.CommitTrans
      End If
      Set rs = Nothing
   
   
   End Select
   
   
   vta_ABM_cli.Show
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



Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
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

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
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
    Call TabEnter2(Me, 16)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_tipoiva(c_iva)
Call carga_vendedores(c_vend)
Call carga_proveedores(c_prov)
c_prov.AddItem "<No es proveedor>", 0
c_prov.ListIndex = 0
Call carga_provincias(c_provincia)

End Sub

Private Sub t_cp_LostFocus()
If t_cp = "" Then
  t_cp = "2705"
Else
  Call NULOS(t_cp)
End If
End Sub

Private Sub t_credito_LostFocus()
If t_credito = "" Then
  t_credito = "999999.99"
Else
  t_credito = Format$(Val(t_credito), "######0.00")
End If
End Sub

Private Sub t_cuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
Else
 Call solonum(KeyAscii, 0)
End If
End Sub

Private Sub t_cuit_LostFocus()
If t_cuit <> "" Then
 Call NULOS(t_cuit)
Else
  t_cuit = "0"
End If
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub



Private Sub t_direccion_LostFocus()
Call NULOS(t_direccion)
End Sub

Private Sub t_direccionlocal_LostFocus()
If t_direccionlocal = "" Then
  t_direccionlocal = t_direccion
End If
End Sub

Private Sub t_email_LostFocus()
Call NULOS(t_email)
End Sub

Private Sub t_localidad_LostFocus()
If t_localidad = "" Then
  t_localidad = "Rojas"
Else
  Call NULOS(t_localidad)
End If
End Sub

Private Sub t_operadorgranos_LostFocus()
t_operadorgranos = Format$(t_operadorgranos, ">@")
Select Case t_operadorgranos
Case Is = "S", Is = "N"

Case Else
  t_operadorgranos = "N"
End Select

End Sub

Private Sub t_percib_LostFocus()
T_PERCIB = Format$(T_PERCIB, ">@")
Select Case T_PERCIB
Case Is = "S", Is = "N"

Case Else
  T_PERCIB = "S"
End Select

End Sub


Private Sub t_saldoi_LostFocus()
t_saldoi = Format$(t_saldoi, ">@")
Select Case t_saldoi
Case Is = "S", Is = "N"

Case Else
  t_saldoi = "N"
End Select

End Sub

Private Sub t_te_LostFocus()
Call NULOS(t_te)
End Sub
