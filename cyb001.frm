VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_abm_bancos1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BANCOS"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5070
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3495
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   8295
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2640
         Width           =   3615
      End
      Begin VB.ComboBox c_cuenta3 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   3615
      End
      Begin VB.ComboBox c_cuentad 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   6120
         Picture         =   "cyb001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   6120
         Picture         =   "cyb001.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox t_abreviatura 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   23
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id. Proveedor"
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
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Ch.Dif. a Pagar"
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
         Index           =   2
         Left            =   480
         TabIndex        =   21
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Ch.Dif. a Pagar"
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
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Abreviatura"
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
         Index           =   3
         Left            =   480
         TabIndex        =   17
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Contable Banco"
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
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "id. Forma de Pago"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Forma de Pago"
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
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6840
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "cyb001.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb001.frx":0E96
         Style           =   1  'Graphical
         TabIndex        =   8
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
      TabIndex        =   6
      Top             =   4815
      Width           =   8745
      _ExtentX        =   15425
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
Attribute VB_Name = "cyb_abm_bancos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   'On Error GoTo ERRORGRABA
    
   Select Case t_funcion
      
   Case "A"
      Set rs = New ADODB.Recordset
      q = "select * from g0  where [sucursal] = 0 "
      rs.Open q, cn1, adOpenStatic, adLockOptimistic
         
         nb = rs("ult_num_banco") + 1
         rs("ult_num_banco") = nb
         rs.Update
         
      Set rs = Nothing
      
      QUERY = "INSERT INTO cyb_01([id_forma_pago], [descripcion], [id_cuenta_cont], [abreviatura], [CAJA],[id_cuenta_ch_dif], [id_cuenta_dep_dif], [id_proveedor01])"
      QUERY = QUERY & " VALUES (" & nb & ", '" & t_descripcion & "', " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & t_abreviatura & "', 'N', " & c_cuentad.ItemData(c_cuentad.ListIndex) & ", " & c_cuenta3.ItemData(c_cuenta3.ListIndex) & ", " & c_prov.ItemData(c_prov.ListIndex) & ")"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   Case "M"
   
      QUERY = "update cyb_01 set  [Descripcion]='" & t_descripcion & "' , [id_cuenta_cont]=" & c_cuenta.ItemData(c_cuenta.ListIndex) & " , [abreviatura]='" & t_abreviatura & "', [caja]='N', [id_cuenta_ch_dif]=" & c_cuentad.ItemData(c_cuentad.ListIndex) & ", [id_cuenta_dep_dif]=" & c_cuenta3.ItemData(c_cuenta3.ListIndex) & ", [id_proveedor01]=" & c_prov.ItemData(c_prov.ListIndex)
      QUERY = QUERY & " where [id_forma_pago]= " & Val(t_id)
      'MsgBox (QUERY)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
   Case "B"
      QUERY = "DELETE FROM cyb_01 WHERE [id_forma_pago] = " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   End Select
   
   cyb_ABM_bancos.DataGrid1.Refresh
   cyb_ABM_bancos.Show
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

Private Sub c_cuenta3_LostFocus()
If c_cuenta3.ListIndex < 0 Then
  If Val(c_cuenta3) > 0 Then
    c_cuenta3.ListIndex = buscaindice(c_cuenta3, Val(c_cuenta3))
  Else
    c_cuenta3.ListIndex = 0
  End If
End If
End Sub

Private Sub c_cuentad_LostFocus()
If c_cuentad.ListIndex < 0 Then
  If Val(c_cuentad) > 0 Then
    c_cuentad.ListIndex = buscaindice(c_cuentad, Val(c_cuentad))
  Else
    c_cuentad.ListIndex = 0
  End If
End If
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
End If

If para.cuenta_sel > 0 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
  para.cuenta_sel = 0
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
    Call TabEnter2(Me, 5)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.ListIndex = 0
Call carga_cuentas_cont(c_cuentad, "C", "D")
c_cuenta.ListIndex = 0
Call carga_cuentas_cont(c_cuenta3, "C", "D")
c_cuenta3.ListIndex = 0
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
c_cuenta3.ListIndex = 0


para.cuenta_sel = 0
End Sub


Private Sub t_abreviatura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_abreviatura_LostFocus()
If IsNull(t_abreviatura) Then
 t_abreviatura = "*"
End If
End Sub

Private Sub t_descripcion_LostFocus()
If t_descripcion = "" Then
  t_descripcion = "Null"
End If
End Sub



