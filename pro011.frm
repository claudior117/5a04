VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form prod_manual 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificacion Manual  del Registro de seguimiento de materiales "
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_idobra 
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox t_descprod 
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox t_idprod 
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   8295
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_obs 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1560
         Width           =   5895
      End
      Begin VB.TextBox t_c 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   14
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox c_cli 
         Height          =   360
         ItemData        =   "pro011.frx":0000
         Left            =   2160
         List            =   "pro011.frx":0010
         TabIndex        =   1
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Fecha:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Observaciones:"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Cantidad:"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Tipo Operacion:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_f 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox t_r 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox t_o 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox t_p 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   5
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Facturado"
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
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Recibidos"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "En OC"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Pedido"
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
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Num. referencia"
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Articulo"
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
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6720
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "pro011.frx":0075
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "pro011.frx":08F7
         Style           =   1  'Graphical
         TabIndex        =   11
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
      TabIndex        =   9
      Top             =   5970
      Width           =   8580
      _ExtentX        =   15134
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
            TextSave        =   "17/11/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:47"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label9 
      Caption         =   "CUIDADO! esta modifccion se realizará con una Minuta Interna y no tendrá un comprobante de respaldo."
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   6495
   End
End
Attribute VB_Name = "prod_manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnacepta_Click()
If verifica Then
  J = MsgBox("Confirma Operacion de modificacion", 4)
  If J = 6 Then
     Call graba
  End If
Else
  MsgBox ("No se actualizaron los Datos")
End If
End Sub

Sub graba()
cn1.BeginTrans

Select Case c_cli.ListIndex
 Case Is = 0
   q = "update pro_04 set [total_pedido] = " & Val(t_c) & " where [num_referencia] = " & Val(t_id)
 Case Is = 1
   q = "update pro_04 set [total_oc] = " & Val(t_c) & " where [num_referencia] = " & Val(t_id)
 Case Is = 2
   q = "update pro_04 set [total_recibido] = " & Val(t_c) & " where [num_referencia] = " & Val(t_id)
 Case Is = 3
   q = "update pro_04 set [total_facturado] = " & Val(t_c) & " where [num_referencia] = " & Val(t_id)
End Select
cn1.Execute q
   
   

   
' grabar movmiento
  numint = saca_ultnumero_int_comp("P")
      
  Set cl_compprod = New comprobantes_produccion
  cl_compprod.sacaultimonumero (2)
  If cl_compprod.numcomp > 0 Then
         QUERY = "INSERT INTO pro_01([num_int], [sucursal], [num_comprobante], [id_tipocomp], [id_obra], [fecha], [id_usuario], [fecha_esperado], [estado], [observaciones])"
         QUERY = QUERY & " VALUES (" & numint & ", " & glo.sucursal & ", " & cl_compprod.numcomp & ", 2, " & Val(t_idobra) & ", '" & t_fecha & "', " & para.id_usuario & ", '" & t_fecha & "', 'P', '" & RTrim$(t_obs) & " " & "')"
         cn1.Execute QUERY
      
           
            QUERY = "INSERT INTO pro_02([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [fecha_esperado], [observaciones], [num_referencia], [num_int_oc], [unidad])"
            QUERY = QUERY & " VALUES (" & numint & ", 1, " & Val(t_idprod) & ", '" & t_descripcion & "', " & Val(t_c) & ", '" & t_fecha & "', '" & RTrim$(t_obs) & " " & "', " & Val(t_id) & ", 0, ' ')"
            cn1.Execute QUERY
         
         
            q = "select * from pro_05 where [num_referencia] = " & Val(t_id)
            Set rs2 = New ADODB.Recordset
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs2.EOF And Not rs2.BOF Then
               rs2.MoveLast
               s = rs2("secuencia") + 1
            Else
               s = 1
            End If
            Set rs2 = Nothing
         
         
            QUERY = "INSERT INTO pro_05([num_referencia], [secuencia], [modulo], [num_int], [cantidad], [tipo_comprobante], [fecha], [unidad], [obs])"
            QUERY = QUERY & " VALUES (" & Val(t_id) & ", " & s & ", 'P', " & numint & ", " & Val(t_c) & ", 2, '" & t_fecha & "','  ','" & Left$(c_cli, 25) & "')"
            cn1.Execute QUERY
      
      J = MsgBox("Imprime Comprobante", 4)
      If J = 6 Then
         cl_compprod.cargar2 (numint)
         If cl_compprod.numint > 0 Then
           cl_compprod.imprimir
         End If
      End If
 End If
 
 cn1.CommitTrans
 
 Call verificaestados
 
 Unload Me

End Sub

Sub verificaestados()
Set rs = New ADODB.Recordset
q = "select * from pro_04 where [num_referencia] = " & Val(t_id)
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
  If rs("total_pedido") > rs("total_oc") Then
     rs("estado_pedido") = "I"
  Else
     rs("estado_pedido") = "C"
 End If
 
  If rs("total_oc") > rs("total_recibido") Then
     rs("estado_oc") = "I"
  Else
     rs("estado_oc") = "C"
 End If
 rs.Update
End If
Set rs = Nothing
 
  


End Sub
Private Sub btnsale_Click()
Me.Hide
End Sub

Function verifica() As Boolean
verifica = True
Select Case c_cli.ListIndex
Case Is = 0
  'pedido
  If Val(t_c) <= 0 Then
    MsgBox ("La cantidad pedda no puede ser menor o igual a cero")
    verifica = False
  End If
End Select

End Function




Private Sub c_cli_LostFocus()
If c_cli.ListIndex < 0 Then
  c_cli.ListIndex = 0
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  
         
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

c_cli.ListIndex = 0
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub




Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
      t_fecha = Format$(Now, "dd/mm/yyyy")
Else
      t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
'If verificaperiodo(t_fecha) = "C" Then
'   MsgBox ("El periodo para el cual se deseas ingresar el comprobante esta CERRADO!!!!!")
'   t_numop.SetFocus
'   t_fecha = ""
'End If
End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub
