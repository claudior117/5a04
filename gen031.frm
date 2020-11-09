VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_cierremes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIERRE MENSUAL y VERIFICACION DE INTEGRIDAD  de CUENTAS"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_funcion 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Verificaciones"
      Height          =   4095
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   10575
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3615
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6376
         _Version        =   393216
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10575
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox c_estado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "gen031.frx":0000
         Left            =   1560
         List            =   "gen031.frx":000A
         TabIndex        =   0
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Mes:"
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
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   $"gen031.frx":0020
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Estado:"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Año:"
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
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10800
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "gen031.frx":01D2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen031.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   3
      Top             =   6915
      Width           =   12495
      _ExtentX        =   22040
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
End
Attribute VB_Name = "gen_cierremes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Dim fechacorte As String


Private Sub btnacepta_Click()
Call graba
End Sub
Sub estadomes()
fechacorte = DateSerial(Val(t_f2), Val(t_f1) + 1, 0)
If verificaperiodog(fechacorte) = "A" Then
    c_estado.ListIndex = 0
    Command1.Caption = "Cerrar Periodo"
    t_funcion = "C"
  Else
    c_estado.ListIndex = 1
    Command1.Caption = "Abrir Periodo"
    t_funcion = "A"
  End If
End Sub

Sub inicia()
  Frame3.Visible = True
  fechacorte = DateSerial(Val(t_f2), Val(t_f1) + 1, 0)
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 4
  msf1.FocusRect = flexFocusNone
  msf1.ColWidth(0) = 1000
  msf1.ColWidth(1) = 3000
  msf1.ColWidth(2) = 2000
  msf1.ColWidth(3) = 2000
  msf1.TextMatrix(0, 0) = "Estado"
  msf1.TextMatrix(0, 1) = "Cuenta"
  msf1.TextMatrix(0, 2) = "Saldo Movimientos"
  msf1.TextMatrix(0, 3) = "Saldo Contable"
  Call estadomes
  Command1.Visible = True

  
End Sub
Sub graba()
If verifica Then
 J = MsgBox("El proceso de verificacion de Integridad de Cuentas puede demorar. Salga del Sistema en todas las terminales y Confirme", 4)
 If J = 6 Then
    Call inicia
    espere.Show
    espere.ProgressBar1.Max = 4
    espere.Label1 = "Inicializando....."
    espere.Refresh
    Call deudores
    Call acreedores
    Call vericaja
    Call veribanco
    Call veriinventario
    
    Unload espere
 End If
End If
End Sub
Sub deudores()
   espere.ProgressBar1.Value = 1
   espere.Label1 = "Verificando cuenta Deudores. Espere...."
   espere.Refresh
   Set rsd = New ADODB.Recordset
   q = "select [id_cliente] from vta_01 where [id_cliente] > 1 and [saldo_incobrable] = 'N'"
   rsd.Open q, cn1
   s = 0
   While Not rsd.EOF
     Set cl_cli = New Clientes
     cl_cli.carga (rsd("id_CLIENTE"))
     s = s + cl_cli.saldo(True, fechacorte, True)
     Set cl_cli = Nothing
     rsd.MoveNext
   Wend
   Set rsd = Nothing
   
   s2 = saldocuentacgr(para.cuenta_deudores, fechacorte)
   
   If Val(Format$(s, "######0.00")) = Val(Format$(s2, "######0.00")) Then
      e = "OK"
   Else
      e = "ERR"
   End If
   msf1.AddItem e & Chr$(9) & "Cuenta Deudores" & Chr$(9) & s & Chr$(9) & s2
   msf1.Refresh
   
End Sub

Sub acreedores()
   espere.ProgressBar1.Value = 2
   espere.Label1 = "Verificando cuenta Acreedores. Espere...."
   espere.Refresh
   Set rsd = New ADODB.Recordset
   q = "select [id_proveedor] from a1 where [id_proveedor] > 1 "
   rsd.Open q, cn1
   s = 0
   While Not rsd.EOF
     Set cl_prov = New proveedores
     cl_prov.carga (rsd("id_proveedor"))
     s = s + cl_prov.saldo(True, fechacorte, True, 0)
     Set cl_prov = Nothing
     rsd.MoveNext
   Wend
   Set rsd = Nothing
   
   s2 = saldocuentacgr(para.cuenta_acreedores, fechacorte)
       
   If Val(Format$(s, "######0.00")) = Val(Format$(s2, "######0.00")) Then
      e = "OK"
   Else
      e = "ERR"
   End If
   msf1.AddItem e & Chr$(9) & "Cuenta Acreedores" & Chr$(9) & Format$(s, "######0.00") & Chr$(9) & Format$(s2, "######0.00")
   msf1.Refresh
   
End Sub


Sub vericaja()
   espere.ProgressBar1.Value = 3
   espere.Label1 = "Verificando cuentas Caja. Espere...."
   espere.Refresh
   
   Set rs = New ADODB.Recordset
   q = "select * from cyb_01 where [caja] = 'S'"
   rs.Open q, cn1
   While Not rs.EOF
     Set rs1 = New ADODB.Recordset
     q = "select * from cyb_05 where [id_forma_pago] = " & rs("id_forma_pago") & " and datevalue([fecha]) <= datevalue('" & fechacorte & "')"
     rs1.Open q, cn1
     sd = 0
     While Not rs1.EOF
       If rs1("UBICACION") = "D" Then
         sd = sd + rs1("importe")
       Else
         sd = sd - rs1("importe")
       End If
        
      rs1.MoveNext
    Wend
    Set rs1 = Nothing
    s2 = saldocuentacgr(rs("id_cuenta_cont"), fechacorte)
    d = rs("descripcion")
    If Val(Format$(sd, "######0.00")) = Val(Format$(s2, "######0.00")) Then
      e = "OK"
   Else
      e = "ERR"
   End If
   msf1.AddItem e & Chr$(9) & d & Chr$(9) & Format$(sd, "######0.00") & Chr$(9) & Format$(s2, "######0.00")
   msf1.Refresh
    
     rs.MoveNext
 Wend
   Set rs = Nothing
End Sub

Sub veriinventario()
   espere.ProgressBar1.Value = 4
   espere.Label1 = "Verificando Cuentas Inventario. Espere...."
   espere.Refresh
   Set rs = New ADODB.Recordset
   q = "select [id_producto], [costoreal] from a2 where [id_producto] > 1 "
   rs.Open q, cn1
   sd = 0
   sp = 0
   Set cl_stock = New STOCK
   While Not rs.EOF
      Call cl_stock.sacastock(rs("id_producto"), fechacorte)
      sp = Format$(cl_stock.stock_movimientos, "#######0.00")
      If IsNull(sp) Then
        sp = "0.00"
      End If
      sd = sd + (Val(sp) * rs("costoreal"))
      rs.MoveNext
   Wend
   Set cl_stock = Nothing
   Set rs = Nothing
   s2 = saldocuentacgr(para.cuenta_inventario, fechacorte)
   
   If Val(Format$(sd, "######0.00")) = Val(Format$(s2, "######0.00")) Then
      e = "OK"
   Else
      e = "ERR"
   End If
   msf1.AddItem e & Chr$(9) & "Inventario" & Chr$(9) & Format$(sd, "######0.00") & Chr$(9) & Format$(s2, "######0.00")
   msf1.Refresh
  End Sub


Sub veribanco()
   espere.ProgressBar1.Value = 3
   espere.Label1 = "Verificando cuentas Bancos. Espere...."
   espere.Refresh
   
   Set rs = New ADODB.Recordset
   q = "select * from cyb_01 where [id_forma_pago] >= 50"
   rs.Open q, cn1
   While Not rs.EOF
     Set rs1 = New ADODB.Recordset
     q = "select * from cyb_04 where [id_banco] = " & rs("id_forma_pago") & " and datevalue([fecha]) <= datevalue('" & fechacorte & "' and [entro] = 'S')"
     rs1.Open q, cn1
     sd = 0
     While Not rs1.EOF
       If rs1("UBICACION") = "D" Then
         sd = sd + rs1("importe")
       Else
         sd = sd - rs1("importe")
       End If
      rs1.MoveNext
    Wend
    Set rs1 = Nothing
    s2 = saldocuentacgr(rs("id_cuenta_cont"), fechacorte)
    d = rs("descripcion") & " conciliado"
    
   If Val(Format$(sd, "######0.00")) = Val(Format$(s2, "######0.00")) Then
      e = "OK"
   Else
      e = "ERR"
   End If
   msf1.AddItem e & Chr$(9) & d & Chr$(9) & Format$(sd, "######0.00") & Chr$(9) & Format$(s2, "######0.00")
   msf1.Refresh
  rs.MoveNext
 Wend
 Set rs = Nothing
   
   
   
   
   
   
End Sub

Function verifica() As Boolean
v = 1
If Val(t_f1) <= 0 Or Val(t_f1) > 12 Then
     MsgBox ("Mes Incorrecto")
     v = 0
End If

If Val(t_f2) < 2008 Or Val(t_f2) > Val(Mid$(Format$(Now, "dd/mm/yyyy"), 7, 4)) Then
     MsgBox ("Año Incorrecto")
     v = 0
End If

If v = 0 Then
  verifica = False
Else
 verifica = True
End If

End Function
Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub c_estado_Change()
If c_estado.ListIndex < 0 Then
  c_estado.ListIndex = 0
End If

End Sub


Private Sub c_estado_LostFocus()
If c_estado.ListIndex < 0 Then
  c_estado.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
J = MsgBox("Confirma " & Command1.Caption, 4)
If J = 6 Then
    If t_funcion = "C" Then
      QUERY = "INSERT INTO g10([periodo], [estado], [id_usuario])"
      QUERY = QUERY & " VALUES (" & Val(Format$(t_f2, "0000") & Format$(t_f1, "00")) & ", 'C', " & para.id_usuario & ")"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
    
    Else
      QUERY = "update g10 set  [estado]='A', [id_usuario]=" & para.id_usuario
      QUERY = QUERY & " where [periodo]= " & Val(Format$(t_f2, "0000") & Format$(t_f1, "00"))
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
    End If
End If

End Sub

Private Sub Form_Load()
Frame3.Visible = False
Command1.Visible = False
t_funcion = ""
t_f1 = Format$(Val(Mid$(Format$(Now, "dd/mm/yyyy"), 4, 2)), "00")
t_f2 = Format$(Val(Mid$(Format$(Now, "dd/mm/yyyy"), 7, 4)), "0000")
If verificaperiodog(Now) = "A" Then
  c_estado.ListIndex = 0
Else
  c_estado.ListIndex = 1
End If


End Sub



Private Sub t_f1_Change()
Call estadomes
End Sub

Private Sub t_f2_Change()
Call estadomes
End Sub

Private Sub UpDown1_DownClick()
If Val(t_f1) > 1 Then
  t_f1 = Val(t_f1) - 1
End If

End Sub

Private Sub UpDown1_UpClick()
If Val(t_f1) < 12 Then
  t_f1 = Val(t_f1) + 1
End If

End Sub

Private Sub UpDown2_DownClick()
If Val(t_f2) > 2005 Then
  t_f2 = Val(t_f2) - 1
End If
End Sub

Private Sub UpDown2_UpClick()
If Val(t_f2) < 2050 Then
  t_f2 = Val(t_f2) + 1
End If

End Sub
