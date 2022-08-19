VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_movint 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MOVIMIENTOS DE AJUESTE EN STOCK"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Ajuste Definitivo del Stock en $"
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   7680
      Width           =   2655
      Begin VB.TextBox t_ajuste 
         Height          =   405
         Left            =   240
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   9000
      TabIndex        =   14
      Top             =   0
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"stk006.frx":0000
         Height          =   1215
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4935
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox c_cli 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1800
         Width           =   5775
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   8160
         Picture         =   "stk006.frx":0095
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   2
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox t_numoc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Stock por Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta contra inventario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Detalle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Ingreso:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Movimiento Interno Nro.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   8
      Top             =   7560
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk006.frx":039F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "stk006.frx":0C21
         Style           =   1  'Graphical
         TabIndex        =   9
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
      TabIndex        =   7
      Top             =   8550
      Width           =   11985
      _ExtentX        =   21140
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
            TextSave        =   "19/08/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:47 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_movint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 4500
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 3200
msf1.ColWidth(5) = 800
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1000
msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Descripcion"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Detalle"
msf1.TextMatrix(0, 5) = "Tipo"
msf1.TextMatrix(0, 6) = "Costo Unit."
msf1.TextMatrix(0, 7) = "Costo Tot."



End Sub







Private Sub c_cli_LostFocus()
If c_cli.ListIndex < 0 Then
 c_cli.ListIndex = 0
End If

End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
If para.cuenta_sel > 0 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
End If
End Sub

Function sacatotales()
 J = 1
 t = 0
 For i = J To msf1.Rows - 1
  If msf1.TextMatrix(i, 5) = "E" Then
    t = t + Val(msf1.TextMatrix(i, 7))
  Else
     t = t - Val(msf1.TextMatrix(i, 7))
  End If
 
 Next i
 sacatotales = t
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call armagrid
Call barraesag(Me)
Call carga_cuentas_cont(c_cuenta, "C", "C")
c_cuenta.AddItem "<Sin Imputar>", 0
c_cuenta.ListIndex = 0

Call carga_clientes(c_cli)
c_cli.AddItem "<Todos>", 0
c_cli.ListIndex = 0

t_ajuste = "0.00"
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
    t_ajuste = sacatotales()
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF9 Then
 If msf1.Rows > 1 Then
  t_ajuste = sacatotales()
  J = MsgBox("Confirma Grabar Movimiento Interno ", 4)
  If J = 6 Then
   If verificaperiodog(t_fecha) = "A" Then
     Call graba
   Else
     MsgBox ("Periodo cerrado. Imposible grabar operacion")
   End If
  End If
 End If
End If

If KeyCode = vbKeyInsert Then
   stk_movint2.t_renglon = ""
   stk_movint2.t_cantidad = ""
   stk_movint2.Show
End If
End Sub

Sub graba()
If EXISTE = "S" Then
   'borro mov stock
    Set cl_stock = New STOCK
    Call cl_stock.borra_mov_stk(Val(t_numoc), "S")
    EXISTE = "N"
End If


If EXISTE = "N" Then
   'oc nueva
      'On Error GoTo ERRORGRABA
      
      Set rs = New ADODB.Recordset
      q = "Select * from g0 where sucursal=0"
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      numcomp = rs("ult_num_ajuste_stock") + 1
      rs("ult_num_ajuste_stock") = numcomp
      rs.Update
      Set rs = Nothing
      
      cn1.BeginTrans
      QUERY = "INSERT INTO stk_02([fecha], [letra], [num_comprobante], [id_usuario], [detalle], [sucursal], [tipo_comprobante], [id_proveedor], [id_obra])"
      QUERY = QUERY & " VALUES ('" & t_fecha & "', 'X', " & numcomp & ", " & para.id_usuario & ", '" & t_detalle & " ', 0, 1, 1,1)"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      numint = rs.Fields("NewID").Value

      If c_cli.ListIndex = 0 Then
         idcli = 0
      Else
         idcli = c_cli.ItemData(c_cli.ListIndex)
     End If
      
      
      For i = 1 To msf1.Rows - 1
        QUERY = "INSERT INTO stk_03([num_int], [RENGLON], [id_producto], [descripcion], [unidad], [detalle], [cantidad], [ubicacion])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', ' ', '" & msf1.TextMatrix(i, 4) & "', " & msf1.TextMatrix(i, 3) & " , '" & msf1.TextMatrix(i, 5) & "')"
        cn1.Execute QUERY
      
        QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo], [id_cliente])"
        QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & msf1.TextMatrix(i, 3) & ", '" & msf1.TextMatrix(i, 5) & "', 'Mov.Int.Stk " & Format$(numint, "00000000") & _
        "', '" & RTrim$(msf1.TextMatrix(i, 4)) & " ', " & numint & ", 'S', " & idcli & ")"
        cn1.Execute QUERY
      
        If msf1.TextMatrix(i, 5) = "E" Then
          QUERY = "update a2 set  [stock]=[stock] + " & Val(msf1.TextMatrix(i, 3))
          QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
          cn1.Execute QUERY
        Else
          QUERY = "update a2 set  [stock]=[stock] - " & Val(msf1.TextMatrix(i, 3))
          QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
          cn1.Execute QUERY
        End If
      
      Next i
      
      
      If c_cuenta.ListIndex > 0 And Val(t_ajuste) <> 0 Then
        'grabo asiento
           'graba asiento
            numintcgr = saca_ultnumero_int_comp("G")
            cta = para.cuenta_inventario
         
            If Val(t_ajuste) > 0 Then
            'inventario al debe porque es un aumento
               u2 = "D"
               u1 = "H"
               importe = Val(t_ajuste)
             Else
             'inventario al haber porque es una disminucion
               u2 = "H"
               u1 = "D"
               importe = -Val(t_ajuste)
             End If
         
         
             Set rs = New ADODB.Recordset
             q = "select [descripcion] from c_01 where [id_cuenta] = " & cta
             rs.MaxRecords = 1
             rs.Open q, cn1
             If Not rs.EOF And Not rs.BOF Then
               dcta = rs("descripcion")
             Else
               dcta = "Cuenta Inexistente"
             End If
             Set rs = Nothing
         
             'grabo asiento
    
            QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
            QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Stock] Ajuste " & Format$(numint, "00000000") & "', 'S', " & numint & ", " & importe & ", " & importe & ", " & para.id_usuario & ", '" & Left$(RTrim$(t_detalle), 50) & "')"
            cn1.Execute QUERY
      
           'cuenta madre
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u2 & "', " & importe & ", 'Ajuste " & Format$(numint, "00000000") & "')"
           cn1.Execute QUERY
         
            ic = 2
            cta = c_cuenta.ItemData(c_cuenta.ListIndex)
            QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
            QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & importe & ", 'Ajuste " & Format$(numint, "00000000") & "')"
            cn1.Execute QUERY
 
      End If
      
      
      cn1.CommitTrans
      Set rs = Nothing
      
      J = MsgBox("Imprime Movimiento Interno", 4)
      If J = 6 Then
         Set cl_stock = New STOCK
         cl_stock.imprimir (numint)
         Set cl_stock = Nothing
      End If

      
      Call INICIALIZA2(Me)
      Call armagrid
      t_numoc.SetFocus
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    stk_movint2.t_renglon = msf1.Row
    stk_movint2.t_basico = msf1.TextMatrix(msf1.Row, 1)
    stk_movint2.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    stk_movint2.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    stk_movint2.t_renglonp = msf1.TextMatrix(msf1.Row, 5)
    stk_movint2.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
End Sub


Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub

Private Sub t_numoc_GotFocus()
Call INICIALIZA2(Me)
End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
 Call carga_oc
End Sub
Sub carga_oc()
q = "SELECT * FROM STK_02, STK_03 WHERE STK_02.[NUM_INT] = " & Val(t_numoc) & " and [tipo_comprobante] = 1 AND STK_02.[NUM_INT] = STK_03.[NUM_INT]"
Set rs = New ADODB.Recordset
rs.Open q, cn1
r = 0
If Not rs.EOF And Not rs.BOF Then
    EXISTE = "S"
    t_fecha = rs("FECHA")
    t_detalle = rs("stk_02.DETALLE")
    Call armagrid
    r = 1
    While Not rs.EOF
      msf1.AddItem r & Chr$(9) & rs("id_producto") & Chr$(9) & rs("descripcion") & Chr$(9) & rs("cantidad") & Chr$(9) & rs("stk_03.detalle") & Chr$(9) & rs("ubicacion")
      r = r + 1
      rs.MoveNext
    Wend
Else
  EXISTE = "N"
  Call armagrid
  t_fecha = ""
  t_detalle = ""
End If

End Sub
