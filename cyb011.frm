VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_movcaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos Internos de Caja"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Reimprimir"
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operacion"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   4320
      Width           =   3735
      Begin VB.TextBox t_op 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4680
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "cyb011.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb011.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Movimiento"
      Height          =   3375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   6255
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   5160
         Picture         =   "cyb011.frx":1104
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2520
         Width           =   495
      End
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "cyb011.frx":140E
         Left            =   1440
         List            =   "cyb011.frx":1418
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "c_cuenta"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.ComboBox c_caja 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox t_importe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox t_destino 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_numint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Concepto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Importe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800080&
         Caption         =   "Num. Interno:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4950
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
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
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_movcaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub graba()
   Set rs1 = New ADODB.Recordset
   q = "select * from cyb_01 where [id_forma_pago] = " & c_caja.ItemData(c_caja.ListIndex)
   rs1.Open q, cn1
   If Not rs1.BOF And Not rs1.EOF Then
      cta = rs1("id_cuenta_cont")
   Else
      cta = 0
   End If
   Set rs1 = Nothing
   
   Set rs = New ADODB.Recordset
   Set rs2 = New ADODB.Recordset
   If t_numint <> "" Then
     q = "SELECT * FROM CYB_05 where [num_mov_caja] = " & Val(t_numint)
     rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
     If Not rs2.EOF And Not rs2.BOF Then
       If rs2("modulo") <> "J" Then
         k = MsgBox("El movimiento que desea modificar fue generado por : " & t_op & ". Confirma modificar solo la caja eliminando la asociacion entre ambos movimientos.", 4)
         If k = 6 Then
            q = "SELECT * FROM CYB_05 "
            rs.Open q, cn1, adOpenDynamic, adLockOptimistic
            rs.AddNew
            rs2.Delete
            
                
         Else
           Exit Sub
         End If
       Else
            q = "SELECT * FROM CYB_05 "
            rs.Open q, cn1, adOpenDynamic, adLockOptimistic
            rs.AddNew
                 
            rs2.Delete
     
            Call borracontabilidad(Val(t_numint), "J")
            
       End If
     
       'borra asiento
     
     Else
       Exit Sub
     End If
    Else
     q = "SELECT * FROM CYB_05 "
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     rs.AddNew
   End If
   Set rs2 = Nothing
   
   
   If Mid$(c_tipo, 1, 1) = "I" Then
      u = "D"
   Else
      u = "H"
   End If
   rs("ID_forma_pago") = c_caja.ItemData(c_caja.ListIndex)
   rs("id_cuenta_caja") = cta
   rs("id_cuenta_contra") = c_cuenta.ItemData(c_cuenta.ListIndex)
   rs("Descripcion") = t_destino
   rs("Importe") = Val(t_importe)
   rs("ubicacion") = u
   rs("fecha") = t_fecha
   rs("num_mov_int") = rs("num_mov_caja")
   rs("modulo") = "J"
   rs("Operacion") = "Mov.Caja " & Format$(rs("num_mov_caja"), "00000000")
   rs("id_usuario") = para.id_usuario
   numint = rs("num_mov_caja")
   t_numint = numint
   rs.Update
   Set rs = Nothing
   
   'sacar numero inetrno
   
   
   'graba asiento

  If Generaasientosauto Then
    
    numintcgr = saca_ultnumero_int_comp("G")
    Set rs = New ADODB.Recordset
    q = "select * from cyb_01 where [id_forma_pago] = " & c_caja.ItemData(c_caja.ListIndex)
    rs.MaxRecords = 1
    rs.Open q, cn1
    If Not rs.EOF And Not rs.BOF Then
         cta = rs("id_cuenta_cont")
    Else
         cta = para.cuenta_caja
    End If
    Set rs = Nothing
         
    If c_tipo.ListIndex = 0 Then 'ingreso
          u2 = "D"
          u1 = "H"
    Else
          u2 = "H"
          u1 = "D"
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
    cn1.BeginTrans
    QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
    QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Caja] " & c_tipo & Format$(numint, "00000000") & "', 'J', " & numint & ", " & Val(t_importe) & ", " & Val(t_importe) & ", " & para.id_usuario & ", '" & Left$(RTrim$(t_destino), 50) & "')"
    cn1.Execute QUERY
      
   'cuenta madre caja
    QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
    QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u2 & "', " & Val(t_importe) & ", '" & c_tipo & " " & Format$(numint, "00000000") & "')"
    cn1.Execute QUERY
         
    ic = 2
    cta = c_cuenta.ItemData(c_cuenta.ListIndex)
    QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
    QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Val(t_importe) & ", '" & c_tipo & " " & Format$(numint, "00000000") & "')"
    cn1.Execute QUERY
    
    cn1.CommitTrans
  End If
   
End Sub


Private Sub btnacepta_Click()
J = MsgBox("Confirma Grabar movimiento", 4)
If J = 6 Then
 If estadocaja(t_fecha) = "A" Then
   If verificaperiodog(t_fecha) = "A" Then
        Call graba
        J = MsgBox("Imprime Vale de Caja", 4)
        If J = 6 Then
          Call imprime
        End If
        Call limpia
        c_caja.SetFocus
        Me.Hide
    Else
        MsgBox ("Periodo cerrado. Imposible grabar operacion")
    End If
   Else
     MsgBox ("Caja CERRADA. Imposible realizar operacion")
 End If
End If
End Sub

Sub imprime()
k = InputBox$("Inrese cantidad de Copias", "Impresion de Vales de Caja", 1)
If Val(k) > 0 And Val(k) <= 4 Then
 For i = 1 To Val(k)
  Call imprimeempresa(14)
  Printer.FontName = "Courier New"
  Printer.Print
  Printer.Print Tab(50); "Vale de Caja Nro.:" & Format$(Val(t_numint), "0000000")
  Printer.Print
  Printer.FontBold = True
  Printer.Print "Fecha........: ";
  Printer.FontBold = False
  Printer.Print t_fecha
  Printer.Print
  Printer.FontBold = True
  Printer.Print "Destino......: ";
  Printer.FontBold = False
  Printer.Print t_destino
  Printer.Print
  Printer.FontBold = True
  Printer.Print "Tipo.........: ";
  Printer.FontBold = False
  Printer.Print c_tipo
  Printer.Print
  Printer.FontBold = True
  Printer.Print "Imputacion...: ";
  Printer.FontBold = False
  Printer.Print c_cuenta; "   -"; c_cuenta.ItemData(c_cuenta.ListIndex); " -"
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print Tab(40); "***************************"
  Printer.Print Tab(45); "$ "; asteriscos(Format$(Val(t_importe), "######0.00"), 15)
  Printer.Print Tab(40); "***************************"
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print "_________________________________________________________________________________"
  Printer.Print "Fecha Imp." & Now & "     Emitido por: " & glo.usuario
  If i < Val(k) Then
   Printer.NewPage
  End If
Next i
  Printer.EndDoc
End If
End Sub
Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub c_caja_LostFocus()
If c_caja.ListIndex < 0 Then
  c_caja.ListIndex = 0
End If

End Sub

Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Sub limpia()
t_fecha = ""
t_fechadif = ""
t_destino = ""
t_importe = ""
t_numint = ""
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

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If

End Sub

Private Sub Command1_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Command2_Click()

  Call imprime
End Sub

Private Sub Form_Activate()
Call barraesag(Me)
If t_numint <> "" Then
  Command2.Visible = True
Else
  Command2.Visible = False
End If
If para.cuenta_sel > 0 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
End If
End Sub
Sub cargacuentas()
  If c_tipo.ListIndex = 0 Then
    Call carga_cuentas_cont(c_cuenta, "C", "D", "I")
  Else
    Call carga_cuentas_cont(c_cuenta, "C", "D", "E")
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then

 Call tabup(Me)
End If


If KeyCode = vbKeyF9 Then
J = MsgBox("Confirma Grabar movimiento", 4)
If J = 6 Then
 If estadocaja(t_fecha) = "A" Then
  Call graba
  Call limpia
  c_caja.SetFocus
   Me.Hide
 Else
  MsgBox ("Caja CERRADA. Imposible realizar operacion")
 End If
End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   
End If

If KeyAscii = 13 Then
  Call TabEnter2(Me, 5)
End If
End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_formas_pago(c_caja, "O")
Call carga_cuentas_cont(c_cuenta, "C", "D")
para.cuenta_sel = 0
End Sub

Private Sub t_destino_LostFocus()
If t_destino = "" Then
  t_destino = "*"
End If
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Or t_fecha = "" Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub


