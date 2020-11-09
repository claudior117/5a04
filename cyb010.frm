VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_chpropios2 
   Caption         =   "Emitir Cheques Propios"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Movimiento"
      Height          =   2775
      Left            =   360
      TabIndex        =   21
      Top             =   2040
      Width           =   8295
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Text            =   "c_cuenta"
         Top             =   1680
         Width           =   6015
      End
      Begin VB.TextBox t_op 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox t_importe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox t_destino 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox t_fechadif 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_numint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Imputacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Tipo Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Importe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Destino:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Diferida:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Emision:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800080&
         Caption         =   "Num. Interno:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   4920
      Width           =   6615
      Begin VB.CommandButton Command5 
         Caption         =   "Anula Ch."
         Height          =   375
         Left            =   4080
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Borra Ch."
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Borra Mov."
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Emitir Ch."
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cheque"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_banco 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   20
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox t_idbanco 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_estado 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   17
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox t_chequera 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox t_ch 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
         Caption         =   "Estado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Nro. Chequera:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Nro. Cheque"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Bancaria:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   5805
      Width           =   8895
      _ExtentX        =   15690
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
            TextSave        =   "20/03/2014"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "18:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_chpropios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub graba()
   Set rs = New ADODB.Recordset
   q = "SELECT * FROM CYB_04"
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   rs.AddNew
   rs("ID_BANCO") = Val(t_idbanco)
   rs("FECHA") = t_fecha
   rs("IMPORTE") = Val(t_importe)
   rs("ID_TIPOMOV") = 1 'EMISION
   rs("FECHA_DIF") = t_fechadif
   rs("UBICACION") = "D"
   rs("ENTRO") = "N"
   rs("FECHA_ACREED") = t_fecha
   rs("NUM_COMP") = Val(t_ch)
   rs("DETALLE") = Left$(t_destino, 39) & " "
   rs("Modulo") = "B"
   rs("num_mov_int") = rs("num_mov_banco")
   rs("id_tipodbcr") = 1
   rs("num_mov_int_compras") = 0
   rs.Update
   
   nmi = rs("num_mov_banco")
   Set rs = Nothing
   
   
   Set rs = New ADODB.Recordset
   q = "SELECT * FROM CYB_02 where [id_banco] = " & Val(t_idbanco) & " and [num_cheque] = " & Val(t_ch)
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs.BOF And Not rs.EOF Then
     rs("fecha_emision") = t_fecha
     rs("fecha_dif") = t_fechadif
     rs("estado") = t_op
     rs("destino") = t_destino
     rs("importe") = Val(t_importe)
     rs("num_mov_banco") = nmi
     rs.Update
   End If
   Set rs = Nothing
   

   If Generaasientosauto Then
    If c_cuenta.ListIndex > 0 Then
      numintcgr = saca_ultnumero_int_comp("G")
      q = "select * from cyb_01 where [id_forma_pago] = " & Val(t_idbanco)
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
         cta = rs("id_cuenta_ch_dif")
         u1 = "H"
         u2 = "D"
         
         Set rs1 = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs1 = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Bancos] Emision Ch.Nro." & Format$(Val(t_ch), "00000000") & "', 'B', " & nmi & ", " & Val(t_importe) & ", " & Val(t_importe) & ", " & para.id_usuario & ", '" & Left$(RTrim$(t_destino), 50) & "')"
         cn1.Execute QUERY
      
         
         'cuenta madre banco
         ic = 1
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Val(t_importe) & ", '" & dcta & "')"
         
         cn1.Execute QUERY
         ic = ic + 1
      
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_importe) & ", '" & "Emision Ch. Propio" & "')"
         cn1.Execute QUERY
      
      End If
    End If
   End If
   
   If t_op = "J" Then
         QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
         QUERY = QUERY & " VALUES (" & para.cuenta_caja & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & t_destino & "', " & Val(t_importe) & ", 'D', '" & Format$(t_fecha, "DD/MM/YYYY") & "', " & nmi & ", 'B', 'Cobro ChP." & Format$(t_ch, "0000000000") & "' ,1, 0, " & para.id_usuario & ")"
         cn1.Execute QUERY
   End If

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

Private Sub Command1_Click()
If t_estado = "P" Or t_estado = "A" Or t_estado = "D" Then
   Call limpia
   Frame3.Enabled = True
   t_fecha.SetFocus
Else
   MsgBox ("El cheque ya fue Emitido")
End If
End Sub
Sub limpia()
t_fecha = ""
t_fechadif = ""
t_destino = ""
t_importe = ""
t_op = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
J = MsgBox("Confirma Eliminar Movimiento Bancario", 4)
If J = 6 Then
   Call borramovbanco
   MsgBox ("Movimiento Eliminado")
End If
  
End Sub
Sub borramovbanco()
 If Val(t_numint) > 0 Then
  Set rs = New ADODB.Recordset
  q = "select * from cyb_04 where [modulo] = 'B' and [num_mov_int] = " & Val(t_numint)
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs.EOF And Not rs.BOF Then
    rs.Delete
    rs.Update
  End If
  Set rs = Nothing
  
  
  Set rs = New ADODB.Recordset
  q = "select * from cyb_05 where [modulo] = 'B' and [num_mov_int] = " & Val(t_numint)
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs.EOF And Not rs.BOF Then
    rs.Delete
    rs.Update
  End If
  Set rs = Nothing
         
  Set rs = New ADODB.Recordset
  q = "select * from c_02 where [modulo] = 'B' and [num_mov_int] = " & Val(t_numint)
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs.EOF And Not rs.BOF Then
    Set rs2 = New ADODB.Recordset
    q = "select * from c_03 where [num_interno] = " & rs("num_interno")
    rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
    While Not rs2.EOF
      rs2.Delete
      rs2.MoveNext
    Wend
    Set rs2 = Nothing
    rs.Delete
    rs.Update
  End If
  Set rs = Nothing
 
End If
  
  
  
   Set rs = New ADODB.Recordset
   q = "SELECT * FROM CYB_02 where [id_banco] = " & Val(t_idbanco) & " and [num_cheque] = " & Val(t_ch)
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs.BOF And Not rs.EOF Then
     rs("estado") = "P"
     rs("destino") = " "
     rs("importe") = 0
     rs("num_mov_banco") = 0
     rs.Update
   End If
   Set rs = Nothing

 
End Sub
Private Sub Command4_Click()
J = MsgBox("Confirma Eliminar ch. nro " & t_ch, 4)
If J = 6 Then
  Set rs = New ADODB.Recordset
  q = "SELECT * FROM CYB_02 where [id_banco] = " & Val(t_idbanco) & " and [num_cheque] = " & Val(t_ch)
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs.BOF And Not rs.EOF Then
     If rs("estado") = "P" Then
        rs.Delete
        rs.Update
        MsgBox ("Cheque Eliminado!!!!")
     Else
        MsgBox ("¡ERROR!: El cheque tiene un movimiento asociado, elimine primero el movimiento y luego el cheque. Operacion NO REALIZADA!!!!")
     End If
  Else
   MsgBox ("El Cheque no Existe")
  End If
  Set rs = Nothing
End If
End Sub

Private Sub Command5_Click()
J = MsgBox("Confirma Anular ch. nro. " & t_ch, 4)
If J = 6 Then
  Set rs = New ADODB.Recordset
  q = "SELECT * FROM CYB_02 where [id_banco] = " & Val(t_idbanco) & " and [num_cheque] = " & Val(t_ch)
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs.BOF And Not rs.EOF Then
     If rs("estado") = "P" Then
        rs("estado") = "A"
        rs.Update
        MsgBox ("Cheque Anulado!!!!")
     Else
        MsgBox ("¡ERROR!: El cheque tiene un movimiento asociado, elimine primero el movimiento y luego anule el cheque. Operacion NO REALIZADA!!!!")
     End If
  Else
   MsgBox ("El Cheque no Existe")
  End If
  Set rs = Nothing
End If

End Sub

Private Sub Form_Activate()
Frame1.Enabled = False
Frame2.Enabled = True
Frame3.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then

 Call tabup(Me)
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Frame2.Enabled = True
   Frame3.Enabled = False
End If

If KeyAscii = 13 Then
  Call TabEnter2(Me, 5)
End If
End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "Sin Imputacion", 0
c_cuenta.ListIndex = 0
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

Private Sub t_fechadif_LostFocus()
If Not IsDate(t_fechadif) Or t_fechadif = "" Then
  t_fechadif = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub

Private Sub t_op_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[E]Entrega a prov. - [J] Cobrado(a Caja) "

End Sub

Private Sub t_op_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If t_op <> "" Then
  t_op = UCase$(t_op)
  Select Case t_op
   Case Is = "E", Is = "J"
      J = MsgBox("Graba Operacion", 4)
      If J = 6 Then
        If verificaperiodog(t_fecha) = "A" Then
          Call graba
        Else
          MsgBox ("Periodo Cerrado. Imposible grabar operacion")
        End If
          Call limpia
          Unload Me
      End If
   End Select
 End If
End If
End Sub

