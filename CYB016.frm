VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_VENTACH 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "VENTA DE CHEQUES"
   ClientHeight    =   8595
   ClientLeft      =   1725
   ClientTop       =   2160
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8595
   ScaleWidth      =   12000
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   15
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CYB016.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CYB016.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle de Valores "
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   11655
      Begin VB.TextBox t_ingresado 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   12
         Top             =   4080
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid msf2 
         Height          =   3735
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6588
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Total de Valores:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   13
         Top             =   4080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operacion"
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   11655
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox c_banco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Banco:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Detalle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   49
      TabIndex        =   4
      Top             =   6960
      Width           =   7335
   End
   Begin VB.TextBox t_numint 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   8340
      Width           =   12000
      _ExtentX        =   21167
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Observaciones:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Nro. Interno:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "cyb_VENTACH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim EXISTE As String
Dim phoy As Double
Dim crg As Integer
Dim pagomes As Double
Dim retmes As Double
Dim impnosujret As Double
Dim rg_alicuota As Double
Dim excedente As Double
Dim gnumintop As Long








Private Sub btnacepta_Click()
If estadocaja(t_fecha) = "A" Then
 J = MsgBox("Confirma emision de deposito", 4)
 If J = 6 Then
  If verificaperiodog(t_fecha) = "A" Then
  Call graba
  Call pi3
 Else
  MsgBox ("Periodo cerrado. Imposible grabar operacion")
 End If
 End If
Else
 MsgBox ("Caja Cerrada. Imposible realizar operacion")
End If
End Sub

Private Sub btnsale_Click()
inicio_bancos.Show
Unload Me
End Sub

Private Sub c_banco_GotFocus()
Call pi3
End Sub

Private Sub c_banco_LostFocus()
If c_banco.ListIndex < 0 Then
  c_banco.ListIndex = 0
End If
End Sub

Private Sub detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Detalle = Detalle & " "
   If Val(t_ingresado) > 0 Then
     btnacepta.Enabled = True
     btnacepta.SetFocus
   End If
End If



End Sub




Sub armagrid2()
msf2.clear
msf2.Rows = 1
msf2.Cols = 10
msf2.ColWidth(0) = 600
msf2.ColWidth(1) = 1200
msf2.ColWidth(2) = 1200
msf2.ColWidth(3) = 2500
msf2.ColWidth(4) = 1700
msf2.ColWidth(5) = 1700
msf2.ColWidth(6) = 1000
msf2.ColWidth(7) = 1000
msf2.ColWidth(8) = 1000
msf2.ColWidth(9) = 1000

msf2.TextMatrix(0, 0) = "Cod."
msf2.TextMatrix(0, 1) = "Forma Pago"
msf2.TextMatrix(0, 2) = "Num.Cheque"
msf2.TextMatrix(0, 3) = "Detalle/Banco"
msf2.TextMatrix(0, 4) = "Sucursal"
msf2.TextMatrix(0, 5) = "Titular"
msf2.TextMatrix(0, 6) = "Importe"
msf2.TextMatrix(0, 7) = "Fecha Dif."
msf2.TextMatrix(0, 8) = "Num.Int."
msf2.TextMatrix(0, 9) = "Cuenta"


End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  'Case Is = 27
  '      Me.Hide
End Select
End Sub

Sub graba()
            
         Set rs = New ADODB.Recordset
         q = "select * from cyb_01 where [id_forma_pago] = " & c_banco.ItemData(c_banco.ListIndex)
         rs.Open q, cn1
         ctabanco = rs("id_CUENTA_CONT")
         Set rs = Nothing

      cn1.BeginTrans
      
      Set rs = New ADODB.Recordset
      q = "select * from cyb_04"
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      rs.AddNew
        rs("id_banco") = c_banco.ItemData(c_banco.ListIndex)
        rs("fecha") = t_fecha
        rs("importe") = Val(t_ingresado)
        rs("id_tipomov") = 40
        rs("fecha_dif") = t_fecha
        rs("ubicacion") = "H"
        rs("entro") = "N"
        rs("fecha_acreed") = t_fecha
        rs("num_comp") = Val(t_numcomp)
        rs("detalle") = Detalle & " "
        rs("Modulo") = "B"
        rs("num_mov_int") = rs("num_mov_banco")
        rs("id_tipodbcr") = 1
        rs("num_mov_int_compras") = 0
      rs.Update
      
      numintb = rs("num_mov_banco")
      Set rs = Nothing
      
      
      'actualiza pagos
      
      For i = 1 To msf2.Rows - 1
       If Val(msf2.TextMatrix(i, 0)) = 3 Then 'ch. terceros
        Set rs = New ADODB.Recordset
        q = "select * from cyb_03 where [num_interno] = " & Val(msf2.TextMatrix(i, 8))
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          rs("estado") = "V"
          rs("destino") = "Vendido" & c_banco
          rs("num_int_op") = 0
          rs("fecha_salida") = t_fecha
          rs("tipo_salida") = "V"
          rs("NUM_MOV_BANCO_E") = numintb
          numintch = Val(msf2.TextMatrix(i, 8))
          rs.Update
         Else
          numintch = 0
        End If
        Set rs = Nothing
       Else
        numintch = 0
       End If
      
      
      If Val(msf2.TextMatrix(i, 0)) >= 50 Then 'ch. propios
         
        Set rs = New ADODB.Recordset
        q = "select * from cyb_02 where [id_banco] = " & Val(msf2.TextMatrix(i, 0)) & " and [num_cheque] = " & Val(msf2.TextMatrix(i, 2))
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          If rs("estado") = "P" Then
             rs("estado") = "V"
             rs("fecha_dif") = msf2.TextMatrix(i, 7)
             rs("fecha_emision") = t_fecha
             rs("destino") = "Vendido" & c_banco
             rs("importe") = Val(msf2.TextMatrix(i, 6))
             rs("num_int_op") = 0
             rs("num_mov_banco") = numintb
             rs.Update
             
          Else
             MsgBox ("Error al asignar ch. propio")
          End If
        End If
        Set rs = Nothing
       
        
        'emito ch.
        QUERY = "INSERT INTO cyb_04([id_banco], [fecha], [importe], [id_tipomov], [fecha_dif], [ubicacion], [entro], [fecha_acreed], [num_comp], [detalle], [modulo], [num_mov_int], [id_tipodbcr])"
        QUERY = QUERY & " VALUES (" & Val(msf2.TextMatrix(i, 0)) & ", '" & t_fecha & "', " & Val(msf2.TextMatrix(i, 6)) & ", 1, '" & msf2.TextMatrix(i, 7) & "', 'D', 'N', '" & t_fecha & "', " & Val(t_numop) & ", 'Ch." & Format$(Val(msf2.TextMatrix(i, 2)), "0000000000") & "', 'B', " & numintb & ", 1)"
        cn1.Execute QUERY
        
       
       End If
 
      
      
      'para cada item si mueve caja
      Set rs = New ADODB.Recordset
      q = "select * from cyb_01 where [id_forma_pago] = " & Val(msf2.TextMatrix(i, 0))
      rs.Open q, cn1
      If Not rs.BOF And Not rs.EOF Then
        If rs("caja") = "S" Then
          'grabo mov caja
           
           If Val(msf2.TextMatrix(i, 0)) = 3 Then 'ch. terc.
             nicht = Val(msf2.TextMatrix(i, 8))
           Else
             nicht = 0
           End If
           
           QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
           QUERY = QUERY & " VALUES (" & Val(msf2.TextMatrix(i, 9)) & ", " & ctabanco & ", '" & Left$(Detalle & " Nro." & t_numcomp, 49) & " ', " & Val(msf2.TextMatrix(i, 6)) & ", 'H', '" & t_fecha & "', " & numintb & ", 'B', 'Vta.Ch. " & Format$(numintb, "00000000") & "', " & Val(msf2.TextMatrix(i, 0)) & ", " & nicht & ", " & para.id_usuario & ")"
           cn1.Execute QUERY
        End If
      End If
      Set rs = Nothing
      
      Next i
     
      
      
      
'contabilidad
If Generaasientosauto Then
      
      Set rs = New ADODB.Recordset
      q = "select * from cyb_06 where [id_tipomov] = 40"
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        contab = rs("contabilidad")
      Else
        contab = "N"
      End If
      Set rs = Nothing
      
      If contab <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")
         u1 = contab
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
          
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Vta.Ch.] N.I." & Format$(numintb, "00000000") & "', 'B', " & numintb & ", " & Val(t_ingresado) & ", " & Val(t_ingresado) & ", " & para.id_usuario & ", '" & Left$(Detalle & " Vta.Ch." & t_numcomp, 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre bancos
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & ctabanco & ", '" & u1 & "', " & Val(t_ingresado) & ", 'Vta.Ch." & Format$(numintb, "00000000") & "')"
         cn1.Execute QUERY
      
         'formas de pago
         ic = 2
         For i = 1 To msf2.Rows - 1
            Set rs = New ADODB.Recordset
            q = "select * from cyb_01 where [id_forma_pago] = " & Val(msf2.TextMatrix(i, 0))
            rs.Open q, cn1
            If Not rs.BOF And Not rs.EOF Then
              d = Left$(RTrim$(msf2.TextMatrix(i, 3)), 35) & " " & msf2.TextMatrix(i, 2)
              cta = rs("id_cuenta_cont")
              QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
              QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u2 & "', " & Val(msf2.TextMatrix(i, 6)) & ", '" & d & "')"
              cn1.Execute QUERY
              ic = ic + 1
            End If
         Next i
      
      End If
      
  End If
      
      
 cn1.CommitTrans
        

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos. Proc:Venta Ch.  Funcion:Graba")
  


End Sub

Private Sub Form_Load()
c_banco.clear
Call carga_formas_pago(c_banco, "B")
Call armagrid2


End Sub

Private Sub msf2_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F1] Ch.Terc. - [F2] Ch.Propios - [F9] Continua - [F5] Elimina Renglon "
If msf2.Rows > 0 Then
  msf2.FocusRect = flexFocusNone
Else
  msf2.FocusRect = flexFocusLight
End If
t_ingresado = Format$(suma_msflexgrid(msf2, 6), "######0.00")


End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
   Load op_fp2
   op_fp2.t_modulo = "V"
  op_fp2.Show
End If

If KeyCode = vbKeyF1 Then
  Load op_fp1
  op_fp1.t_modulo = "V"
  op_fp1.Show
End If


If KeyCode = vbKeyF9 Then
    Detalle.SetFocus
End If


If KeyCode = vbKeyF5 Then
 If msf2.Rows > 2 Then
    msf2.RemoveItem (msf2.Row)
 Else
   Call armagrid2
 End If
End If

End Sub

Private Sub msf2_LostFocus()
Call barra(Me)
t_ingresado = suma_msflexgrid(msf2, 6)
msf2.FocusRect = flexFocusLight
End Sub




Private Sub pi3()
   Call INICIALIZA2(Me)
   Call armagrid2
   btnacepta.Enabled = False
   c_banco.SetFocus
End Sub


Sub totales()
k = 1
t = 0
While k <= msf1.Rows - 1
   t = t + Val(msf1.TextMatrix(k, 2))
   k = k + 1
Wend
t_pago = Format$(t, "######0.00")
t_total = Format$(Val(t_pago) - Val(retencion) - Val(t_retib), "######0.00")
t_totald = Format$(Val(t_total) / Val(fdolar), "######0.00")
t_diferencia = Format$(Val(t_total) - Val(t_ingresado), "######0.00")
End Sub

Sub totales2()
t_pago = Format$(t_pago, "######0.00")
t_total = Format$(Val(t_pago) - Val(retencion) - Val(t_retib), "######0.00")
't_totald = Format$(Val(t_total) / Val(fdolar), "######0.00")
t_diferencia = Format$(Val(t_total) - Val(t_ingresado), "######0.00")

End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
      t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub
