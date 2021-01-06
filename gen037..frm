VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_citicom 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CITI COMPRAS UNIFICADO"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Salida:"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Width           =   9975
      Begin VB.TextBox t_carpeta 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   8055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Carpeta destino:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   7920
      Width           =   9975
      Begin VB.Label Label6 
         Caption         =   "Genera 2 archivos c:\citicom.txt con las operaciones a exportar y c:\alucuo.txt con las alicuotas"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   615
      Left            =   7320
      TabIndex        =   14
      Top             =   720
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Verifica Totales"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   67567617
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen037..frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen037..frx":0882
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   2
      Top             =   8550
      Width           =   12060
      _ExtentX        =   21273
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
            TextSave        =   "23/11/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:07 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5535
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9763
      _Version        =   393216
   End
End
Attribute VB_Name = "gen_citicom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim c5 As Double

Sub carga()
 'Dim cr(2) As Long
 Dim nic As String
 espere.Show
 espere.Label1 = "Espere...... Generando Listado para CITI"
 espere.Refresh
 Call armagrid
  q = "select * from a5, a1, g2 where [grabado] <> 'N' and  [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  q = q & " order by [fecha], [letra], [num_comprobante]"
  Set rs = New adodb.Recordset
  
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  While Not rs.EOF
     er = ""
     obserr = ""
     F = Mid$(Format$(rs("fecha"), "dd/mm/yyyy"), 7, 4) & Mid$(Format$(rs("fecha"), "dd/mm/yyyy"), 4, 2) & Mid$(Format$(rs("fecha"), "dd/mm/yyyy"), 1, 2)
     Select Case rs("letra")
     Case Is = "A"
          tc = Format$(rs("cod_afip_a"), "000")
          cic = "80"  'cuit
          nic = Format$(rs("cuit05"), "00000000000000000000")
          tio = "0"
          If Val(nic) < 20000000000# Then
            er = er & "*"
            obserr = obserr & "Cuit "
          End If
     Case Is = "B"
          tc = Format$(rs("cod_afip_b"), "000")
          tio = "0"
          nic = Format$(rs("cuit05"), "00000000000000000000")
          
     Case Is = "C"
          tc = Format$(rs("cod_afip_c"), "000")
          cic = "80"  'cuit
          nic = Format$(rs("cuit05"), "00000000000000000000")
          tio = "0"
     Case Else
          tc = Format$(rs("cod_afip_a"), "000")
          er = er & "*"
          obserr = obserr & "Letra Comp.- "
          tio = "0"
     End Select
          
     
     PtV = Format$(rs("sucursal"), "00000")
     nc = Format$(rs("num_comprobante"), "00000000000000000000")
     If Val(PtV) <= 0 Or Val(nc) <= 0 Then
         er = er & "*"
         obserr = obserr & "PV/Nro.- "
     End If
     
     If rs("moneda") = "P" Then
       c5 = 1
       moneda = "PES"
     Else
       c5 = rs("cotiz_dolar")
       moneda = "DOL"
     End If
     cambio2 = Format$(c5, "###0.0000")
         
     t = Format$(rs("total"), "######0.00")
     i = Format$(rs("a5.iva"), "######0.00")
     s = Format$(rs("subtotal"), "######0.00")
     ng = Format$((rs("no_grabado")), "######0.00")
     
     'definicion alicuita
     If Val(i) > 0 And Val(s) > 0 Then
        a = Val(i) / Val(s)
        e = 0
        ct = 1
        
        If a > 2 And a < 3 Then
           ti = 2.5
           e = 1
        End If
        
       If a > 4 And a < 6 Then
           ti = 5
           e = 1
        End If
       
        
        If a > 9 And a < 11 Then
           ti = 10.5
           e = 1
        End If
        
        If a > 19 And a < 22 Then
           ti = 21
           e = 1
           
        End If
        
        If a > 26 And a < 28 Then
           ti = 27
           e = 1
        End If
        
        
        
        If e = 0 Then
          ti = 21
        End If
        
     Else
       If tc < 11 Then
         ct = 1
         ti = 0
       Else
         ct = 0
         ti = 0
       End If
       
     End If
     
     
     'busco percpciones
     Set rs2 = New adodb.Recordset
     q = "select * from A13 where [num_int] = " & rs("num_int")
     rs2.Open q, cn1
     piv = 0
     pin = 0
     pip = 0
     While Not rs2.EOF
         Select Case rs2("id_percepcion")
         Case Is = 1 'perc iva
           piv = piv + rs2("importe")
         Case Is = 2 'perc ib
           pip = pip + rs2("importe")
         Case Else   'perc gan
           pin = pin + rs2("importe")
        End Select
        rs2.MoveNext
     Wend
     Set rs2 = Nothing
     
     msf1.AddItem er & Chr$(9) & F & Chr(9) & rs("proveedor05") & Chr(9) & cic & Chr$(9) & nic & Chr$(9) & tc & Chr$(9) & PtV & Chr$(9) & nc & Chr(9) & t & Chr(9) & ng & Chr$(9) & s & Chr(9) & ti & Chr(9) & i & Chr$(9) & ct & Chr(9) & rs("num_int") & Chr(9) & obserr & Chr(9) & moneda & Chr(9) & cambio2 & Chr(9) & tio & Chr(9) & pin & Chr(9) & pip & Chr(9) & piv
       
    rs.MoveNext
  Wend
  Unload espere
     
End Sub
Private Sub btnacepta_Click()
  Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub







Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1.Value
Else
  t_fecha2 = cal1.Value
End If
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub



Private Sub Command1_Click()
If t_fecha = "" Or t_fecha2 = "" Then
  MsgBox ("Debe indicar un periodo de trabajo para realizar esta operacion")
  Exit Sub
End If
h = MsgBox("Verificacion de Totales. Asegurese de haber indicado correctamente el periodo de trabajo y No apague la maquina ni cancele este proceso. ¿Esta seguro que quiere actualizar? ", 4)
If h = 6 Then
espere.Show
espere.Refresh
qm = "select * from vta_02 where  [grabado] <> 'N'"
c = " and "
If IsDate(t_fecha) Then
    qm = qm & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
End If
  
If IsDate(t_fecha2) Then
   qm = qm & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
End If
  
If c_sucursal.ListIndex > 0 Then
    qm = qm & c & " and [sucursal_ingreso] = " & Val(c_sucursal)
End If
Set rs2 = New adodb.Recordset
rs2.Open qm, cn1
a = 1
While Not rs2.EOF
 Call verifica_tasa_iva(rs2("num_int"))
 
 rs2.MoveNext
Wend
Set rs2 = Nothing
Unload espere
MsgBox ("Proceso Terminado")
End If

End Sub

Private Sub Command2_Click()
Load gen_seleccionacarpeta
gen_seleccionacarpeta.t_llamada = "2"
gen_seleccionacarpeta.Show

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 22
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 900
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 1800
msf1.ColWidth(5) = 500
msf1.ColWidth(6) = 800
msf1.ColWidth(7) = 1800
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1100
msf1.ColWidth(10) = 1000
msf1.ColWidth(11) = 1100
msf1.ColWidth(12) = 900
msf1.ColWidth(13) = 1100
msf1.ColWidth(14) = 800
msf1.ColWidth(15) = 3000
msf1.ColWidth(16) = 500
msf1.ColWidth(17) = 1000
msf1.ColWidth(18) = 500
msf1.ColWidth(19) = 1000
msf1.ColWidth(20) = 1000
msf1.ColWidth(21) = 1000
msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Proveedor"
msf1.TextMatrix(0, 3) = "Tipo Doc."
msf1.TextMatrix(0, 4) = "Nro.Cuit/Dni "
msf1.TextMatrix(0, 5) = "Tipo Comp."
msf1.TextMatrix(0, 6) = "PV"
msf1.TextMatrix(0, 7) = "Numero "
msf1.TextMatrix(0, 8) = "Total"
msf1.TextMatrix(0, 9) = "No Grav."
msf1.TextMatrix(0, 10) = "Gravado"
msf1.TextMatrix(0, 11) = "% iva"
msf1.TextMatrix(0, 12) = "Iva"
msf1.TextMatrix(0, 13) = "Cant.tasas "
msf1.TextMatrix(0, 14) = "nro Int."
msf1.TextMatrix(0, 15) = "Errores"
msf1.TextMatrix(0, 16) = "Moneda"
msf1.TextMatrix(0, 17) = "Cambio"
msf1.TextMatrix(0, 18) = "T.OP."
msf1.TextMatrix(0, 19) = "Perc I.Nac"
msf1.TextMatrix(0, 20) = "Perc I.Prov"
msf1.TextMatrix(0, 21) = "Perc Iva"

For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 14
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0
t_carpeta = "C:\"
Call barraesag(Me)
cal1.Visible = False
Call armagrid
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F2] Cambia DNI/Cuit - [F5] Exporta -  [F7] Imprime - [F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    c(6) = 6
    c(7) = 7
    c(8) = 8

    For i = 9 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE IVA VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF2 Then
  t = InputBox$("CITI ventas", "Cambia valor celda", msf1.TextMatrix(msf1.Row, msf1.col))
  If t <> "" Then
    t = Format$(Val(t), "00000000000")
    msf1.TextMatrix(msf1.Row, msf1.col) = t
    QUERY = "update vta_02 set  [cuit02]='" & t & "'"
    QUERY = QUERY & " where [num_int]= " & Val(msf1.TextMatrix(msf1.Row, 14))
    cn1.BeginTrans
    cn1.Execute QUERY
    cn1.CommitTrans
  End If
End If


If KeyCode = vbKeyF5 Then
  J = MsgBox("Confirma genera archivos para importar del aplicativo Citi Compras. Carpeta destino " & t_carpeta, 4)
  If J = 6 Then
    If t_carpeta <> "" Then
      Call exporta
    Else
      MsgBox ("Debe seleccionar una carpeta destino")
   End If
  End If
  
End If




End Sub
Sub exporta()
Dim Detalle As String
k = 1
a1 = t_carpeta & "citicom.txt"
Open a1 For Output As #1
ni15 = "000000000000000"
ni16 = "                "

Detalle = String(75, " ")
While k <= msf1.Rows - 1
   t = Mid$(Format$(Val(msf1.TextMatrix(k, 8)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 8)), "0000000000000.00"), 15, 2)
   ng = Mid$(Format$(Val(msf1.TextMatrix(k, 9)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 9)), "0000000000000.00"), 15, 2)
   g = Mid$(Format$(Val(msf1.TextMatrix(k, 10)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 10)), "0000000000000.00"), 15, 2)
   a = Mid$(Format$(Val(msf1.TextMatrix(k, 11)), "00.00"), 1, 2) & Mid$(Format$(Val(msf1.TextMatrix(k, 11)), "00.00"), 4, 2)
   i = Mid$(Format$(Val(msf1.TextMatrix(k, 12)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 12)), "0000000000000.00"), 15, 2)
   e = "000000000000000"
   pnc = "000000000000000"
   pn = Mid$(Format$(Val(msf1.TextMatrix(k, 19)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 19)), "0000000000000.00"), 15, 2)
   pp = Mid$(Format$(Val(msf1.TextMatrix(k, 20)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 20)), "0000000000000.00"), 15, 2)
   pi = Mid$(Format$(Val(msf1.TextMatrix(k, 21)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 21)), "0000000000000.00"), 15, 2)
   ca = Format$(Val(msf1.TextMatrix(k, 13)), "0")
   nd = Format$(Val(msf1.TextMatrix(k, 4)), "00000000000000000000")
   tc = Mid$(Format$(Val(msf1.TextMatrix(k, 17)), "0000.000000"), 1, 4) & Mid$(Format$(Val(msf1.TextMatrix(k, 17)), "0000.000000"), 6, 6)
    
   l = msf1.TextMatrix(k, 1) & msf1.TextMatrix(k, 5) & msf1.TextMatrix(k, 6) & msf1.TextMatrix(k, 7) & ni16 & msf1.TextMatrix(k, 3) & nd & Left$(Format$(msf1.TextMatrix(k, 2), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30) & t & ng & e & pi & pn & pp & e & e & msf1.TextMatrix(k, 16) & tc & msf1.TextMatrix(k, 13) & msf1.TextMatrix(k, 18) & i & e & "00000000000" & "                              " & e
   Print #1, l
   k = k + 1
Wend
Close #1

k = 1
a1 = t_carpeta & "alicuo.txt"
Open a1 For Output As #1
ni15 = "000000000000000"
ni16 = "                "
While k <= msf1.Rows - 1
 If Val(msf1.TextMatrix(k, 13)) > 0 Then
   g = Mid$(Format$(Val(msf1.TextMatrix(k, 10)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 10)), "0000000000000.00"), 15, 2)
   a = Mid$(Format$(Val(msf1.TextMatrix(k, 11)), "00.00"), 1, 2) & Mid$(Format$(Val(msf1.TextMatrix(k, 11)), "00.00"), 4, 2)
   i = Mid$(Format$(Val(msf1.TextMatrix(k, 12)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 12)), "0000000000000.00"), 15, 2)
   e = "000000000000000"
   pnc = "000000000000000"
   nd = Format$(Val(msf1.TextMatrix(k, 4)), "00000000000000000000")
   l = msf1.TextMatrix(k, 5) & msf1.TextMatrix(k, 6) & msf1.TextMatrix(k, 7) & msf1.TextMatrix(k, 3) & nd & g & "0005" & i
   Print #1, l
 End If
   k = k + 1
Wend
Close #1


MsgBox ("Operacion Terminada. Ingrse al prgrama Citi Compras del siap e Importe los registros generados")

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 14)
    vta_cc_detalle.Show
  End If
End If

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"


End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yyyy")
  End If
End If

End Sub
