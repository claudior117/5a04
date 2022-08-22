VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_citi 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CITI  VENTAS"
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
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   7080
      Width           =   9975
      Begin VB.Label Label6 
         Caption         =   "El archivo de exportacion para importar del siap se guarda como c:\citivta.txt"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   8895
      End
      Begin VB.Label Label5 
         Caption         =   "Los comprobantes con mas de una tasa generan tantas filas como tasas tengan."
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   8895
      End
      Begin VB.Label Label4 
         Caption         =   "Los comprobantes B mayores a $1000.00 deben tener como identicacion del cliente el DNI ."
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   9375
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
         BackColor       =   &H00800080&
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
      StartOfWeek     =   173342721
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
         BackColor       =   &H00800080&
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
         BackColor       =   &H00800080&
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
         Picture         =   "VTA038.frx":0000
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
         Picture         =   "VTA038.frx":0882
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
            TextSave        =   "22/08/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:39 p.m."
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
Attribute VB_Name = "vta_citi"
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
  q = "select * from VTA_02, vta_01, vta_06, g3 where [grabado] <> 'N' and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente] and vta_01.[id_tipoiva] = [cod_tipoiva] and vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_sucursal.ListIndex > 0 Then
    q = q & c & " and [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
 q = q & c & "  (vta_02.[id_tipocomp] < 10 or vta_02.[id_tipocomp] > 200)"
  q = q & " order by [fecha], [letra], [num_comp]"
  Set rs = New ADODB.Recordset
  
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
          tc = Format$(rs("cod_afip_a"), "00")
          cic = "80"  'cuit
          nic = Format$(rs("cuit02"), "00000000000")
          If Val(nic) < 20000000000# Then
            er = er & "*"
            obserr = obserr & "Cuit "
          End If
     Case Is = "B"
          tc = Format$(rs("cod_afip_b"), "00")
          If rs("id_tipo_iva02") = 3 And rs("total") < 1000 Then
             cic = "99"
             nic = "00000000000"
          Else
            If rs("id_tipo_iva02") = 3 Then
               cic = "96"
               nic = Format$(rs("cuit02"), "00000000000")  'dni
                If Val(nic) <= 0 Then
                  er = er & "*"
                  obserr = obserr & "DNI "
                End If
            Else
               cic = "80"  'cuit
               nic = Format$(rs("cuit02"), "00000000000")
               If Val(nic) < 20000000000# Then
                er = er & "*"
                obserr = obserr & "Cuit "
               End If
            End If
          End If
     Case Is = "E"
          tc = Format$(rs("cod_afip_e"), "00")
          cic = "80"  'cuit
          nic = Format$(rs("cuit02"), "00000000000")
     Case Else
          tc = Format$(rs("cod_afip_a"), "00")
          er = er & "*"
          obserr = obserr & "Letra Comp.- "
     End Select
          
     'If Val(nic) <= 0 And Val(cic) <> 96 Then
     '  er = er & "*"
     '  obserr = obserr & "Cuit/DNI- "
     'End If
     
     If Val(nic) = Val(glo.CUIT) Then
       er = er & "*"
       obserr = obserr & "Cuit"
     
     End If
     PtV = Format$(rs("vta_02.sucursal"), "0000")
     nc = Format$(rs("num_comp"), "00000000000000000000")
     If Val(PtV) <= 0 Or Val(nc) <= 0 Then
         er = er & "*"
         obserr = obserr & "PV/Nro.- "
     End If
     
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     
     If rs("vta_02.id_tipocomp") <> 101 Then 'retencion de iva
         t = Format$(rs("total") * c5, "######0.00")
         i = Format$(rs("VTA_02.iva") * c5, "######0.00")
         s = Format$(rs("subtotal") * c5, "######0.00")
         ng = Format$((rs("impuestos") + rs("perc_ib") + rs("perc_gan")) * c5, "######0.00")
         rp = Format$(rs("perc_iva") * c5, "######0.00") 'ret/perc iva
     Else
        t = Format$(rs("total") * c5, "######0.00")
        i = Format$(0, "######0.00")
        s = Format$(0, "######0.00")
        ng = Format$(0, "######0.00")
        rp = Format$(rs("total") * c5, "######0.00") 'ret/perc iva
     End If
   
     q = "select * from vta_09 where [num_int] = " & rs("num_int")
     Set rs2 = New ADODB.Recordset
     rs2.Open q, cn1
     If Not rs2.EOF And Not rs2.BOF Then
           cr = rs2.GetRows
           r = UBound(cr, 2) + 1
           rs2.MoveFirst
     Else
         r = 1
     End If
     p = 1
     If er <> "" Then
       er = "ERR"
     End If
     
     If Val(cic) = 80 Then
       If verificacuit(nic) = 0 Then
        er = "ERR"
        obserr = obserr & " Nro. Cuit "
       End If
     End If
     
     While Not rs2.EOF
        n = Format$(rs2("neto") * c5, "#######0.00")
        iv = Format$(rs2("iva") * c5, "#######0.00")
        ti = Format$(rs2("tasa_iva"), "#######0.00")
        
        If p = r Then
           msf1.AddItem er & Chr$(9) & F & Chr(9) & rs("cliente02") & Chr(9) & cic & Chr$(9) & nic & Chr$(9) & tc & Chr$(9) & PtV & Chr$(9) & nc & Chr(9) & t & Chr(9) & "" & Chr$(9) & n & Chr(9) & ti & Chr(9) & iv & Chr(9) & r & Chr(9) & rs("num_int") & Chr(9) & obserr
        Else
           msf1.AddItem er & Chr$(9) & F & Chr(9) & rs("cliente02") & Chr(9) & cic & Chr$(9) & nic & Chr$(9) & tc & Chr$(9) & PtV & Chr$(9) & nc & Chr(9) & "" & Chr(9) & "" & Chr$(9) & n & Chr(9) & ti & Chr(9) & iv & Chr(9) & r & Chr(9) & rs("num_int") & Chr(9) & obserr
        End If
        p = p + 1
      rs2.MoveNext
     Wend
     Set rs2 = Nothing
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
Set rs2 = New ADODB.Recordset
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
msf1.Cols = 17
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

msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Cliente"
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
  J = MsgBox("Confirma genera archivo para importar del aplicativo Citi Ventas. Archivo: c:\Citiv.txt", 4)
  If J = 6 Then
    Call exporta
  End If
  
End If


End Sub
Sub exporta()
Dim Detalle As String
k = 1
Open "c:\citiv.txt" For Output As #1
ni15 = "000000000000000"
Detalle = String(75, " ")
While k <= msf1.Rows - 1
 
   t = Mid$(Format$(Val(msf1.TextMatrix(k, 8)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 8)), "0000000000000.00"), 15, 2)
   ng = Mid$(Format$(Val(msf1.TextMatrix(k, 9)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 9)), "0000000000000.00"), 15, 2)
   g = Mid$(Format$(Val(msf1.TextMatrix(k, 10)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 10)), "0000000000000.00"), 15, 2)
   a = Mid$(Format$(Val(msf1.TextMatrix(k, 11)), "00.00"), 1, 2) & Mid$(Format$(Val(msf1.TextMatrix(k, 11)), "00.00"), 4, 2)
   i = Mid$(Format$(Val(msf1.TextMatrix(k, 12)), "0000000000000.00"), 1, 13) & Mid$(Format$(Val(msf1.TextMatrix(k, 12)), "0000000000000.00"), 15, 2)
   e = "000000000000000"
   ca = Format$(Val(msf1.TextMatrix(k, 13)), "0")
   nd = Format$(Val(msf1.TextMatrix(k, 4)), "00000000000")
   l = "1" & msf1.TextMatrix(k, 1) & msf1.TextMatrix(k, 5) & " " & msf1.TextMatrix(k, 6) & msf1.TextMatrix(k, 7) & msf1.TextMatrix(k, 7) & msf1.TextMatrix(k, 3) & nd & Left$(Format$(msf1.TextMatrix(k, 2), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30) & t & ng & g & a & i & ni15 & e & ni15 & ni15 & ni15 & ni15 & "00" & "   " & "0000000000" & ca & " 000000000000000000000000000000" & Detalle & "00000000" & ni15
   Print #1, l
   k = k + 1
Wend
Close #1
MsgBox ("Operacion Terminada. Ingrse al prgrama Citi Ventas del siap e Importe los registros generados")

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
