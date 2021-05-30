VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_ivacompras 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SUBDIARIO DE IVA COMPRAS"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   7200
      Width           =   3255
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra totales por Tasa"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3240
      TabIndex        =   9
      Top             =   240
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   218103809
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   7455
      Begin VB.TextBox t_p 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro. página inicial:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   240
         Width           =   1455
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
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Con003.frx":0000
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
         Picture         =   "Con003.frx":0882
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
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
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
            TextSave        =   "30/05/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:28 a.m."
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
Attribute VB_Name = "con_ivacompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Sub carga()
  Call armagrid
  q = "select * from a5, g2, a1, g3 where [grabado] <> 'N' and  [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] and a1.[cod_tipoiva] = g3.[cod_tipoiva]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  q = q & " order by [fecha]"
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  tdbf = 0
  tcrf = 0
  While Not rs.EOF
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("g2.abreviatura")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     If rs("moneda") = "P" Then
        c5 = 1
     Else
        c5 = rs("cotiz_dolar")
     End If
     
     If rs("percep_ret") <> 0 Then
        q = "select * from a13, a12 where a13.[id_percepcion] = a12.[id_percepcion] and [num_int] = " & rs("num_int")
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1
        perc_iva = 0
        perc_otras = 0
        While Not rs1.EOF
          If rs1("impuesto12") = "I" Then 'iva
            perc_iva = perc_iva + rs1("importe")
          Else
            perc_otras = perc_otras + rs1("importe")
          End If
          rs1.MoveNext
        Wend
        Set rs1 = Nothing
      Else
        perc_iva = 0
        perc_otras = 0
      End If
     
     
     If rs("id_tipocomp") <> 97 Then
      If rs("grabado") = "S" Then
        t = Format$(rs("total") * c5, "######0.00")
        i = Format$(rs("a5.iva") * c5, "######0.00")
        s = Format$(rs("subtotal") * c5, "######0.00")
        ng = Format$((rs("no_grabado") + perc_otras) * c5, "######0.00")
        rp = Format$(perc_iva * c5, "######0.00")
        tcrf = tcrf + Val(i)
      Else
        t = Format$(-rs("total") * c5, "######0.00")
        i = Format$(-rs("a5.iva") * c5, "######0.00")
        s = Format$(-rs("subtotal") * c5, "######0.00")
        ng = Format$(-(rs("no_grabado") + perc_otras) * c5, "######0.00")
        rp = Format$(-perc_iva * c5, "######0.00")
        tdbf = tdbf + (-Val(i))
      End If
     Else
       'retencion iva
         t = Format$(rs("total") * c5, "######0.00")
        i = Format$(0, "######0.00")
        s = Format$(0, "######0.00")
        ng = Format$(0, "######0.00")
        rp = Format$(rs("total") * c5, "######0.00")
     End If
     
     tt = tt + Val(t)
     ti = ti + Val(i)
     ts = ts + Val(s)
     tng = tng + Val(ng)
     trp = trp + Val(rp)
     
     msf1.AddItem F & Chr(9) & rs("proveedor05") & Chr(9) & rs("cuit05") & " " & rs("g3.abreviatura") & Chr(9) & tc & " " & nc & Chr(9) & s & Chr(9) & rp & Chr(9) & i & Chr(9) & ng & Chr(9) & t & Chr(9) & Format$(rs("num_int"), "00000")

    rs.MoveNext
  Wend
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Totales:" & Chr(9) & Format$(ts, "######0.00") & Chr(9) & Format$(trp, "######0.00") & Chr(9) & Format$(ti, "######0.00") & Chr(9) & Format$(tng, "######0.00") & Chr(9) & Format$(tt, "######0.00")
  msf1.AddItem " "
  msf1.AddItem " "
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Total Db. Fiscal  :" & Chr(9) & Format$(tdbf, "######0.00")
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Total Cr. Fiscal  :" & Chr(9) & Format$(tcrf, "######0.00")
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Total Ret/Perc Iva:" & Chr(9) & Format$(trp, "######0.00")
   
 salto = 0
 If Check1 = 1 Then
   msf1.AddItem "*"
   Call portasa
   salto = 1
 End If
  
   
End Sub

Sub portasa()
  'acumulacion por tasa iva
  espere.Label1 = "Espere...... Calculando Totales por Tasa"
  espere.Refresh
  
  msf1.AddItem " " & Chr$(9) & "Totales por Tasa Iva -CREDITOS-"
  msf1.AddItem " "
  q = "select * from a5, a23 where [grabado] <> 'N' and  a5.[num_int] = a23.[num_int] "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  q = q & " order by [tasa_iva]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  nt = 0
  nd = 0
  nf = 0
  IT = 0
  id = 0
  if2 = 0
  tt = 0
  td = 0
  tf = 0
  tasa = 0
  retd = 0
  retf = 0
  impu = 0
  impud = 0
  impuf = 0
  nc = 0
  ic = 0
  tc = 0
  If Not rs.EOF And Not rs.BOF Then
   ni = rs("a5.num_int")
   While Not rs.EOF
    nt = 0
    IT = 0
    tt = 0
    ret = 0
    ntc = 0
    ITc = 0
    If rs("moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotiz_dolar")
     End If
    
    If rs("tasa_iva") = tasa Then

        If rs("grabado") = "S" Then
           nt = rs("neto")
           IT = rs("a23.iva")
        Else
           'ntc = rs("neto")
           'ITc = rs("vta_09.iva")
        End If
        
      nd = nd + Format(nt * c5, "########0.00")
      id = id + Format(IT * c5, "########0.00")
      td = td + (nt + IT)
             
             
      rs.MoveNext
 
    Else
       msf1.AddItem " " & Chr$(9) & " " & Chr$(9) & " " & Chr$(9) & "Tasa " & tasa & "%" & Chr$(9) & Format$(nd, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(id, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(td, "######0.00")
       tasa = rs("tasa_iva")
       nf = nf + nd
       if2 = if2 + id
       tf = tf + (td + id)
       nd = 0
       id = 0
       td = 0
    End If
  Wend
  nf = nf + nd
  if2 = if2 + id
  tf = nf + if2 + tng + trp
  msf1.AddItem " " & Chr$(9) & " " & Chr$(9) & " " & Chr$(9) & "Tasa " & tasa & "%" & Chr$(9) & Format$(nd, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(id, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(td, "######0.00")
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr$(9) & " " & Chr$(9) & "______________________" & Chr$(9) & "______________________" & Chr$(9) & "______________________" & Chr$(9) & "______________________" & Chr$(9) & "______________________"
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr$(9) & "Totales:" & Chr(9) & Format$(nf, "######0.00") & Chr$(9) & Format$(trp, "#####0.00") & Chr$(9) & Format$(if2, "######0.00") & Chr$(9) & Format$(tng, "#####0.00") & Chr$(9) & Format$(tf, "######0.00")
  Set rs = Nothing
 End If
  msf1.AddItem " "
    
  
  msf1.AddItem " " & Chr$(9) & "Totales por Tasa Iva -DEBITOS-"
  msf1.AddItem " "
  q = "select * from a5, a23 where [grabado] <> 'N' and  a5.[num_int] = a23.[num_int] "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  q = q & " order by [tasa_iva]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  nt = 0
  nd = 0
  nf = 0
  IT = 0
  id = 0
  if2 = 0
  tt = 0
  td = 0
  tf = 0
  tasa = 0
  retd = 0
  retf = 0
  impu = 0
  impud = 0
  impuf = 0
  nc = 0
  ic = 0
  tc = 0
  If Not rs.EOF And Not rs.BOF Then
   ni = rs("a5.num_int")
   While Not rs.EOF
    nt = 0
    IT = 0
    tt = 0
    ret = 0
    ntc = 0
    ITc = 0
    If rs("moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotiz_dolar")
     End If
    
    If rs("tasa_iva") = tasa Then

        If rs("grabado") = "R" Then
           nt = rs("neto")
           IT = rs("a23.iva")
        Else
           'ntc = rs("neto")
           'ITc = rs("vta_09.iva")
        End If
        
      
      nd = nd + Format(nt * c5, "########0.00")
      id = id + Format(IT * c5, "########0.00")
      td = td + (nt + IT)
             
             
      rs.MoveNext
 
    Else
       msf1.AddItem " " & Chr$(9) & " " & Chr$(9) & " " & Chr$(9) & "Tasa " & tasa & "%" & Chr$(9) & Format$(nd, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(id, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(td, "######0.00")
       tasa = rs("tasa_iva")
       nf = nf + nd
       if2 = if2 + id
       tf = tf + (td + id)
       nd = 0
       id = 0
       td = 0
    End If
  Wend
  nf = nf + nd
  if2 = if2 + id
  tf = nf + if2 + tng + trp
  msf1.AddItem " " & Chr$(9) & " " & Chr$(9) & " " & Chr$(9) & "Tasa " & tasa & "%" & Chr$(9) & Format$(nd, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(id, "######0.00") & Chr$(9) & " " & Chr$(9) & Format$(td, "######0.00")
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr$(9) & " " & Chr$(9) & "______________________" & Chr$(9) & "______________________" & Chr$(9) & "______________________" & Chr$(9) & "______________________" & Chr$(9) & "______________________"
  msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr$(9) & "Totales:" & Chr(9) & Format$(nf, "######0.00") & Chr$(9) & Format$(trp, "#####0.00") & Chr$(9) & Format$(if2, "######0.00") & Chr$(9) & Format$(tng, "#####0.00") & Chr$(9) & Format$(tf, "######0.00")
  Set rs = Nothing
 End If
 
 
 
 
 
 espere.Label1 = "Espere...... Calculando totales por Tipo de Contribuyente"
 espere.Refresh

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
msf1.Cols = 10
msf1.ColWidth(0) = 900
msf1.ColWidth(1) = 3000
msf1.ColWidth(2) = 1800
msf1.ColWidth(3) = 2200
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1000


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Proveedor"
msf1.TextMatrix(0, 2) = "Cuit "
msf1.TextMatrix(0, 3) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 4) = "Subtotal  "
msf1.TextMatrix(0, 5) = "Ret/Per Iva"
msf1.TextMatrix(0, 6) = "Iva"
msf1.TextMatrix(0, 7) = "No Grav."
msf1.TextMatrix(0, 8) = "Total"
msf1.TextMatrix(0, 9) = "Num.Int."

For i = 0 To 3
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 4 To 8
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()

Call barraesag(Me)
cal1.Visible = False
Call armagrid
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F6] Archivo Texto - [F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
 J = MsgBox("Prepare Impresora y confirme", 4)
 If J = 6 Then
  Call imprimegrid_l
 End If
End If



If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Sub imprimegrid_l()
 
     
   'On Error GoTo errifg
      
   Printer.Orientation = 1
     
    fuente = 6
    linea = 2
    pag = Val(t_p)
    Call imprimeempresa(14)
    linea = linea + 5
    Printer.FontSize = fuente + 4
    Printer.Print
    Printer.Print "LISTADO DE IVA COMPRAS "
    Printer.Print
    Printer.FontName = "Courier New"
    Printer.Print "Periodo: " & t_fecha & " - " & t_fecha2
    Printer.Print Tab(70); "Hoja Nro: " & pag
    Printer.FontSize = fuente
    nh = 1
    Fila = 0
    linea = linea + 3
    cab = 0
    lph = 90
    t = "____________________________________________________________________________________________________"
    margen = "    "
    While Fila < msf1.Rows
     If linea <= lph And msf1.TextMatrix(Fila, 0) <> "*" Then
       Text = ""
       For col = 0 To 8   'columnas
           tamañocol = Int(msf1.ColWidth(col) / 100)
           item = Space$(tamañocol)
           If msf1.ColAlignment(col) = 1 Then 'izq textos
              LSet item = msf1.TextMatrix(Fila, col)
           Else
              RSet item = msf1.TextMatrix(Fila, col)
           End If
           Text = Text & "  " & item
       Next col
       If cab = 0 Then
          If Fila = 0 Then
             t = "_"
             For i = 1 To Len(margen & Text)
                t = t & "_"
             Next i
             primera = Text
             Fila = Fila + 1
          End If
          Call imprimelinea(t, fuente, False, False, 1)
          Call imprimelinea(margen & primera, fuente, False, False, 1)
          Call imprimelinea(t, fuente, False, False, 1)
          cab = 1
       Else
          Call imprimelinea(margen & Text, fuente, False, False, 1)
          Fila = Fila + 1
          linea = linea + 1
       End If
     Else
      If para.imprime_pie_reportes = True Then
       If msf1.TextMatrix(Fila, 0) = "*" Then
          For i = linea To lph
            Printer.Print
          Next i
       End If
       Printer.Print "________________________________________________________________"
       Printer.Print "Fecha Imp." & Format$(Now, "dd/mm/yyyy") & "   Nro.Hoja: " & Format$(nh, "000") & "     Emitido por: " & glo.usuario
      End If
      Printer.NewPage
      nh = nh + 1
      fuente = 6
      linea = 2
      pag = pag + 1
      Call imprimeempresa(14)
      linea = linea + 5
      Printer.FontSize = fuente + 4
      Printer.Print
      If msf1.TextMatrix(Fila, 0) <> "*" Then
        Printer.Print "LISTADO DE IVA COMPRAS "
      Else
        Printer.Print "LISTADO DE IVA COMPRAS TOTALES"
        Fila = Fila + 1
      End If
      Printer.Print
      Printer.FontName = "Courier New"
      Printer.FontSize = fuente + 2
      Printer.Print "Periodo: " & t_fecha & " - " & t_fecha2
      Printer.Print Tab(70); "Hoja Nro: " & pag
      cab = 0
      End If
     Wend
     If para.imprime_pie_reportes = True Then
      For i = linea To lph
        Printer.Print
      Next i
     
      Printer.Print "________________________________________________________________"
      Printer.Print "Fecha Imp." & Format$(Now, "dd/mm/yyyy") & "   Nro.Hoja: " & Format$(nh, "000") & "     Emitido por: " & glo.usuario
     End If
     Printer.EndDoc

Exit Sub
errifg:
g = MsgBox("Error de Impresion. Continua?", 4)
If g = 6 Then
   Resume
Else
   Printer.KillDoc
   Exit Sub
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 9)
    cc_detalle.Show
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
