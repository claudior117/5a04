VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form gen_posicioniva 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POSICION IVA"
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
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   10335
      Begin VB.TextBox t_montoutilizado 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_saldotecnico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_saldolibre 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Monto utilizado"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7080
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Saldo Tecnico a favor anterior"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Saldo libre disp. a favor anterior"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Formato"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   7680
      Width           =   4575
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Papeles trabajo Siap"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Planilla Resumen"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   735
      Left            =   7200
      TabIndex        =   16
      Top             =   7680
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "Verifica Totales"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   7320
      TabIndex        =   13
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3720
      TabIndex        =   11
      Top             =   1800
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   226885633
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
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
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   5
      Top             =   7560
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "GEN023.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "GEN023.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
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
            TextSave        =   "01/04/2025"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "01:00 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5535
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9763
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
End
Attribute VB_Name = "gen_posicioniva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim c5 As Double

Sub carga()
 'saca el debito fiscal
 'las nc de venta son creditos fiscal
 'tambien se separan las ret iva
 l = "________________________________"
 ll = "--------------------------->>"
 Dim netoV(9) As Double
 Dim ivaV(9) As Double
 Dim Neto_creditosV(9) As Double
 Dim iva_creditosV(9) As Double
 espere.Show
 espere.Label1 = "Espere...... Generando Listado de Iva"
 espere.Refresh
 
ret_iva = 0
'ventas
 q = "select * from VTA_02 where [grabado] <> 'N' "
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
  q = q & " order by [fecha], [letra], [num_comp]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  
  For i = 0 To 8
    netoV(i) = 0
    ivaV(i) = 0
    Neto_creditosV(i) = 0
    iva_creditosV(9) = 0
  Next i
  
  While Not rs.EOF
     If rs("moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     If rs("id_tipocomp") <> 101 Then 'retencion de iva
       If rs("grabado") = "S" Then
        netoV(rs("id_tipo_iva02")) = netoV(rs("id_tipo_iva02")) + rs("subtotal") * c5
        ivaV(rs("id_tipo_iva02")) = ivaV(rs("id_tipo_iva02")) + rs("iva") * c5
        ng = Format$((rs("impuestos") + rs("perc_ib") + rs("perc_gan")) * c5, "######0.00")
        rp = Format$(rs("perc_iva") * c5, "######0.00") 'ret/perc iva
       Else
        Neto_creditosV(rs("id_tipo_iva02")) = Neto_creditosV(rs("id_tipo_iva02")) + rs("subtotal") * c5
        iva_creditosV(rs("id_tipo_iva02")) = iva_creditosV(rs("id_tipo_iva02")) + rs("iva") * c5

       End If
       
        If rs("id_tipocomp") >= 205 And rs("id_tipocomp") <= 207 Then 'venta directa
            q = "select * from vta_012, a12 where [id_retencion] = [id_percepcion] and [num_int] = " & rs("num_int")
            Set rs1 = New ADODB.Recordset
            rs1.Open q, cn1
            r_iva = 0
            r_otras = 0
            While Not rs1.EOF
              If rs1("impuesto12") = "I" Then 'iva
                r_iva = r_iva + rs1("importe")
              Else
                r_otras = r_otras + rs1("importe")
              End If
             rs1.MoveNext
            Wend
            Set rs1 = Nothing
            
            If rs("id_tipocomp") <> 207 Then
              'fat y nd
              ng = Format$(Val(ng) + (r_otras * c5), "######0.00")
              rp = Format$(Val(rp) + (r_iva * c5), "######0.00") 'ret/perc iva
           Else
              ng = Format$(Val(ng) - (r_otras * c5), "######0.00")
              rp = Format$(Val(rp) - (r_iva * c5), "######0.00") 'ret/perc iva
           End If
       End If
       
     Else
        ret_iva = ret_iva + rs("total") * c5
     End If
   
     tng = tng + Val(ng)
     trp = trp + Val(rp)
     
    rs.MoveNext
  Wend
  
  msf1.AddItem "VENTAS"
  msf1.AddItem "    Debito Fiscal"
  msf1.AddItem "        Op. Resp. Insc." & Chr(9) & Format$(netoV(1) + netoV(6), "######0.00") & Chr(9) & Format$(ivaV(1) + ivaV(6), "######0.00")
  msf1.AddItem "        Op. Resp. No Insc." & Chr(9) & Format$(netoV(2), "######0.00") & Chr(9) & Format$(ivaV(2), "######0.00")
  msf1.AddItem "        Op. CF, Exento y No Alc." & Chr(9) & Format$(netoV(3) + netoV(5) + netoV(7), "######0.00") & Chr(9) & Format$(ivaV(3) + ivaV(5) + ivaV(7), "######0.00")
  msf1.AddItem "        Op. Monotirbuto" & Chr(9) & Format$(netoV(4), "######0.00") & Chr(9) & Format$(ivaV(4), "######0.00")
  msf1.AddItem "        Op. Exentas Exportaciones" & Chr(9) & Format$(netoV(8), "######0.00") & Chr(9) & Format$(ivaV(8), "######0.00")
  msf1.AddItem "        " & Chr(9) & l & Chr(9) & l
  tt = 0
  ti = 0
  ttc = 0
  tic = 0
  ti2 = 0
  For i = 0 To 8
    tt = tt + Format(netoV(i), "######0.00")
    ti = ti + Format(ivaV(i), "######0.00")
    ttc = ttc + Format(Neto_creditosV(i), "######0.00")
    tic = tic + Format(iva_creditosV(i), "######0.00")
   
  Next i
  msf1.AddItem "" & Chr(9) & Format$(tt, "######0.00") & Chr(9) & Format$(ti, "######0.00")
  Set rs = Nothing
  ti2 = ti
 
 '---------------------------------------------------------------------------------------------------
 'compras
   
  q = "select * from a5 where [grabado] <> 'N' "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  
  tnc = 0
  tinc = 0
  
  perc_iva = 0
  While Not rs.EOF
     If rs("moneda") = "P" Then
        c5 = 1
     Else
        c5 = rs("cotiz_dolar")
     End If
     
     
     If rs("percep_ret") <> 0 Then
        q = "select * from a13 where [num_int] = " & rs("num_int")
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1
        While Not rs1.EOF
          If rs1("id_percepcion") = 1 Then 'iva
            perc_iva = perc_iva + (rs1("importe") * c5)
          End If
          rs1.MoveNext
        Wend
        Set rs1 = Nothing
     End If
     
     If rs("id_tipocomp") <> 97 Then
      If rs("grabado") = "S" Then
        t = t + (rs("total") * c5)
        i = i + (rs("iva") * c5)
        s = s + (rs("subtotal") * c5)
        ng = ng + ((rs("no_grabado") + perc_otras) * c5)
        rp = rp + (perc_iva * c5)
      Else
        'nota credito compras
        tinc = tinc + (rs("iva") * c5)
        rp = Format$(-perc_iva * c5, "######0.00")
      End If
     Else
       'retencion iva
         perc_iva = perc_iva + (rs("total") * c5)
     End If
     
     
    rs.MoveNext
  Wend
  
  msf1.AddItem "    Debito Fiscal por NC Compras" & Chr$(9) & "" & Chr$(9) & Format$(tinc, "######0.00") 'nota credito de compras
  tdf = ti2 + tic
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & ll & Chr(9) & Format$(tdf, "######0.00")
  msf1.AddItem ""
  msf1.AddItem ""
  msf1.AddItem "COMPRAS"
  msf1.AddItem "    Credito Fiscal"
  msf1.AddItem "        Compras a Inscriptos" & Chr(9) & Format$(s, "######0.00") & Chr(9) & Format$(i, "######0.00")
  
  msf1.AddItem "    Credito Fiscal por NC Ventas"
  msf1.AddItem "        Op. Resp. Insc." & Chr(9) & Format$(Neto_creditosV(1) + Neto_creditosV(6), "######0.00") & Chr(9) & Format$(iva_creditosV(1) + iva_creditosV(6), "######0.00")
  msf1.AddItem "        Op. Resp. No Insc." & Chr(9) & Format$(Neto_creditosV(2), "######0.00") & Chr(9) & Format$(iva_creditosV(2), "######0.00")
  msf1.AddItem "        Op. CF, Exento y No Alc." & Chr(9) & Format$(Neto_creditosV(3) + Neto_creditosV(5) + Neto_creditosV(7), "######0.00") & Chr(9) & Format$(iva_creditosV(3) + iva_creditosV(5) + iva_creditosV(7), "######0.00")
  msf1.AddItem "        Op. Monotirbuto" & Chr(9) & Format$(Neto_creditosV(4), "######0.00") & Chr(9) & Format$(iva_creditosV(4), "######0.00")
  msf1.AddItem "        Op. Exentas Exportaciones" & Chr(9) & Format$(Neto_creditosV(8), "######0.00") & Chr(9) & Format$(iva_creditosV(8), "######0.00")
  msf1.AddItem "        " & Chr(9) & l & Chr(9) & l
  msf1.AddItem "" & Chr(9) & Format$(ttc, "######0.00") & Chr(9) & Format$(tic, "######0.00")
  tcf = i + tic
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & ll & Chr(9) & Format$(tcf, "######0.00")

  Set rs = Nothing

  
  msf1.AddItem "    Otros"
  msf1.AddItem "        Retenciones/Percepcionesd IVA" & Chr(9) & "" & Chr(9) & Format$(perc_iva + ret_iva, "######0.00")
  msf1.AddItem ""
  saldot = Val(t_saldotecnico) - ti2 + tic + i - tinc - tinc
  If saldot <= 0 Then
    saldot = 0
  End If
  msf1.AddItem "    Saldo a Favor Tecnico" & Chr(9) & ll & Chr(9) & Format$(saldot, "######0.00")
 
  saldol = Val(t_saldolibre) + perc_iva + ret_iva - Val(t_montoutilizado)
  If saldol <= 0 Then
    saldol = 0
  End If
  msf1.AddItem "    Saldo a Favor Libre Disponibilidad" & Chr(9) & ll & Chr(9) & Format$(saldol, "######0.00")
  
  saldol = Val(t_saldolibre) + perc_iva + ret_iva - Val(t_montoutilizado)
  
  msf1.AddItem "    Monto a pagar" & Chr(9) & ll
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  Set rs = Nothing
 Unload espere
     
End Sub


Sub carga2()
 'saca el debito fiscal
 'las nc de venta son creditos fiscal
 'tambien se separan las ret iva
 l = "________________________________"
 ll = "--------------------------->>"
 Dim netoV(9) As Double
 Dim ivaV(9) As Double
 Dim Neto_creditosV(9) As Double
 Dim iva_creditosV(9) As Double
 espere.Show
 espere.Label1 = "Espere...... Generando Listado de Iva"
 espere.Refresh
 
ret_iva = 0
'ventas
 q = "select * from VTA_02 where [grabado] <> 'N' "
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
  q = q & " order by [fecha], [letra], [num_comp]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  
  For i = 0 To 8
    netoV(i) = 0
    ivaV(i) = 0
    Neto_creditosV(i) = 0
    iva_creditosV(9) = 0
  Next i
  
  While Not rs.EOF
     If rs("moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     If rs("id_tipocomp") <> 101 Then 'retencion de iva
       If rs("grabado") = "S" Then
        netoV(rs("id_tipo_iva02")) = netoV(rs("id_tipo_iva02")) + rs("subtotal") * c5
        ivaV(rs("id_tipo_iva02")) = ivaV(rs("id_tipo_iva02")) + rs("iva") * c5
        ng = Format$((rs("impuestos") + rs("perc_ib") + rs("perc_gan")) * c5, "######0.00")
        rp = Format$(rs("perc_iva") * c5, "######0.00") 'ret/perc iva
       Else
        Neto_creditosV(rs("id_tipo_iva02")) = Neto_creditosV(rs("id_tipo_iva02")) + rs("subtotal") * c5
        iva_creditosV(rs("id_tipo_iva02")) = iva_creditosV(rs("id_tipo_iva02")) + rs("iva") * c5
         netoV(rs("id_tipo_iva02")) = netoV(rs("id_tipo_iva02")) - rs("subtotal") * c5
        ivaV(rs("id_tipo_iva02")) = ivaV(rs("id_tipo_iva02")) - rs("iva") * c5
       End If
       
        If rs("id_tipocomp") >= 205 And rs("id_tipocomp") <= 207 Then 'venta directa
            q = "select * from vta_012, a12 where [id_retencion] = [id_percepcion] and [num_int] = " & rs("num_int")
            Set rs1 = New ADODB.Recordset
            rs1.Open q, cn1
            r_iva = 0
            r_otras = 0
            While Not rs1.EOF
              If rs1("impuesto12") = "I" Then 'iva
                r_iva = r_iva + rs1("importe")
              Else
                r_otras = r_otras + rs1("importe")
              End If
             rs1.MoveNext
            Wend
            Set rs1 = Nothing
            
            If rs("id_tipocomp") <> 207 Then
              'fat y nd
              ng = Format$(Val(ng) + (r_otras * c5), "######0.00")
              rp = Format$(Val(rp) + (r_iva * c5), "######0.00") 'ret/perc iva
           Else
              ng = Format$(Val(ng) - (r_otras * c5), "######0.00")
              rp = Format$(Val(rp) - (r_iva * c5), "######0.00") 'ret/perc iva
           End If
       End If
       
     Else
        ret_iva = ret_iva + rs("total") * c5
     End If
   
     tng = tng + Val(ng)
     trp = trp + Val(rp)
     
    rs.MoveNext
  Wend
  
  msf1.AddItem "VENTAS Debito Fiscal"
  msf1.AddItem "        Op. Resp. Insc." & Chr(9) & Format$(netoV(1) + netoV(6), "######0.00") & Chr(9) & Format$(ivaV(1) + ivaV(6), "######0.00")
  msf1.AddItem "        Op. Resp. No Insc." & Chr(9) & Format$(netoV(2), "######0.00") & Chr(9) & Format$(ivaV(2), "######0.00")
  msf1.AddItem "        Op. CF, Exento y No Alc." & Chr(9) & Format$(netoV(3) + netoV(5) + netoV(7), "######0.00") & Chr(9) & Format$(ivaV(3) + ivaV(5) + ivaV(7), "######0.00")
  msf1.AddItem "        Op. Monotirbuto" & Chr(9) & Format$(netoV(4), "######0.00") & Chr(9) & Format$(ivaV(4), "######0.00")
  msf1.AddItem "        Op. Exentas Exportaciones" & Chr(9) & Format$(netoV(8), "######0.00") & Chr(9) & Format$(ivaV(8), "######0.00")
  msf1.AddItem "        " & Chr(9) & l & Chr(9) & l
  tt = 0
  ti = 0
  ttc = 0
  tic = 0
  ti2 = 0
  For i = 0 To 8
    tt = tt + Format(netoV(i), "######0.00")
    ti = ti + Format(ivaV(i), "######0.00")
    ttc = ttc + Format(Neto_creditosV(i), "######0.00")
    tic = tic + Format(iva_creditosV(i), "######0.00")
   
  Next i
  Set rs = Nothing
  ti2 = ti
  tt2 = tt
 '---------------------------------------------------------------------------------------------------
 'compras
   
  q = "select * from a5 where [grabado] <> 'N' "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  
  tnc = 0
  tinc = 0
  
  perc_iva = 0
  While Not rs.EOF
     If rs("moneda") = "P" Then
        c5 = 1
     Else
        c5 = rs("cotiz_dolar")
     End If
     
     
     If rs("percep_ret") <> 0 Then
        q = "select * from a13 where [num_int] = " & rs("num_int")
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1
        While Not rs1.EOF
          If rs1("id_percepcion") = 1 Then 'iva
            perc_iva = perc_iva + (rs1("importe") * c5)
          End If
          rs1.MoveNext
        Wend
        Set rs1 = Nothing
     End If
     
     If rs("id_tipocomp") <> 97 Then
      If rs("grabado") = "S" Then
        t = t + (rs("total") * c5)
        i = i + (rs("iva") * c5)
        s = s + (rs("subtotal") * c5)
        ng = ng + ((rs("no_grabado") + perc_otras) * c5)
        rp = rp + (perc_iva * c5)
      Else
        'nota credito compras
        tinc = tinc + (rs("iva") * c5)
        rp = Format$(-perc_iva * c5, "######0.00")
        
        i = i - (rs("iva") * c5)
        s = s - (rs("subtotal") * c5)
      
      
      End If
     Else
       'retencion iva
         perc_iva = perc_iva + (rs("total") * c5)
     End If
     
     
    rs.MoveNext
  Wend
  
  msf1.AddItem "TOTALES" & Chr(9) & Format$(tt2, "######0.00") & Chr(9) & Format$(ti2, "######0.00") & Chr$(9) & Format$(tinc, "######0.00") & Chr(9) & Format$(s, "######0.00") & Chr(9) & Format$(i, "######0.00") & Chr(9) & Format$(perc_iva + ret_iva, "######0.00")
  tdf = ti2 + tic
  tcf = i + tic
  
  msf1.AddItem ""
  
  msf1.AddItem "    Total Debito Fiscal " & Chr$(9) & ll & Chr(9) & Format$(tdf, "######0.00")
  msf1.AddItem "    Total Credito Fiscal " & Chr$(9) & ll & Chr(9) & Format$(tcf, "######0.00")
  

  Set rs = Nothing

  
  msf1.AddItem ""
  
  saldot = Val(t_saldotecnico) - ti2 - tinc + i
  If saldot <= 0 Then
    saldot = 0
  End If
  msf1.AddItem "    Saldo a Favor Tecnico" & Chr(9) & ll & Chr(9) & Format$(saldot, "######0.00")
 
  saldol = Val(t_saldolibre) + perc_iva + ret_iva - Val(t_montoutilizado)
  If saldol <= 0 Then
    saldol = 0
  End If
  msf1.AddItem "    Saldo a Favor Libre Disponibilidad" & Chr(9) & ll & Chr(9) & Format$(saldol, "######0.00")
  
  
  msf1.AddItem "    Monto a pagar" & Chr(9) & ll
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  Set rs = Nothing
 Unload espere
     
End Sub




Private Sub btnacepta_Click()
If Option2 Then
    Call armagrid
    Call carga 'papeles de trabajo
Else
    Call armagrid2
    Call carga2 'planilla resumen
End If
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
msf1.Cols = 6
msf1.ColWidth(0) = 5000
msf1.ColWidth(1) = 2000
msf1.ColWidth(2) = 2000
msf1.ColWidth(3) = 1600
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 1800


'msf1.TextMatrix(0, 0) = ""
'msf1.TextMatrix(0, 1) = "Subtotal  "
'msf1.TextMatrix(0, 2) = "Ret/Per Iva"
'msf1.TextMatrix(0, 3) = "Iva"
'msf1.TextMatrix(0, 4) = "No Grav./Otros"
'msf1.TextMatrix(0, 5) = "Total"

msf1.TextMatrix(0, 0) = "Titulo"
msf1.TextMatrix(0, 1) = "Neto  "
msf1.TextMatrix(0, 2) = "Iva"


For i = 0 To 0
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 1 To 5
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub


Sub armagrid2()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.ColWidth(0) = 5000
msf1.ColWidth(1) = 1600
msf1.ColWidth(2) = 1600
msf1.ColWidth(3) = 1600
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 1600
msf1.ColWidth(6) = 1600




msf1.TextMatrix(0, 0) = "Titulo"
msf1.TextMatrix(0, 1) = "Neto Ventas "
msf1.TextMatrix(0, 2) = "Debito Fiscal"
msf1.TextMatrix(0, 3) = "DF NC Compras"
msf1.TextMatrix(0, 4) = "Neto Compras"
msf1.TextMatrix(0, 5) = "Credito Fiscal"
msf1.TextMatrix(0, 6) = "Retenciones"


For i = 0 To 0
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 1 To 5
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub


Private Sub Form_Load()
Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0

Call barraesag(Me)
cal1.Visible = False
Option1 = True
Call armagrid
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel"

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
    
    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "POSICION FRENTE AL IVA", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
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
