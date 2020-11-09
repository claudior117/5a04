VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_estadoresultado 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ESTADO DE RESULTADO"
   ClientHeight    =   8715
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Caption         =   "Opciones"
      Height          =   855
      Left            =   8880
      TabIndex        =   23
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox Check2 
         Caption         =   "Calc. Gastos Bancarios"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calc. Costo Mercaderia"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.ComboBox c_sucursal 
      Height          =   315
      Left            =   5640
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   360
      TabIndex        =   20
      Top             =   6960
      Width           =   9135
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   $"gen020.frx":0000
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   6360
      Width           =   9135
      Begin VB.Label Label2 
         Caption         =   "Este proceso involucra informacion historica, puede demorar y hacer bajar el rendimiento de la aplicacion. "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Meses"
      Height          =   855
      Left            =   5280
      TabIndex        =   15
      Top             =   0
      Width           =   1335
      Begin VB.TextBox t_meses 
         Height          =   495
         Left            =   120
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   495
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultado Anterior"
      Height          =   855
      Left            =   6720
      TabIndex        =   13
      Top             =   0
      Width           =   2055
      Begin VB.TextBox t_saldoant 
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox t_ultimodia 
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox t_primerdia 
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   4695
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox T_FECHA2 
         Height          =   435
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         Height          =   450
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen020.frx":0219
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen020.frx":0A9B
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   8460
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
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
            TextSave        =   "09:40"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9551
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
Attribute VB_Name = "gen_estadoresultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim o As Integer
Dim gcol As Integer
Dim gii, gfi As Integer  'intervalo ingresos
Dim glinea, l_rb, l_r As Integer



Private Sub btnacepta_Click()
Call limpia
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub




Sub limpia()
  Call armagrid
  espere.Show
  espere.Label1 = "ESPERE........  [Generando Informe]"
  espere.Refresh
  Call opcion1
  Unload espere
  msf1.SetFocus
  End Sub

Sub opcion1()
 Call armagrid
 msf1.AddItem "VENTAS NETAS(s/iva)"
 For i = 0 To c_sucursal.ListCount - 1
   msf1.AddItem "P.V. " & c_sucursal.List(i)
   msf1.AddItem "   Contado"
   msf1.AddItem "   Cta.Cte."
   msf1.AddItem "   Notas Cr."
   msf1.AddItem ""
   msf1.AddItem "TOT. VENTAS P.V. " & c_sucursal.List(i)
   msf1.AddItem ""
 Next i
 
'msf1.AddItem ""
msf1.AddItem "        VENTAS TOTALES "
l_tv = msf1.Rows - 1
 msf1.AddItem ""
 msf1.AddItem "COSTOS VENTAS"
 msf1.AddItem "   Stock Inicial"
 msf1.AddItem "   Compras Merc."
 msf1.AddItem "   Stock Final"
 msf1.AddItem ""
 msf1.AddItem "(SI+Compras-SF)=COSTO "
 l_c = msf1.Rows - 1
 msf1.AddItem ""
 msf1.AddItem "RESULTADO BRUTO"
 l_rb = msf1.Rows - 1
 msf1.AddItem ""
 msf1.AddItem ""
 msf1.AddItem "COMPRAS NETAS(s/iva) "
 glinea = 19
 For i = 1 To msf1.Cols - 3
   Call ventas(i)
   Call COSTOS(i)
   msf1.TextMatrix(glinea + 2, i) = Format$(Val(msf1.TextMatrix(l_tv, i)) - Val(msf1.TextMatrix(l_c, i)), "######0.00") 'resultado bruto
 Next i
 glinea = glinea + 5
 Call sacacompras


 
 
 
End Sub

Sub ventas(ByVal i As Integer)
ttv = 0
l = 3
fecha = msf1.TextMatrix(0, i)
primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
ultimo = DateSerial(Year(fecha), Month(fecha) + 1, 0)

For J = 0 To c_sucursal.ListCount - 1
   q = "select * from vta_02 where [venta] <> 'N' and datevalue([fecha]) >= datevalue('" & primer & "') and datevalue([fecha]) <= datevalue('" & ultimo & "') and [sucursal_ingreso] = " & Val(c_sucursal.List(J))
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   espere.Label1 = "ESPERE........  [Obteniendo informacion de ventas]"
   espere.Refresh
   tcdo = 0
   tctacte = 0
   tnc = 0
   While Not rs.EOF
     If rs("venta") = "R" Then
        tnc = tnc + rs("subtotal")
     Else
       If rs("contado") = "N" Then
          tctacte = tctacte + rs("subtotal")
       Else
          tcdo = tcdo + rs("subtotal")
       End If
     End If
     rs.MoveNext
   Wend
   msf1.TextMatrix(l, i) = Format$(tcdo, "######0.00")
   msf1.TextMatrix(l + 1, i) = Format$(tctacte, "######0.00")
   msf1.TextMatrix(l + 2, i) = Format$(-tnc, "######0.00")
   msf1.TextMatrix(l + 4, i) = Format$(Val(msf1.TextMatrix(l, i)) + Val(msf1.TextMatrix(l + 1, i)) - Val(msf1.TextMatrix(l + 2, i)), "######0.00") 'total ventas
   ttv = ttv + Val(msf1.TextMatrix(l + 4, i))
   l = l + 7
   Set rs = Nothing
 Next J
 
 
 
l = l - 1
msf1.TextMatrix(l, i) = Format$(ttv, "######0.00")
glinea = l
End Sub
Sub sacacompras()
   LI = glinea
   l = glinea
   q = "select * from c_01 where [tipo] = 'C' "
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   espere.Label1 = "ESPERE........  [Obteniendo informacion de Compras]"
   espere.Refresh
   ttc = 0
   'compras
   While Not rs.EOF
    p = 0
    If rs("id_cuenta") <> para.cuenta_inventario Then 'saco todas las compras menos las de mercaderia
     For i = 1 To msf1.Cols - 3
       fecha = msf1.TextMatrix(0, i)
       primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
       ultimo = DateSerial(Year(fecha), Month(fecha) + 1, 0)
       q = "select * from a5 where [compra] <> 'N' and datevalue([fecha]) >= datevalue('" & primer & "') and datevalue([fecha]) <= datevalue('" & ultimo & "') and [id_cuenta] = " & rs("id_cuenta")
       Set rs2 = New ADODB.Recordset
       'MsgBox (q)
       rs2.Open q, cn1
       tc = 0
       While Not rs2.EOF
          If rs2("moneda") = "P" Then
            m = 1
          Else
            m = rs2("cotiz_dolar")
          End If
          If rs2("compra") = "E" Then
            tc = tc + (rs2("subtotal") * m)
          Else
            tc = tc + (rs2("subtotal") * m)
          End If
          
          rs2.MoveNext
       Wend
       Set rs2 = Nothing
       If tc > 0 Then
         If p = 0 Then
           p = 1
           msf1.AddItem rs("descripcion")
           l = l + 1
         End If
         msf1.TextMatrix(l, i) = Format$(tc, "######0.00")
         ttc = ttc + tc
       End If
       
      Next i
     End If
     rs.MoveNext
   Wend
   msf1.AddItem ""
   msf1.AddItem "TOTAL COMPRAS "
   If l > LI Then
     l = l + 2
     For i = 1 To msf1.Cols - 3
       tm = 0
       For J = LI To l
          tm = tm + Val(msf1.TextMatrix(J, i))
       Next J
       msf1.TextMatrix(l, i) = Format$(tm, "######0.00")
     Next i
     lc1 = l
     
   End If

msf1.AddItem ""
msf1.AddItem "OTROS EGRESOS "
   
'SACA GASTOS
   l = l + 2
   LI = l
   q = "select * from c_01 where [tipo] = 'C' "
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   espere.Label1 = "ESPERE........  [Obteniendo informacion de Caja]"
   espere.Refresh
   ttc = 0
   While Not rs.EOF
     p = 0
     For i = 1 To msf1.Cols - 3
       fecha = msf1.TextMatrix(0, i)
       primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
       ultimo = DateSerial(Year(fecha), Month(fecha) + 1, 0)
       q = "select * from cyb_05 where [ubicacion] = 'H' and [modulo] = 'J' and  datevalue([fecha]) >= datevalue('" & primer & "') and datevalue([fecha]) <= datevalue('" & ultimo & "') and [id_cuenta_contra] = " & rs("id_cuenta")
       Set rs2 = New ADODB.Recordset
       'MsgBox (q)
       rs2.Open q, cn1
       tc = 0
       While Not rs2.EOF
         tc = tc + (rs2("importe"))
         rs2.MoveNext
       Wend
       Set rs2 = Nothing
       If tc > 0 Then
         If p = 0 Then
           p = 1
           msf1.AddItem rs("descripcion")
           l = l + 1
         End If
         msf1.TextMatrix(l, i) = Format$(tc, "######0.00")
         ttc = ttc + tc
       End If
      Next i
      rs.MoveNext
   Wend
   msf1.AddItem ""
   msf1.AddItem "TOTAL OTROS EGRESOS "
   l = l + 2
   If l > LI Then
     
     For i = 1 To msf1.Cols - 3
       tm = 0
       For J = LI To l
          tm = tm + Val(msf1.TextMatrix(J, i))
       Next J
       msf1.TextMatrix(l, i) = Format$(tm, "######0.00")
     Next i
     lc2 = l
   
   End If
   
 'gastos bancarios
   

 msf1.AddItem ""
 msf1.AddItem "Db y Cr BANCARIOS "
   
 l = l + 2
 LI = l
 If Check2 = 1 Then
   q = "select * from c_01 where [tipo] = 'C' "
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   espere.Label1 = "ESPERE........  [Obteniendo informacion de bancos]"
   espere.Refresh
   ttc = 0
   While Not rs.EOF
     p = 0
     For i = 1 To msf1.Cols - 3
       fecha = msf1.TextMatrix(0, i)
       primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
       ultimo = DateSerial(Year(fecha), Month(fecha) + 1, 0)
       q = "select * from cyb_04 where ([id_tipomov] = 20 or [id_tipomov] = 30)  and  datevalue([fecha]) >= datevalue('" & primer & "') and datevalue([fecha]) <= datevalue('" & ultimo & "') and [id_cuenta04] = " & rs("id_cuenta") & " and [num_mov_int_compras] = 0"
       Set rs2 = New ADODB.Recordset
       'MsgBox (q)
       rs2.Open q, cn1
       tc = 0
       While Not rs2.EOF
        If rs2("ubicacion") = "D" Then
         tc = tc + (rs2("importe"))
        Else
         tc = tc - (rs2("importe"))
        End If
         rs2.MoveNext
       Wend
       Set rs2 = Nothing
       If tc > 0 Then
         If p = 0 Then
           p = 1
           msf1.AddItem rs("descripcion")
           l = l + 1
         End If
         msf1.TextMatrix(l, i) = Format$(tc, "######0.00")
         ttc = ttc + tc
       End If
      Next i
      rs.MoveNext
   Wend
  End If
   msf1.AddItem ""
   msf1.AddItem "TOTAL Db y Cr bancarios "
   l = l + 2

   If l > LI Then
     For i = 1 To msf1.Cols - 3
       tm = 0
       For J = LI To l
          tm = tm + Val(msf1.TextMatrix(J, i))
       Next J
       msf1.TextMatrix(l, i) = Format$(tm, "######0.00")
     Next i
     lc2 = l
     
   End If
   
     
   
   msf1.AddItem ""
   msf1.AddItem "TOTAL EGRESOS "
   l = l + 2
   For i = 1 To msf1.Cols - 3
       msf1.TextMatrix(l, i) = Format$(Val(msf1.TextMatrix(lc1, i)) + Val(msf1.TextMatrix(lc2, i)), "######0.00")
   Next i
   lc3 = l
   
   
   'TOTALES
    msf1.AddItem ""
 msf1.AddItem "*********************"
 msf1.AddItem "     RESULTADO"
 msf1.AddItem "*********************"
 msf1.AddItem "TOTAL  VENTAS"
 msf1.AddItem "TOTAL GASTOS"
 l = l + 5
 For i = 1 To msf1.Cols - 3
      msf1.TextMatrix(l, i) = Format$(Val(msf1.TextMatrix(l_rb, i)), "######0.00")
 Next i
 l = l + 1
 For i = 1 To msf1.Cols - 3
      msf1.TextMatrix(l, i) = Format$(Val(msf1.TextMatrix(lc3, i)), "######0.00")
 Next i
 l = l + 1
 msf1.AddItem ""
 For i = 1 To msf1.Cols - 3
      msf1.TextMatrix(l, i) = "-------------------------------"
 Next i
 
  msf1.AddItem ""
   l = l + 1
  For i = 1 To msf1.Cols - 3
       msf1.TextMatrix(l, i) = Format$(Val(msf1.TextMatrix(l_rb, i)) - Val(msf1.TextMatrix(lc3, i)), "######0.00")
   Next i
   L4 = l
   l_r = l
   msf1.AddItem ""
   msf1.AddItem "******ACUMULADO******* "
   l = l + 2
   msf1.TextMatrix(l, 1) = Format$(Val(msf1.TextMatrix(L4, 1)) + Val(t_saldoant), "######0.00")
   For i = 2 To msf1.Cols - 3
       msf1.TextMatrix(l, i) = Format$(Val(msf1.TextMatrix(L4, i)) + Val(msf1.TextMatrix(l, i - 1)), "######0.00")
   Next i


End Sub
Sub COSTOS(ByVal i As Integer)
   espere.Label1 = "ESPERE........  [Obteniendo informacion de Costos]"
   espere.Refresh
   fecha = msf1.TextMatrix(0, i)
   primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
   ultimo = DateSerial(Year(fecha), Month(fecha) + 1, 0)
   l = glinea + 3
   'STOCK INICIAL
   If Check1 = 1 Then
    q = "select * from a2 where [id_producto] > 1"
    Set rs = New ADODB.Recordset
    rs.Open q, cn1
    sit = 0
   
   While Not rs.EOF
      Set rs2 = New ADODB.Recordset
      q = "select * from stk_01 where [id_producto] = " & rs("id_producto") & " and datevalue([fecha]) < datevalue('" & primer & "')"
      'MsgBox (q)
      rs2.Open q, cn1
      s = 0
      While Not rs2.EOF
       If rs2("ubicacion") = "E" Then
          s = s + rs2("cantidad")
       Else
          s = s - rs2("cantidad")
       End If
       rs2.MoveNext
      Wend
      Set rs2 = Nothing
      sit = sit + (s * rs("costoreal"))
      rs.MoveNext
    Wend
    Set rs = Nothing
    
    
    msf1.TextMatrix(l, i) = Format$(sit, "######0.00")
    
  End If
  
   'compras del periodo DE LA CUENTA MERCADERIA
   q = "select * from a5 where datevalue([fecha]) >= datevalue('" & primer & "') and datevalue([fecha]) <= datevalue('" & ultimo & "') and [id_cuenta] = " & para.cuenta_inventario & " and [compra] <> 'N'"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   compras = 0
   While Not rs.EOF
     If rs("compra") = "E" Then
         compras = compras + rs("subtotal")
     Else
         compras = compras - rs("subtotal")
     End If
     rs.MoveNext
   Wend
   Set rs = Nothing
   msf1.TextMatrix(l + 1, i) = Format$(compras, "######0.00")
   
  'stock final
  
  If Check1 = 1 Then
   q = "select * from a2 where [id_producto] > 1"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   sfT = 0
   While Not rs.EOF
      Set rs2 = New ADODB.Recordset
      q = "select * from stk_01 where [id_producto] = " & rs("id_producto") & " and datevalue([fecha]) <= datevalue('" & ultimo & "')"
      'MsgBox (q)
      rs2.Open q, cn1
      s = 0
      While Not rs2.EOF
       If rs2("ubicacion") = "E" Then
          s = s + rs2("cantidad")
       Else
          s = s - rs2("cantidad")
       End If
       rs2.MoveNext
      Wend
      Set rs2 = Nothing
      sfT = sfT + (s * rs("costoreal"))
      rs.MoveNext
    Wend
    Set rs = Nothing
    msf1.TextMatrix(l + 2, i) = Format$(sfT, "######0.00")
    
   End If
   msf1.TextMatrix(l + 4, i) = Format$(sit + compras - sfT, "######0.00") 'costo
   glinea = l + 4
End Sub

Sub armagrid()
'armar grilla
gcol = Val(t_meses) + 3
msf1.clear
msf1.Rows = 1
msf1.Cols = gcol
'msf1.FixedCols = 0
'msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 2500
For i = 1 To Val(t_meses)
msf1.ColWidth(i) = 1100
Next i
'msf1.ColWidth(2) = 800
'msf1.ColWidth(3) = 800
'msf1.ColWidth(4) = 800
'msf1.ColWidth(5) = 800
'msf1.ColWidth(6) = 800
'msf1.ColWidth(7) = 800
'msf1.ColWidth(8) = 800
'msf1.ColWidth(9) = 800
'msf1.ColWidth(10) = 800
'msf1.ColWidth(11) = 800
'msf1.ColWidth(12) = 800
msf1.ColWidth(Val(t_meses) + 1) = 1000
msf1.ColWidth(Val(t_meses) + 2) = 1000

msf1.TextMatrix(0, 0) = "Concepto"
msf1.TextMatrix(0, Val(t_meses) + 1) = "Totales"
msf1.TextMatrix(0, Val(t_meses) + 2) = "Promedios"
mp = Val(t_fecha)
ap = Val(t_fecha2)
For i = 1 To Val(t_meses) + 1
  If i <= Val(t_meses) Then
   msf1.TextMatrix(0, i) = Format$(mp, "00") & "/" & Format$(ap, "00")
  Else
   t_ultimodia = Format$(DateSerial(ap, mp, 1) - 1, "dd/mm/yyyy")
  End If
  mp = mp + 1
  If mp > 12 Then
    mp = 1
    ap = ap + 1
  End If
Next i
t_primerdia = "01/" & msf1.TextMatrix(0, 1)

End Sub




Private Sub Command1_Click()
F = "01/" & t_fecha & "/" & t_fecha2
t_saldoant = Format$(saldoanterior(F, 0, 0), "######0.00")

End Sub

Private Sub Form_Load()
Call barraesag(Me)
t_fecha = "01"
t_fecha2 = Format$(Now, "yyyy")
t_meses = 12
Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = 0
Call armagrid
Check1 = False
Check2 = False
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F10] Grafica -  [F11] EXCEL"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF7 Then
Dim c(15) As Double
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 5 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
   For i = 0 To Val(t_meses) + 2
    c(i) = i
   Next i
   For i = Val(t_meses) + 3 To 14
      c(i) = -1
    Next i
    
    t1 = "Periodo desde:......: " & t_fecha & " / " & t_fecha2
    T2 = "Meses...............: " & t_meses
    t3 = "Saldo Inicial.......: " & Format$(t_saldoant, "#######0.00")
    
    Call imprimegrid(msf1, c(), "ESTADO DE RESULTADOS", T2, t1, t3, 55, 8, True, False, "H")
  End If
Else
 Call sinpermisos
End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF10 Then
  J = MsgBox("Confirma armado de grafica", 4)
  If J = 6 Then
    Load gen_graficaresultado
    gen_graficaresultado.MSChart1.Title = "Resultado"
    gen_graficaresultado.MSChart1.chartType = VtChChartType2dLine
    gen_graficaresultado.MSChart1.DataGrid.ColumnCount = 1
    gen_graficaresultado.MSChart1.DataGrid.RowCount = msf1.Cols - 3
    gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(1, 1) = "Resultado"
       
    For i = 1 To msf1.Cols - 3
       gen_graficaresultado.MSChart1.DataGrid.RowLabel(i, 1) = msf1.TextMatrix(0, i)
       gen_graficaresultado.MSChart1.DataGrid.SetData i, 1, msf1.TextMatrix(l_r, i), 0
     
    Next i
   gen_graficaresultado.Show
  
  End If
End If



End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If msf1.Row >= gii And msf1.Row <= gfi Then
  h = InputBox$("Ingrese nuevo valor para la celda", "MODIFICACION CASH FLOW", msf1.TextMatrix(msf1.Row, msf1.col))
  msf1.TextMatrix(msf1.Row, msf1.col) = Format$(Val(h), "######0.00")
  Call CALCULATOTALES
 End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
End Sub


Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Val(t_fecha) < 1 Or Val(t_fecha) > 12 Then
    t_fecha = "01"
  End If
Else
   t_fecha = "01"
End If
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Val(t_fecha2) < 2000 Or Val(t_fecha2) > 2100 Then
    t_fecha2 = Format$(Now, "yyyy")
  End If
End If
End Sub




Private Sub t_meses_LostFocus()
If Val(t_meses) < 1 Or Val(t_meses) > 12 Then
  t_meses = 12
End If
End Sub

Private Sub t_saldoant_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_saldoant_LostFocus()
t_saldoant = Format$(Val(t_saldoant), "######0.00")
End Sub

Private Sub UpDown1_DownClick()
If Val(t_fecha) > 1 Then
  t_fecha = Format$(Val(t_fecha) - 1, "00")
Else
 t_fecha = "12"
End If
End Sub

Private Sub UpDown1_UpClick()
If Val(t_fecha) < 12 Then
  t_fecha = Format$(Val(t_fecha) + 1, "00")
Else
 t_fecha = "01"
End If

End Sub

Private Sub UpDown2_DownClick()
  t_fecha2 = Format$(Val(t_fecha2) - 1, "00")
End Sub

Private Sub UpDown2_UpClick()
 t_fecha2 = Format$(Val(t_fecha2) + 1, "00")
End Sub

Private Sub UpDown3_DownClick()
If Val(t_meses) > 1 Then
  t_meses = Val(t_meses) - 1
End If
End Sub

Private Sub UpDown3_UpClick()
If Val(t_meses) < 12 Then
  t_meses = Val(t_meses) + 1
End If

End Sub
Sub CALCULATOTALES()
'CALCULA TOTALES POR CUENTA y Promedio
  If gii >= 2 Then
    For i = 2 To gfi
      tc = 0
      For J = 1 To Val(t_meses)
        tc = tc + Val(msf1.TextMatrix(i, J))
      Next J
      msf1.TextMatrix(i, Val(t_meses) + 1) = Format$(tc, "######0.00")
      msf1.TextMatrix(i, Val(t_meses) + 2) = Format$(tc / Val(t_meses), "######0.00")
    Next i
  End If
  
  
  If gii >= 2 Then
    For i = 1 To Val(t_meses) + 1
      tc = 0
      For J = 2 To gfi
        tc = tc + Val(msf1.TextMatrix(J, i))
      Next J
      msf1.TextMatrix(gfi + 2, i) = Format$(tc, "######0.00")
    Next i
   ' msf1.TextMatrix(gfi, Val(t_meses) + 2) = Format$(Val(msf1.TextMatrix(gfi, Val(t_meses) + 1)) / Val(t_meses), "######0.00")
  End If
End Sub
