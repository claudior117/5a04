VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_gerencial1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME GERENCIAL1"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo"
      Height          =   855
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Width           =   8175
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ambos Informes"
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ingresos y Egresos"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compras y Ventas"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   6015
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10610
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
      Caption         =   "Periodo"
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   3375
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
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
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9960
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta029.frx":0000
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
         Picture         =   "vta029.frx":0882
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
      Top             =   8550
      Width           =   12165
      _ExtentX        =   21458
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_gerencial1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim ttvcc, ttvco, ttccc, ttcco As Double


Sub carga()
  espere.Show
  espere.Refresh
  Call armagrid

  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Informe Gerencial 1 " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 13, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans


If Option5 = True Then
  Call ventas
  Call compras
Else
  If Option6 = True Then
      Call caja2
  Else
     Call ventas
     Call compras
     Call caja2
  End If
 End If
  Unload espere
End Sub
Sub compras()
'BUSCO COMPRAS
  q = "SELECT * FROM A5 where [compra] <> 'N' "
  c = " and "
  
  If t_fecha <> "" Then
    q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
  End If
  
  If t_fecha2 <> "" Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
  End If
  
  q = q & " order by [id_cuenta]"
  
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
   
  tcco = 0
  tccc = 0
  ttcco = 0
  ttccc = 0
  p = 0
  
  If Option6 = False Then
   msf1.AddItem ""
   msf1.AddItem ""
   msf1.AddItem "" & Chr$(9) & "COMPRAS" & Chr$(9) & "Cta.Cte" & Chr$(9) & "Contado" & Chr$(9) & "Total"
  End If
  
  While Not rs.EOF
     If p = 0 Then
        cta = rs("id_cuenta")
        p = 1
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing
     End If
     
     If cta <> rs("id_cuenta") Then
        'muestro
        ttcco = ttcco + tcco
        ttccc = ttccc + tccc
                
                
        If Option6 = False Then
          msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tccc, "#######0.00") & Chr$(9) & Format$(tcco, "#######0.00") & Chr$(9) & Format$(tccc + tcco, "#######0.00")
        End If
               
        tcco = 0
        tccc = 0
        cta = rs("id_cuenta")
        
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing

     Else
       If rs("contado") = "S" Then
          'COMPRAS contado
          If rs("COMPRA") = "S" Then
             'suma COMPRAS contado
              tcco = tcco + rs("total")
          Else
              tcco = tcco - rs("total")
          End If
       Else
          'COMPRAS ctacte
           If rs("compra") = "S" Then
             'suma venta contado
              tccc = tccc + rs("total")
          Else
              tccc = tccc - rs("total")
          End If
       End If
       rs.MoveNext
     End If
     
 Wend

  
   ttcco = ttcco + tcco
   ttccc = ttccc + tccc
        
  If Option6 = False Then
   msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tccc, "#######0.00") & Chr$(9) & Format$(tcco, "#######0.00") & Chr$(9) & Format$(tccc + tcco, "#######0.00")
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & "========================" & Chr(9) & "========================" & Chr(9) & "========================"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & Format$(ttccc, "#######0.00") & Chr$(9) & Format$(ttcco, "#######0.00") & Chr$(9) & Format$(ttccc + ttcco, "#######0.00")
   msf1.AddItem ""
  msf1.AddItem ""
  msf1.AddItem "" & Chr$(9) & "RESULTADO VENTAS-COMPRAS" & Chr(9) & Format$(ttvcc - ttccc, "#######0.00") & Chr$(9) & Format$(ttvco - ttcco, "#######0.00") & Chr$(9) & Format$((ttvcc + ttvco - ttccc - ttcco), "#######0.00")
  msf1.AddItem ""
  msf1.AddItem ""
End If


End Sub
Sub ventas()
 'BUSCO VENTAS
  q = "SELECT * FROM VTA_02 where [venta] <> 'N' "
  c = " and "
  
  If t_fecha <> "" Then
    q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
  End If
  
  If t_fecha2 <> "" Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
  End If
  
  q = q & " order by [id_cuenta]"
  
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
   
  tVco = 0
  tvcc = 0
  ttvco = 0
  ttvcc = 0
  p = 0
  
  If Option6 = False Then
    msf1.AddItem "" & Chr$(9) & "VENTAS" & Chr$(9) & "Cta.Cte" & Chr$(9) & "Contado" & Chr$(9) & "Total"
  End If
  While Not rs.EOF
     If p = 0 Then
        cta = rs("id_cuenta")
        p = 1
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing
     End If
     
     If cta <> rs("id_cuenta") Then
        'muestro
        ttvco = ttvco + tVco
        ttvcc = ttvcc + tvcc
                
        If Option6 = False Then
         msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tvcc, "#######0.00") & Chr$(9) & Format$(tVco, "#######0.00") & Chr$(9) & Format$(tvcc + tVco, "#######0.00")
        End If
               
        tVco = 0
        tvcc = 0
        cta = rs("id_cuenta")
        
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing

     Else
       If rs("contado") = "S" Then
          'venta contado
          If rs("venta") = "S" Then
             'suma venta contado
              tVco = tVco + rs("total")
          Else
              tVco = tVco - rs("total")
          End If
       Else
          'venta ctacte
           If rs("venta") = "S" Then
             'suma venta contado
              tvcc = tvcc + rs("total")
          Else
              tvcc = tvcc - rs("total")
          End If
       End If
       rs.MoveNext
     End If
     
 Wend

  
   ttvco = ttvco + tVco
   ttvcc = ttvcc + tvcc
  If Option6 = False Then
    msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tvcc, "#######0.00") & Chr$(9) & Format$(tVco, "#######0.00") & Chr$(9) & Format$(tvcc + tVco, "#######0.00")
    msf1.AddItem "" & Chr(9) & "" & Chr(9) & "========================" & Chr(9) & "========================" & Chr(9) & "========================"
    msf1.AddItem "" & Chr(9) & "" & Chr(9) & Format$(ttvcc, "#######0.00") & Chr$(9) & Format$(ttvco, "#######0.00") & Chr$(9) & Format$(ttvcc + ttvco, "#######0.00")
  End If
End Sub

Sub caja()
 'BUSCO mov. caja
  q = "SELECT * FROM cyb_05 where [modulo] = 'J'"
  c = " and "
  
  If t_fecha <> "" Then
    q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
  End If
  
  If t_fecha2 <> "" Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
  End If
  
  q = q & " order by [id_cuenta_contra]"
  
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
   
  tic = 0
  tec = 0
  ttic = 0
  ttec = 0
  p = 0
  
  msf1.AddItem "" & Chr$(9) & "INGRESOS - EGRESOS" & Chr$(9) & "Ingresos" & Chr$(9) & "Egresos" & Chr$(9) & "Resultado"
  
  While Not rs.EOF
     If p = 0 Then
        cta = rs("id_cuenta_contra")
        p = 1
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta_contra")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing
     End If
     
     If cta <> rs("id_cuenta_contra") Then
        'muestro
        ttic = ttic + tic
        ttec = ttec + tec
                
                       
        msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tic, "#######0.00") & Chr$(9) & Format$(tec, "#######0.00") & Chr$(9) & Format$(tic - tec, "#######0.00")
               
               
        tic = 0
        tec = 0
        cta = rs("id_cuenta_contra")
        
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta_contra")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing

     Else
       If rs("ubicacion") = "H" Then
              tec = tec + rs("importe")
          Else
              tic = tic + rs("importe")
       End If
       rs.MoveNext
     End If
     
 Wend

  
   ttic = ttic + tic
   ttec = ttec + tec
        
  msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tic, "#######0.00") & Chr$(9) & Format$(tec, "#######0.00") & Chr$(9) & Format$(tic - tec, "#######0.00")
        
  
 'Operaciones Contado contado
  msf1.AddItem "" & Chr$(9) & "Ventas y Compras Contado" & Chr$(9) & Format$(ttvco, "#######0.00") & Chr$(9) & Format$(ttcco, "#######0.00") & Chr$(9) & Format$(ttvco - ttcco, "#######0.00")
    
   ttic = ttic + ttvco
   ttec = ttec + ttcco
  
  
  'busco recibos
  Set rs1 = New ADODB.Recordset
  q = "select * from vta_02 where [id_tipocomp] = 50 "
  c = " and "
  If t_fecha <> "" Then
    q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
  End If
  
  If t_fecha2 <> "" Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
  End If
  
  rs1.Open q, cn1
  tir = 0
  While Not rs1.EOF
    tir = tir + rs1("total")
    rs1.MoveNext
  Wend
  Set rs1 = Nothing
  
  ttic = ttic + tir
  
  
'busco OP
  Set rs1 = New ADODB.Recordset
  q = "select * from a5 where [id_tipocomp] = 50 "
  c = " and "
  If t_fecha <> "" Then
    q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
  End If
  
  If t_fecha2 <> "" Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
  End If
  
  rs1.Open q, cn1
  teop = 0
  While Not rs1.EOF
    teop = teop + rs1("total")
    rs1.MoveNext
  Wend
  Set rs1 = Nothing
  
   ttec = ttec + teop

  msf1.AddItem "" & Chr$(9) & "Rbos. y Pagos Emitidos" & Chr$(9) & Format$(tir, "#######0.00") & Chr$(9) & Format$(teop, "#######0.00") & Chr$(9) & Format$(tir - teop, "#######0.00")
  
  'totales ingresos y egresos
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "========================" & Chr(9) & "========================" & Chr(9) & "========================"
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & Format$(ttic, "#######0.00") & Chr$(9) & Format$(ttec, "#######0.00") & Chr$(9) & Format$(ttic - ttec, "#######0.00")

End Sub


Sub caja2()
  msf1.AddItem "" & Chr$(9) & "INGRESOS - EGRESOS" & Chr$(9) & "Ingresos" & Chr$(9) & "Egresos" & Chr$(9) & "Resultado"
 'BUSCO mov. caja
  'q = "SELECT * FROM cyb_05 where ([modulo] = 'J' or [modulo] = 'V' or [modulo] = 'C')"
  q = "SELECT * FROM cyb_05"
  c = " where "
  
  If t_fecha <> "" Then
    q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
  End If
  
  If t_fecha2 <> "" Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
  End If
  
  q = q & " order by [id_cuenta_contra]"
  
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
   
  tic = 0
  tec = 0
  ttic = 0
  ttec = 0
  p = 0
  
  
  While Not rs.EOF
     If p = 0 Then
        cta = rs("id_cuenta_contra")
        p = 1
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta_contra")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing
     End If
     
     If cta <> rs("id_cuenta_contra") Then
        'muestro
        ttic = ttic + tic
        ttec = ttec + tec
                
                       
        msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tic, "#######0.00") & Chr$(9) & Format$(tec, "#######0.00") & Chr$(9) & Format$(tic - tec, "#######0.00")
               
               
        tic = 0
        tec = 0
        cta = rs("id_cuenta_contra")
        
        Set rs1 = New ADODB.Recordset
        q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta_contra")
        rs1.MaxRecords = 1
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           dc = rs1("descripcion")
        Else
           dc = "Cuenta Inexiastente"
        End If
        Set rs1 = Nothing

     Else
       If rs("ubicacion") = "H" Then
              tec = tec + rs("importe")
          Else
              tic = tic + rs("importe")
       End If
       rs.MoveNext
     End If
     
 Wend

  
 ttic = ttic + tic
 ttec = ttec + tec
        
 msf1.AddItem cta & Chr$(9) & dc & Chr$(9) & Format$(tic, "#######0.00") & Chr$(9) & Format$(tec, "#######0.00") & Chr$(9) & Format$(tic - tec, "#######0.00")
        
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "========================" & Chr(9) & "========================" & Chr(9) & "========================"
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & Format$(ttic, "#######0.00") & Chr$(9) & Format$(ttec, "#######0.00") & Chr$(9) & Format$(ttic - ttec, "#######0.00")
  

End Sub


Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 4500 'cod prov
msf1.ColWidth(2) = 1600
msf1.ColWidth(3) = 1600
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 300

For i = 0 To 1
    msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 5
    msf1.ColAlignment(i) = 9 'der
Next i


End Sub










Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call armagrid
Call barraesag(Me)
Option5 = True

End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F11] Excel "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 5
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "INFORME DE RESULTADOS ", "Periodo del: " & t_fecha & " hasta : " & t_fecha2o, "", "", 75, 8, True, False, "V")
  End If

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If


End Sub


Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  Else
   t_fecha = Format$(t_fecha, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  Else
   t_fecha2 = Format$(t_fecha2, "dd/mm/yyyy")
  End If
End If

End Sub
