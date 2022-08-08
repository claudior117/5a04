VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_ib 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLANILLA INGRESOS BRUTOS"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   6615
      Begin VB.OptionButton Option3 
         Caption         =   "Detallado x Comprobante"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Detallado x Tasa"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Width           =   6255
      Begin VB.ComboBox c_acti 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Concepto IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   10616833
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
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "VTA015.frx":0000
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
         Picture         =   "VTA015.frx":0882
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
      Top             =   8475
      Width           =   12045
      _ExtentX        =   21246
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
            TextSave        =   "08/08/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:09 p.m."
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
Attribute VB_Name = "vta_ib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Sub carga()
 Call armagrid
    
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Informe Ingresos Brutos " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 14, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
  
  
  'para cada actividad
  q = "select * from g8"
  If c_acti.ListIndex > 0 Then
    q = q & " where [id_actividad] = " & c_acti.ItemData(c_acti.ListIndex)
  End If
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tr = 0
  ti = 0
  While Not rs.EOF
    'busco comprobantes para cada actividad ordenados por tasa ib
    q = "select *  from VTA_02, vta_06, VTA_03, a2 where  vta_02.[num_int] = vta_03.[num_int] and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and [ib] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]  and a2.[id_producto] = vta_03.[id_producto] and [id_actividad] = " & rs("id_actividad")
    c = " and "
  
     
    If IsDate(t_fecha) Then
      q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    End If
  
    If IsDate(t_fecha2) Then
      q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
  
    If c_vend.ListIndex > 0 Then
      q = q & c & " [id_tasaib] = " & c_vend.ItemData(c_vend.ListIndex)
    End If
    
    q = q & " order by [id_tasaib], [fecha], [letra], [num_comp]"
  
    Set rs1 = New ADODB.Recordset
    rs1.Open q, cn1

    p = 0
    While Not rs1.EOF
     If p = 0 Then
        
       a = rs1("id_tasaib")
       Set rs2 = New ADODB.Recordset
       q = "select * from g12 where [id_tasaib] = " & a
       rs2.Open q, cn1
       If Not rs2.EOF And Not rs2.BOF Then
         d = rs2("descripcionib")
         t = rs2("tasaib")
       Else
         d = "Error"
         t = "0.0"
       End If
       Set rs2 = Nothing
       l1 = "======================================"
       msf1.AddItem ""
       msf1.AddItem "" & Chr$(9) & rs("descripcion") & Chr(9) & "" & Chr$(9) & d
       t1 = ""
       For i = 0 To 8
         t1 = t1 & l1 & Chr$(9)
       Next i
       msf1.AddItem t1
       p = 1
       r = 0
       s = 0
       ib = 0
     End If
     If a <> rs1("id_tasaib") Then
       't = (s * a / 100)
       ts = ts + s
       ti = ti + ib
       p = 0
       msf1.AddItem "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr$(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
       msf1.AddItem "TOTALES" & Chr(9) & rs("descripcion") & Chr$(9) & " " & Chr(9) & d & Chr$(9) & Format$(s, "######0.00") & Chr(9) & "" & Chr(9) & Format$(ib, "######0.00")
     Else
       If rs1("vta_02.moneda") = "P" Then
         c5 = 1
       Else
         c5 = rs1("cotizacion_dolar")
       End If
       'If rs1("vta_02.id_tipocomp") <> 100 Then 'retencion de ib
        n = (rs1("importe") * c5)
        i = Format((n * rs1("tasaib") / 100), "#####0.00")
        If rs1("ib") = "S" Then
         s = s + n
         ib = ib + i
        Else
         s = s - n
         ib = ib - i
        End If
        If n <> 0 Then
         msf1.AddItem rs1("fecha") & Chr$(9) & rs1("cliente02") & Chr(9) & rs1("vta_06.abreviatura") & " " & rs1("letra") & Format$(rs1("vta_02.sucursal"), "0000") & "-" & Format$(rs1("num_comp"), "00000000") & Chr(9) & rs("descripcion") & Chr$(9) & Format$(n, "######0.00") & Chr(9) & Format$(rs1("tasaib"), "0.00") & Chr(9) & Format$(i, "####0.00") & Chr(9) & "" & Chr$(9) & rs1("vta_02.num_int") & Chr$(9) & rs1("renglon")
        End If
       rs1.MoveNext
      End If
     Wend 'rs1
     Set rs1 = Nothing
     If p <> 0 Then
      r = sacaretib(rs("id_actividad"))
      ts = ts + s
      ti = ti + ib
      msf1.AddItem "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr$(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
      msf1.AddItem "TOTALES" & Chr(9) & rs("descripcion") & Chr$(9) & " " & Chr(9) & d & Chr$(9) & Format$(s, "######0.00") & Chr(9) & "" & Chr(9) & Format$(ib, "######0.00")
      tr = tr + r
     End If
     rs.MoveNext
   Wend
   msf1.AddItem ""
   msf1.AddItem "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr$(9) & "______________________" & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________"
   msf1.AddItem "TOTALES" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "" & Chr$(9) & Format$(ts, "######0.00") & Chr(9) & "" & Chr$(9) & Format$(ti, "######0.00") & Chr(9) & Format$(tr, "######0.00")
      
  
  Set rs = Nothing
   
   
   
End Sub

Sub carga3()
 Call armagrid
    
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Informe Ingresos Brutos " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 14, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
  
  
    'busco comprobantes
    q = "select * from VTA_02, vta_06, VTA_03, a2 where  vta_02.[num_int] = vta_03.[num_int] and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and [ib] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]  and a2.[id_producto] = vta_03.[id_producto] "
    c = " and "
  
     
    If IsDate(t_fecha) Then
      q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    End If
  
    If IsDate(t_fecha2) Then
      q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
  
    If c_vend.ListIndex > 0 Then
      q = q & c & " [id_tasaib] = " & c_vend.ItemData(c_vend.ListIndex)
    End If
    
    q = q & " order by  [fecha], [letra], [num_comp]"
  
    Set rs1 = New ADODB.Recordset
    rs1.Open q, cn1

    p = 0
    
    While Not rs1.EOF
     If p <> rs1("vta_02.num_int") Then
       'muestro cabecera factura
       If p <> 0 Then
         'muestro total de comprobante
         msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "-------------------" & Chr$(9) & "" & Chr$(9) & "------------------------"
         msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & Format$(s, "######0.00") & Chr$(9) & "" & Chr$(9) & Format$(ib, "######0.00")
          msf1.AddItem ""
       End If
         
       If rs1("vta_02.moneda") = "P" Then
         c5 = 1
       Else
         c5 = rs1("cotizacion_dolar")
       End If
        n = (rs1("importe") * c5)
        i = Format((n * rs1("tasaib") / 100), "#####0.00")
        
         s = 0
         ib = 0
        
       msf1.AddItem rs1("fecha") & Chr$(9) & rs1("cliente02") & Chr(9) & rs1("vta_06.abreviatura") & " " & rs1("letra") & Format$(rs1("vta_02.sucursal"), "0000") & "-" & Format$(rs1("num_comp"), "00000000") & Chr(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & rs1("vta_02.num_int") & Chr$(9) & rs1("renglon")
       
       p = rs1("vta_02.num_int")
     Else
       'muestro renglon productos
       If rs1("vta_02.moneda") = "P" Then
         c5 = 1
       Else
         c5 = rs1("cotizacion_dolar")
       End If
 
       s = s + (rs1("importe") * c5)
       ib = ib + Val(Format$(rs1("tasaib") * (rs1("importe") * c5) / 100, "####0.00"))
       msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & rs1("vta_03.descripcion") & " (" & rs1("cantidad_original") & rs1("tunidad") & ")" & Chr$(9) & Format$((rs1("importe") * c5), "#####0.00") & Chr$(9) & rs1("tasaib") & Chr$(9) & Format$(rs1("tasaib") * (rs1("importe") * c5) / 100, "####0.00")
       rs1.MoveNext
     End If
     
     Wend 'rs1
     Set rs1 = Nothing
     
     'If p <> 0 Then
     ' r = sacaretib(rs("id_actividad"))
     ' ts = ts + s
     ' ti = ti + ib
     ' msf1.AddItem "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr$(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
     ' msf1.AddItem "TOTALES" & Chr(9) & rs("descripcion") & Chr$(9) & " " & Chr(9) & d & Chr$(9) & Format$(s, "######0.00") & Chr(9) & "" & Chr(9) & Format$(ib, "######0.00")
     ' tr = tr + r
     'End If
     'rs.MoveNext
   'Wend
  ' msf1.AddItem ""
  ' msf1.AddItem "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr$(9) & "______________________" & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________"
  ' msf1.AddItem "TOTALES" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "" & Chr$(9) & Format$(ts, "######0.00") & Chr(9) & "" & Chr$(9) & Format$(ti, "######0.00") & Chr(9) & Format$(tr, "######0.00")
      
  
 ' Set rs = Nothing
   
   
   
End Sub
Sub carga2()
  Call armagrid2
    
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Informe Ingresos Brutos " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 14, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
  
  
  'para cada actividad
  q = "select * from g8"
  If c_acti.ListIndex > 0 Then
    q = q & " where [id_actividad] = " & c_acti.ItemData(c_acti.ListIndex)
  End If
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tr = 0
  ti = 0
  trb = 0
  While Not rs.EOF
    'busco comprobantes para cada actividad ordenados por tasa ib
    q = "select [id_tasaib], [importe], vta_02.[moneda], [cotizacion_dolar], [tasaib], [ib]  from VTA_02, vta_06, VTA_03, a2 where  vta_02.[num_int] = vta_03.[num_int] and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and [ib] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]  and a2.[id_producto] = vta_03.[id_producto] and [id_actividad] = " & rs("id_actividad")
    c = " and "
  
     
    If IsDate(t_fecha) Then
      q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    End If
  
    If IsDate(t_fecha2) Then
      q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
  
    If c_vend.ListIndex > 0 Then
      q = q & c & " [id_tasaib] = " & c_vend.ItemData(c_vend.ListIndex)
    End If
    
    q = q & " order by [id_tasaib]"
  
    Set rs1 = New ADODB.Recordset
    rs1.Open q, cn1

    p = 0
    While Not rs1.EOF
     If p = 0 Then
       a = rs1("id_tasaib")
       p = 1
       r = 0
       s = 0
       ib = 0
       rb = 0
     End If
     If a <> rs1("id_tasaib") Then
       't = (s * a / 100)
       Set rs2 = New ADODB.Recordset
       q = "select * from g12 where [id_tasaib] = " & a
       rs2.Open q, cn1
       If Not rs2.EOF And Not rs2.BOF Then
         d = rs2("descripcionib")
         t = rs2("tasaib")
       Else
         d = "Error"
         t = "0.0"
       End If
       Set rs2 = Nothing
       ts = ts + s
       ti = ti + ib
       msf1.AddItem rs("descripcion") & Chr(9) & d & Chr(9) & Format$(t, "0.00") & Chr(9) & Format$(s, "######0.00") & Chr(9) & Format$(ib, "####0.00")
       msf1.AddItem ""
       a = rs1("id_tasaib")
       s = 0
       ib = 0
     Else
       If rs1("moneda") = "P" Then
         c5 = 1
       Else
         c5 = rs1("cotizacion_dolar")
       End If
       'If rs1("vta_02.id_tipocomp") <> 100 Then 'retencion de ib
        n = (rs1("importe") * c5)
        i = Format((n * rs1("tasaib") / 100), "#####0.00")
        If rs1("ib") = "S" Then
         s = s + n
         ib = ib + i
        Else
         s = s - n
         ib = ib - i
        End If
       
        'End If
        rs1.MoveNext
      End If
     Wend 'rs1
     Set rs1 = Nothing
     If p <> 0 Then
      r = sacaretib(rs("id_actividad"))
      rb = sacaretibbanco(rs("id_actividad"))

      ts = ts + s
      ti = ti + ib
      Set rs2 = New ADODB.Recordset
      q = "select * from g12 where [id_tasaib] = " & a
      rs2.Open q, cn1
      If Not rs2.EOF And Not rs2.BOF Then
         d = rs2("descripcionib")
         t = rs2("tasaib")
      Else
         d = "Error"
         t = "0.0"
      End If
      Set rs2 = Nothing
       
      msf1.AddItem rs("descripcion") & Chr(9) & d & Chr(9) & Format$(t, "0.00") & Chr(9) & Format$(s, "######0.00") & Chr(9) & Format$(ib, "####0.00")
      tr = tr + r
      trb = trb + rb
     End If
     rs.MoveNext
   Wend
   tp = buscaperc
   msf1.AddItem "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________"
   msf1.AddItem "TOTALES" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & Format$(ts, "######0.00") & Chr(9) & Format$(ti, "######0.00")
   msf1.AddItem "" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "Ret.IB........" & Chr(9) & Format$(tr, "######0.00")
   msf1.AddItem "" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "Ret.IB Banco.." & Chr(9) & Format$(trb, "######0.00")
   msf1.AddItem "" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "Perc.IB......." & Chr(9) & Format$(tp, "######0.00")
   
   msf1.AddItem "" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "" & Chr(9) & "====================="
   msf1.AddItem "" & Chr(9) & "" & Chr$(9) & " " & Chr(9) & "A pagar " & Chr(9) & Format$(ti - tr - trb - tp, "######0.00")
   
   
      
  
  Set rs = Nothing
   
   
   
End Sub

Function buscaperc() As Double

Set rs1 = New ADODB.Recordset
q = "select * from a12 where [tipo12] = 'P' and [impuesto12] = 'B'"
rs1.Open q, cn1
totperc = 0
dr = "IB"
tp = 0
While Not rs1.EOF
 
 p = 0
 q = "select * from a5, g2, g3, a1, a13  where  [GRABADO] <> 'N' AND [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor]  AND a5.[num_int] = a13.[num_int] and a13.[id_percepcion] = " & rs1("id_percepcion") & " and a1.[cod_tipoiva] = g3.[cod_tipoiva]"
 c = " and "
 If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
 End If
 If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
 End If
 q = q & " order by [fecha]"
 'MsgBox (q)
 Set rs = New ADODB.Recordset
 rs.Open q, cn1
 tt = 0
 While Not rs.EOF
     If rs("moneda") = "P" Then
        c5 = 1
     Else
        c5 = rs("cotiz_dolar")
     End If
     
     If rs("grabado") = "S" Then
       t = Format$(rs("importe") * c5, "######0.00")
     Else
        t = Format$(-rs("importe") * c5, "######0.00")
     End If
     tt = tt + Val(t)
     rs.MoveNext
  Wend
  tp = tp + tt
  Set rs = Nothing
  rs1.MoveNext
 Wend
 buscaperc = tp
End Function


Function sacaretib(ByVal ia As Long) As Double
 '[total_bultos] tiene el codigo de regimen
 'separo 1 que sn ret comunes y 2 que son ret bancarias
q = "select * from VTA_02, vta_06 where  vta_02.[id_tipocomp] = vta_06.[id_tipocomp]  and [ib] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal] and vta_02.[id_actividad] = " & ia & " and vta_02.[id_tipocomp] = 100 and [total_bultos] = 1"
c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  Set rs1 = New ADODB.Recordset
  rs1.Open q, cn1
  r = 0
  While Not rs1.EOF
    If rs1("vta_02.moneda") = "P" Then
         c5 = 1
    Else
        c5 = rs1("cotizacion_dolar")
    End If
    
   
    
    If rs1("ib") = "S" Then
       r = r + (rs1("total") * c5)
    Else
       r = r - (rs1("total") * c5)
    End If
   rs1.MoveNext
  Wend
  Set rs1 = Nothing
  sacaretib = r
End Function

Function sacaretibbanco(ByVal ia As Long) As Double
 '[total_bultos] tiene el codigo de regimen
 'separo 1 que sn ret comunes y 2 que son ret bancarias
q = "select * from VTA_02, vta_06 where  vta_02.[id_tipocomp] = vta_06.[id_tipocomp]  and [ib] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal] and vta_02.[id_actividad] = " & ia & " and vta_02.[id_tipocomp] = 100 and [total_bultos] = 2"
c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  Set rs1 = New ADODB.Recordset
  rs1.Open q, cn1
  r = 0
  While Not rs1.EOF
    If rs1("vta_02.moneda") = "P" Then
         c5 = 1
    Else
        c5 = rs1("cotizacion_dolar")
    End If
    
   
    
    If rs1("ib") = "S" Then
       r = r + (rs1("total") * c5)
    Else
       r = r - (rs1("total") * c5)
    End If
   rs1.MoveNext
  Wend
  Set rs1 = Nothing
  sacaretibbanco = r
End Function
Private Sub btnacepta_Click()
If Option1 = True Then
  Call carga2
Else
 If Option2 = True Then
   Call carga
 Else
   Call carga3
 End If
End If

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub







Private Sub c_acti_LostFocus()
If c_acti.ListIndex < 0 Then
  c_acti.ListIndex = 0
End If

End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If

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
     Unload Me
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 2500
msf1.ColWidth(2) = 1800
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 800
msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Cliente"
msf1.TextMatrix(0, 2) = "Comprobante"
msf1.TextMatrix(0, 3) = "Concepto"
msf1.TextMatrix(0, 4) = "Imponible"
msf1.TextMatrix(0, 5) = "Tasa IB"
msf1.TextMatrix(0, 6) = "Imp. IB"
msf1.TextMatrix(0, 7) = "Retenciones"
msf1.TextMatrix(0, 8) = "Num.Int."
msf1.TextMatrix(0, 9) = "reng."
For i = 0 To 3
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 4 To 7
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Sub armagrid2()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 3500
msf1.ColWidth(1) = 2500
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1500
msf1.ColWidth(4) = 1500
msf1.ColWidth(5) = 1500

msf1.TextMatrix(0, 0) = "Actividad"
msf1.TextMatrix(0, 1) = "Concepto"
msf1.TextMatrix(0, 2) = "% IB Actual"
msf1.TextMatrix(0, 3) = "Imponible"
msf1.TextMatrix(0, 4) = "Imp. IB"
msf1.TextMatrix(0, 5) = ""

For i = 1 To 4
  msf1.ColAlignment(i) = 9 'der
Next i
  msf1.ColAlignment(0) = 1 'izq

End Sub

Private Sub Form_Load()

Call barraesag(Me)
cal1.Visible = False
Call armagrid2

Call carga_actividades(c_acti)
c_acti.AddItem "<Todas>", 0
c_acti.ListIndex = 0


Call carga_tasaib(c_vend)
c_vend.AddItem "<Todas>", 0
c_vend.ListIndex = 0

Option1 = True
End Sub



Private Sub msf1_GotFocus()

Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F3] Actualiza tasa IB "

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
    
    For i = 5 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LIQUIDACION INGRESOS BRUTOS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If

If KeyCode = vbKeyF3 Then
 If Option2 Then
  ni = Val(msf1.TextMatrix(msf1.Row, 8))
  r = Val(msf1.TextMatrix(msf1.Row, 9))
  If ni > 0 And r > 0 Then
    J = MsgBox("Actualiza tasa IB para Item", 4)
    If J = 6 Then
      Set rs = New ADODB.Recordset
      q = "select g12.[tasaib], vta_03.[tasaib] from vta_03, a2, g12 where [num_int] = " & ni & " and [renglon] = " & r & " and vta_03.[id_producto] = a2.[id_producto] and a2.[id_tasaib] = g12.[id_tasaib]"
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.EOF And Not rs.BOF Then
        tib = rs("g12.tasaib")
        rs("vta_03.tasaib") = tib
        rs.Update
      End If
      Set rs = Nothing
    
    End If
 End If
End If
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 And Option2 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
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
