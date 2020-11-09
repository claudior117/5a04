VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cja_detallemov 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME de CAJA"
   ClientHeight    =   8760
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Importe"
      Height          =   615
      Left            =   5040
      TabIndex        =   30
      Top             =   7440
      Width           =   4455
      Begin VB.TextBox t_importe 
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Caption         =   "< ="
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "> ="
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "="
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3000
      TabIndex        =   28
      Top             =   7920
      Width           =   1815
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin saldo inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agrupado por:"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Concepto"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Agrupar"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtros"
      Height          =   1695
      Left            =   5160
      TabIndex        =   7
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   5040
         Picture         =   "cja_002.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox c_usuario 
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox C_concepto 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   3975
      End
      Begin VB.ComboBox C_tipo 
         Height          =   315
         ItemData        =   "cja_002.frx":030A
         Left            =   960
         List            =   "cja_002.frx":0317
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox C_subrubro 
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Concepto:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo"
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox c_n1 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox T_detalle 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox T_FECHA2 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Titulo"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalle:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Hasta:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
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
         Picture         =   "cja_002.frx":0337
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
         Picture         =   "cja_002.frx":0BB9
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
      Top             =   8505
      Width           =   12105
      _ExtentX        =   21352
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
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cja_detallemov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim o As Integer

Private Sub btnacepta_Click()
Call limpia
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_concepto_LostFocus()
If c_concepto.ListIndex < 0 Then
  c_concepto.ListIndex = 0
End If
End Sub


Private Sub C_subrubro_LostFocus()
If C_subrubro.ListIndex < 0 Then
  If Val(C_subrubro) > 0 Then
    C_subrubro.ListIndex = buscaindice(C_subrubro, Val(C_subrubro))
  Else
    C_subrubro.ListIndex = 0
  End If
End If

End Sub


Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
End Sub

Sub limpia()
  Call armagrid
  espere.Show
  espere.Label1 = "ESPERE........  [Generando Informe]"
  espere.Refresh
  o = 1
  If Option1 = True Then
    o = 1
    msf1.Tag = Option1.Caption
  End If
  
   
  If Option3 = True Then
    o = 3
    msf1.Tag = Option3.Caption
  End If

  If Option4 = True Then
    o = 4
    msf1.Tag = Option4.Caption
  End If

 
  
  Select Case o
   Case Is = 1
     Call opcion1
   Case Is = 3
     Call opcion3
   Case Is = 4
     Call opcion4
  
  End Select
  Unload espere
  End Sub

Sub opcion1()
q = "SELECT * FROM cyb_05, C_01, cyb_01 WHERE  [ID_cuenta_contra] = c_01.[ID_cuenta] AND cyb_05.[ID_forma_pago] = cyb_01.[ID_forma_pago] "
       
       
  If C_subrubro.ListIndex > 0 Then
         q = q & " and [id_cuenta_contra] = " & C_subrubro.ItemData(C_subrubro.ListIndex)
  End If
       
  If c_n1.ListIndex > 0 Then
    Set rs2 = New ADODB.Recordset
    k = "select * from c_01 where [id_cuenta] = " & c_n1.ItemData(c_n1.ListIndex)
    rs2.Open k, cn1
    If Not rs2.EOF And Not rs2.BOF Then
      i = Format$(rs2("pos1"), "0")
      F = Format$(rs2("pos1"), "0")
      If rs2("pos2") > 0 Then
         i = i & Format$(rs2("pos2"), "0")
         F = F & Format$(rs2("pos2"), "0")
         If rs2("pos3") > 0 Then
           i = i & Format$(rs2("pos3"), "00")
           F = F & Format$(rs2("pos3"), "00")
         Else
           i = i & "00"
           F = F & "99"
         End If
      Else
         i = i & "000"
         F = F & "999"
      End If
    End If
    i = i & "00"
    F = F & "99"
    Set rs2 = Nothing
        
    q = q & " and [id_cuenta_contra] > " & i & " and [id_cuenta_contra] <= " & F
  
  End If
  
  If c_concepto.ListIndex > 0 Then
      q = q & " and cyb_05.[id_forma_pago] = " & c_concepto.ItemData(c_concepto.ListIndex)
  End If
       
  If t_detalle <> "" Then
     q = q & " and cyb_05.[descripcion] like '%" & t_detalle & "%'"
  End If
       
   If Mid$(c_tipo, 1, 1) = "I" Then
     q = q & " and [ubicacion] = 'D'"
   Else
     If Mid$(c_tipo, 1, 1) = "E" Then
        q = q & " and [ubicacion] = 'H'"
     End If
   End If
       
       
   If c_usuario.ListIndex > 0 Then
         q = q & " and [id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
   End If
       
  
  If t_fecha <> "" Then
    If IsDate(t_fecha) Then
       qa = q & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
       q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    Else
       sa = 0
       qa = ""
    End If
  Else
    sa = 0
    qa = ""
  End If
  
  If t_fecha2 <> "" Then
    If IsDate(t_fecha2) Then
       q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
  End If
  
  If t_importe <> "" Then
    If Option2 = True Then
      q1 = " = "
    Else
      If Option5 = True Then
        q1 = " >= "
      Else
        q1 = " <= "
      End If
    End If
    q = q & " and cyb_05.[importe] " & q1 & Val(t_importe)
  End If
    
    
    
  q = q & " order by fecha"
  
  Set rs = New ADODB.Recordset
  
    
  rs.Open q, cn1
  
  tci = 0
  tce = 0
  pasada = 0
  fi = Format$(Now, "dd/mm/yyyy")
  ff = Format$(Now, "dd/mm/yyyy")
  
  While Not rs.EOF
    If pasada = 0 Then
      fi = rs("fecha")
      ff = rs("Fecha")
      pasada = 1
    Else
      ff = rs("Fecha")
    End If
    d = rs("cyb_05.descripcion")
    r = rs("operacion")
    s = rs("c_01.descripcion")
    c = rs("cyb_01.descripcion")
    F = Format$(rs("fecha"), "dd/mm/yyyy")
    If rs("ubicacion") = "D" Then
       i = Format$(rs("cyb_05.IMPORTE"), "######0.00")
       e = ""
       tci = tci + rs("cyb_05.IMPORTE")
    Else
       e = Format$(rs("cyb_05.IMPORTE"), "######0.00")
       i = ""
       tce = tce + rs("cyb_05.IMPORTE")
    End If
    msf1.AddItem F & Chr$(9) & r & Chr$(9) & i & Chr$(9) & e & Chr$(9) & s & Chr$(9) & c & Chr$(9) & d & Chr$(9) & rs("num_mov_caja")
    rs.MoveNext
  Wend
  Set rs = Nothing
  
  
  'calcula saldo anterior
  If Check1 = 0 Then
   If qa <> "" Then
    Set rs = New ADODB.Recordset
    rs.Open qa, cn1
    sa = 0
    While Not rs.EOF
     If rs("ubicacion") = "D" Then
       sa = sa + rs("cyb_05.IMPORTE")
     Else
       sa = sa - rs("cyb_05.IMPORTE")
     End If
     rs.MoveNext
    Wend
    Set rs = Nothing
   End If
  Else
   sa = 0
 End If
  
  l1 = "================================"
  'msf1.AddItem ""  & Chr$(9) & "" & Chr$(9) & "--------------------------------" & Chr$(9) & "----------------------------"
  'msf1.AddItem " Saldo Ant.  " & Chr$(9) & Format$(sa, "#######0.00") & "      " & Chr$(9) & Format$(tci, "######0.00") & Chr$(9) & Format$(tce, "######0.00") & Chr$(9) & "Saldo Act. ==> " & Chr$(9) & Format$(sa + tci - tce, "######0.00")
  
  msf1.AddItem l1 & Chr$(9) & l1 & Chr$(9) & l1 & Chr$(9) & l1 & Chr$(9) & l1 & Chr$(9) & l1 & Chr$(9) & l1 & Chr$(9) & l1 & Chr$(9) & l1
  msf1.AddItem ""
  msf1.AddItem ""
  msf1.AddItem "" & Chr$(9) & " Saldo Anterior " & Chr$(9) & Format$(sa, "#######0.00")
  msf1.AddItem "" & Chr$(9) & " INGRESOS (+)   " & Chr$(9) & Format$(tci, "######0.00")
  msf1.AddItem "" & Chr$(9) & " EGRESOS  (-)   " & Chr$(9) & Format$(tce, "######0.00")
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "-----------------------"
  msf1.AddItem "" & Chr$(9) & "Saldo Actual    " & Chr$(9) & Format$(sa + tci - tce, "######0.00")
  msf1.AddItem ""
  'd = DateDiff("d", fi, ff)
  'If d <= 0 Then
  '  d = 1
  'End If
  'pe = Format$(tce / d, "#####0.00")
  'pi = Format$(tci / d, "#####0.00")
  'ps = Format$((tci - tce) / d, "#####0.00")
  'msf1.AddItem "" & Chr$(9) & "Dias Evaluados: " & d
  'msf1.AddItem ""
  'msf1.AddItem "" & Chr$(9) & "Promedio por dia: " & Chr$(9) & pi & Chr$(9) & pe & Chr$(9) & ps
  msf1.SetFocus
  Set rs = Nothing
  
End Sub


Sub opcion3()


If c_n1.ListIndex > 0 Then
    Set rs2 = New ADODB.Recordset
    k = "select * from c_01 where [id_cuenta] = " & c_n1.ItemData(c_n1.ListIndex)
    rs2.Open k, cn1
    If Not rs2.EOF And Not rs2.BOF Then
      i = Format$(rs2("pos1"), "0")
      F = Format$(rs2("pos1"), "0")
      If rs2("pos2") > 0 Then
         i = i & Format$(rs2("pos2"), "0")
         F = F & Format$(rs2("pos2"), "0")
         If rs2("pos3") > 0 Then
           i = i & Format$(rs2("pos3"), "00")
           F = F & Format$(rs2("pos3"), "00")
         Else
           i = i & "00"
           F = F & "99"
         End If
      Else
         i = i & "000"
         F = F & "999"
      End If
    End If
    i = i & "00"
    F = F & "99"
    Set rs2 = Nothing
        
     
  End If

q = "select * from c_01"
c = " where "
If c_n1.ListIndex > 0 Then
   q = q & c & " [id_cuenta] > " & i & " and [id_cuenta] <= " & F
End If
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
tti = 0
tte = 0
fi = Format$(Now, "dd/mm/yyyy")
ff = Format$(Now, "dd/mm/yyyy")



While Not rs2.EOF
  q = "SELECT * FROM cyb_05  "
  c = " where "
       
  If C_subrubro.ListIndex > 0 Then
         q = q & c & " [id_cuenta_contra] = " & C_subrubro.ItemData(C_subrubro.ListIndex)
         c = " and "
  End If
       
  
  If t_fecha2 <> "" Then
    If IsDate(t_fecha2) Then
       q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
       c = " and "
    End If
  End If
       
  If c_concepto.ListIndex > 0 Then
      q = q & c & " [id_forma_pago] = " & c_concepto.ItemData(c_concepto.ListIndex)
      c = " and "
  End If
       
  If t_detalle <> "" Then
     q = q & c & " [descripcion] like '%" & t_detalle & "%'"
     c = " and "
  End If
       
   If Mid$(c_tipo, 1, 1) = "I" Then
     q = q & c & " [ubicacion] = 'D'"
     c = " and "
   Else
     If Mid$(c_tipo, 1, 1) = "E" Then
        q = q & c & " [ubicacion] = 'H'"
        c = " and "
     End If
   End If
       
   If c_usuario.ListIndex > 0 Then
         q = q & c & " [id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
         c = " and "
   End If
    
       
  If t_fecha <> "" Then
    If IsDate(t_fecha) Then
       qa = q & c & " datevalue([fecha]) < datevalue('" & t_fecha & "')"
       q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
        c = " and "
      Else
       qa = ""
    End If
  Else
    qa = ""
  End If
  
  If t_fecha2 <> "" Then
    If IsDate(t_fecha2) Then
       q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
       c = " and "
    End If
  End If
  
  q = q & c & " [id_cuenta_contra] = " & rs2("id_cuenta")
    
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tci = 0
  tce = 0
  p = 0
  While Not rs.EOF
    If p = 0 Then
      d = rs2("descripcion")
      r = ""
      s = rs2("descripcion")
      c = ""
      p = 1
      pe = ""
    End If
    
    If rs("ubicacion") = "D" Then
       tci = tci + rs("IMPORTE")
    Else
       tce = tce + rs("IMPORTE")
    End If
    
    If DateValue(fi) > DateValue(rs("fecha")) Then
       fi = Format$(rs("fecha"), "dd/mm/yyyy")
    End If
    
    If DateValue(ff) < DateValue(rs("fecha")) Then
       ff = Format$(rs("fecha"), "dd/mm/yyyy")
    End If
    rs.MoveNext
  Wend
  Set rs = Nothing
  
   If tci > 0 Or tce > 0 Then
    tti = tti + tci
    tte = tte + tce
    
        F = ""
    i = Format$(tci, "######0.00")
    e = Format$(tce, "######0.00")
    msf1.AddItem F & Chr$(9) & r & Chr$(9) & i & Chr$(9) & e & Chr$(9) & s & Chr$(9) & c & Chr$(9) & "" & Chr$(9) & ""
   End If
  rs2.MoveNext
  Wend
  Set rs2 = Nothing
  
   'calcula saldo anterior
If Check1 = 0 Then
  If qa <> "" Then
    Set rs = New ADODB.Recordset
    rs.Open qa, cn1
    sa = 0
    While Not rs.EOF
     If rs("ubicacion") = "D" Then
       sa = sa + rs("IMPORTE")
     Else
       sa = sa - rs("IMPORTE")
     End If
     rs.MoveNext
    Wend
    Set rs = Nothing
  Else
    sa = 0
  End If
 Else
  sa = 0
 End If
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "--------------------------------" & Chr$(9) & "----------------------------"
  msf1.AddItem "Saldo Ant." & Chr$(9) & Format$(sa, "#######0.00") & " <====" & Chr$(9) & Format$(tti, "######0.00") & Chr$(9) & Format$(tte, "######0.00") & Chr$(9) & "==> " & Format$(sa + tti - tte, "######0.00")
  msf1.AddItem ""
  d = DateDiff("d", fi, ff)
  If d <= 0 Then
    d = 1
  End If
  pe = Format$(tte / d, "#####0.00")
  pi = Format$(tti / d, "#####0.00")
  ps = Format$((tti - tte) / d, "#####0.00")
  msf1.AddItem "" & Chr$(9) & "Dias Evaluados: " & d
  msf1.AddItem ""
  msf1.AddItem "" & Chr$(9) & "Promedio por dia: " & Chr$(9) & pi & Chr$(9) & pe & Chr$(9) & ps

  msf1.SetFocus
  Set rs = Nothing
  
  
  
End Sub


Sub opcion4()
q = "select * from cyb_01"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
tti = 0
tte = 0
While Not rs2.EOF
  q = "SELECT * FROM cyb_05, c_01 WHERE  [ID_cuenta_contra] = c_01.[ID_cuenta] "
       
       
  If C_subrubro.ListIndex > 0 Then
         q = q & " and [id_cuenta_contra] = " & C_subrubro.ItemData(C_subrubro.ListIndex)
  End If
       
  If c_n1.ListIndex > 0 Then
    Set rs = New ADODB.Recordset
    k = "select * from c_01 where [id_cuenta] = " & c_n1.ItemData(c_n1.ListIndex)
    rs.Open k, cn1
    If Not rs.EOF And Not rs.BOF Then
      i = Format$(rs("pos1"), "0")
      F = Format$(rs("pos1"), "0")
      If rs("pos2") > 0 Then
         i = i & Format$(rs("pos2"), "0")
         F = F & Format$(rs("pos2"), "0")
         If rs("pos3") > 0 Then
           i = i & Format$(rs("pos3"), "00")
           F = F & Format$(rs("pos3"), "00")
         Else
           i = i & "00"
           F = F & "99"
         End If
      Else
         i = i & "000"
         F = F & "999"
      End If
    End If
    i = i & "00"
    F = F & "99"
    Set rs = Nothing
        
    q = q & " and [id_cuenta_contra] > " & i & " and [id_cuenta_contra] <= " & F
  
  End If
       
       
       
         
  If c_concepto.ListIndex > 0 Then
      q = q & " and cyb_05.[id_forma_pago] = " & c_concepto.ItemData(c_concepto.ListIndex)
  End If
       
  If t_detalle <> "" Then
     q = q & " and cyb_05.[descripcion] like '%" & t_detalle & "%'"
  End If
       
  If c_tipo.ListIndex > 0 Then
   If Mid$(c_tipo, 1, 1) <> "I" Then
     q = q & " and [ubicacion] = 'D'"
   Else
     If Mid$(c_tipo, 1, 1) <> "E" Then
        q = q & " and [ubicacion] = 'H'"
     End If
   End If
  End If
  
       
   If c_usuario.ListIndex > 0 Then
         q = q & " and [id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
   End If
   
  If t_fecha <> "" Then
    If IsDate(t_fecha) Then
       qa = q & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
       q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    End If
  End If
  
  If t_fecha2 <> "" Then
    If IsDate(t_fecha2) Then
       q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
  End If
 
  q = q & " and [id_forma_pago] = " & rs2("id_forma_pago")
       
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tci = 0
  tce = 0
  p = 0
  While Not rs.EOF
    If p = 0 Then
      d = rs2("descripcion")
      r = ""
      s = ""
      c = rs2("descripcion")
      pe = ""
      p = 1
    End If
    
    
    
    If rs("ubicacion") = "D" Then
       tci = tci + rs("cyb_05.IMPORTE")
    Else
       tce = tce + rs("cyb_05.IMPORTE")
    End If
    rs.MoveNext
  Wend
  Set rs = Nothing
  
   If tci > 0 Or tce > 0 Then
    tti = tti + tci
    tte = tte + tce
    
    F = ""
    i = Format$(tci, "######0.00")
    e = Format$(tce, "######0.00")
    msf1.AddItem F & Chr$(9) & r & Chr$(9) & i & Chr$(9) & e & Chr$(9) & s & Chr$(9) & c & Chr$(9) & "" & Chr$(9) & ""
   End If
  rs2.MoveNext
  Wend
  Set rs2 = Nothing
 
  
   'calcula saldo anterior
  If Check1 = 0 Then
   If qa <> "" Then
    Set rs = New ADODB.Recordset
    rs.Open qa, cn1
    sa = 0
    While Not rs.EOF
     If rs("ubicacion") = "D" Then
       sa = sa + rs("IMPORTE")
     Else
       sa = sa - rs("IMPORTE")
     End If
     rs.MoveNext
    Wend
    Set rs = Nothing
   Else
    sa = 0
   End If
  Else
   sa = 0
  End If
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "--------------------------------" & Chr$(9) & "----------------------------"
  msf1.AddItem "Saldo Ant. " & Chr$(9) & Format$(sa, "#######0.00") & " <====" & Chr$(9) & Format$(tti, "######0.00") & Chr$(9) & Format$(tte, "######0.00") & Chr$(9) & "==> " & Format$(sa + tti - tte, "######0.00")
  msf1.SetFocus
  Set rs = Nothing
    
    
   
End Sub


Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.FixedCols = 0
'msf1.SelectionMode = flexSelectionByRow
'msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 2500
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 1500
msf1.ColWidth(5) = 1500
msf1.ColWidth(6) = 2000
msf1.ColWidth(7) = 800
msf1.TextMatrix(0, 0) = "Fecha  "
msf1.TextMatrix(0, 1) = "Operacion"
msf1.TextMatrix(0, 2) = "INGRESOS"
msf1.TextMatrix(0, 3) = "EGRESOS"
msf1.TextMatrix(0, 4) = "Cuenta"
msf1.TextMatrix(0, 5) = "Concepto"
msf1.TextMatrix(0, 6) = "Detalle"
msf1.TextMatrix(0, 7) = "Num.Mov."

msf1.ColAlignment(1) = 1
For i = 2 To 3
  msf1.ColAlignment(i) = 9
Next i

For i = 4 To 6
  msf1.ColAlignment(i) = 1
Next i

End Sub



Private Sub Command1_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
If para.cuenta_sel > 0 Then
  C_subrubro.ListIndex = buscaindice(C_subrubro, para.cuenta_sel)
  para.cuenta_sel = 0
End If
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_cuentas_cont(C_subrubro, "C", "D")
C_subrubro.AddItem "<Todos>", 0
C_subrubro.ListIndex = 0
Call carga_formas_pago(c_concepto, "T")
c_concepto.AddItem "<Todos>", 0
c_concepto.ListIndex = 0

Call carga_usuarios(c_usuario)
c_usuario.AddItem "<Todos>", 0
c_usuario.ListIndex = 0

c_tipo.ListIndex = 0
Load cyb_movcaja
Option1 = True
Call armagrid

Call carga_cuentas_cont(c_n1, "T", "D")
c_n1.AddItem "<Todos>", 0
c_n1.ListIndex = 0

para.cuenta_sel = 0
Check1 = False
Option2 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload cyb_movcaja
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F1] Agregar - [F7] Imprime - [F8] Borra Mov. - [ENTER] Modifica -[F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 7 Then
      cyb_movcaja.limpia
      cyb_movcaja.c_tipo.ListIndex = 0
      cyb_movcaja.t_fecha = Format$(Now, "DD/MM/YYYY")
      cyb_movcaja.Show
 Else
   Call sinpermisos
 End If
End If


If KeyCode = vbKeyF7 Then
Dim c(15) As Double
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 5 Then
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
   
    For i = 8 To 14
      c(i) = -1
    Next i
    
    'If t_fecha <> "" Or T_FECHA2 <> "" Then
    t1 = "Periodo Desde........: " & t_fecha1 & " Hasta: " & t_fecha2
    T2 = ""
    t3 = "Acumulado por........: " & msf1.Tag
                   
    
    
    Call imprimegrid(msf1, c(), "MOVIMIENTOS DE CAJA", t1, T2, t3, 50, 8, True, False, "H")
  End If
Else
 Call sinpermisos
End If
End If


If KeyCode = vbKeyF8 Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 8 Then
  n = Val(msf1.TextMatrix(msf1.Row, 7))
  If Val(n) > 0 Then
    Call borramovcaja(n)
    Call limpia
  End If
 Else
  Call sinpermisos
 End If

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call nivel_acceso(1)
 If para.id_grupo_modulo_actual > 8 Then
 
  ni = Val(msf1.TextMatrix(msf1.Row, 7))
  If ni > 0 Then
      cyb_movcaja.t_numint = msf1.TextMatrix(msf1.Row, 7)
      q = "SELECT * FROM cyb_05 WHERE [num_mov_caja] = " & ni
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
         cyb_movcaja.limpia
         cyb_movcaja.t_numint = rs("num_mov_caja")
         cyb_movcaja.t_fecha = rs("FECHA")
         cyb_movcaja.t_destino = rs("DEscripcion")
         cyb_movcaja.t_importe = rs("IMPORTE")
         cyb_movcaja.c_caja.ListIndex = buscaindice(cyb_movcaja.c_caja, rs("ID_forma_pago"))
         cyb_movcaja.c_cuenta.ListIndex = buscaindice(cyb_movcaja.c_cuenta, rs("ID_cuenta_contra"))
         cyb_movcaja.t_op = rs("operacion")
         If rs("UBICACION") = "H" Then
            cyb_movcaja.c_tipo.ListIndex = 1
         Else
            cyb_movcaja.c_tipo.ListIndex = 0
         End If
         Set rs = Nothing
         cyb_movcaja.Show
      End If
    End If
  Else
    Call sinpermisos
  End If
 End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
End Sub

Private Sub T_detalle_GotFocus()
t_detalle = ""
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub


Private Sub t_importe_GotFocus()
t_importe = ""
End Sub
