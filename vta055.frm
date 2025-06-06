VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_informevta5 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE VENTAS  por ACUMULADOS(IMPORTES)"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Acumulado por ..."
      Height          =   615
      Left            =   0
      TabIndex        =   24
      Top             =   7200
      Width           =   9615
      Begin VB.OptionButton Option9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Punto Vernta"
         Height          =   195
         Left            =   8160
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marca"
         Height          =   195
         Left            =   7080
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Departamento"
         Height          =   195
         Left            =   5400
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grupo"
         Height          =   195
         Left            =   4200
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clientes"
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Productos"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   3615
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id. "
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   5520
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox c_marca 
         Height          =   315
         Left            =   1440
         TabIndex        =   31
         Top             =   1680
         Width           =   4575
      End
      Begin VB.ComboBox c_grupo 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_dep 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   960
         Width           =   4575
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   4575
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Marca:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Grupo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Departamento:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   5160
      TabIndex        =   9
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   178454529
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   3615
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
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
         Picture         =   "vta055.frx":0000
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
         Picture         =   "vta055.frx":0882
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
      Top             =   8520
      Width           =   12090
      _ExtentX        =   21325
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
            TextSave        =   "06/06/2025"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:55 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8916
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label9 
      Caption         =   $"vta055.frx":1104
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   7920
      Width           =   8895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4200
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "vta_informevta5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'FIXIT: Declare 'ti' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Dim ti, t As Double
'FIXIT: Declare 'reg' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim reg, regi As Integer


Sub carga()
  Call armagrid
  'selecciono productos
  q = "select * from a2  "
  c = " where "
  
   If c_dep.ListIndex > 0 Then
    q = q & c & " [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
    c = " and "
  End If
  
   If c_marca.ListIndex > 0 Then
    q = q & c & " [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
    c = " and "
  End If
  
  If c_grupo.ListIndex > 0 Then
    q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
    c = " and "
  End If
  
  If Option4 = True Then
     q = q & " order by [id_producto]"
  Else
     q = q & " order by [descripcion]"
  End If

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  reg2 = 0
  ttcosto = 0
  espere.Show
  
  While Not rs.EOF
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, vta_01 where [id_producto] = " & rs("id_producto") & " and vta_03.[num_int] = vta_02.[num_int] and vta_02.[id_cliente] = vta_01.[id_cliente] "
      c = " and "
  
      If c_prov.ListIndex > 0 Then
        q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
      End If
  
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_vend.ListIndex > 0 Then
         q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
      End If
        
      Set rs2 = New ADODB.Recordset
      rs2.Open q, cn1
      tsi = 0
      tf = 0
      tc = 0
      tcosto = 0
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If rs2("id_producto") > 1 Then
              costo = rs2("costo")
             Else
              costo = 0
             End If
             
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
               tcosto = tcosto + (costo * c)
               ttcosto = ttcosto + (costo * c)
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
               tcosto = tcosto - (costo * c)
               ttcosto = ttcosto - (costo * c)
            End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      ip = rs("id_producto")
      dp = rs("descripcion")
      If tf > 0 Then
        
        
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00") & Chr(9) & Format$(tcosto, "#####0.00") & Chr(9) & Format$(tsi - tcosto, "#####0.00") & Chr(9) & Format$((tsi / (tcosto + 1)) - 1, "##0%")
        msf1.Refresh
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00") & Chr(9) & Format$(ttcosto, "#####0.00") & Chr(9) & Format$(ttsi - ttcosto, "#####0.00") & Chr(9) & Format$((ttsi / (ttcosto + 1)) - 1, "##0%")
   Set rs = Nothing
   Unload espere
  
   
   
   
End Sub

Sub carga2()
  Call armagrid
  'selecciono clientes
  q = "select * from vta_01  "
  
  If c_prov.ListIndex > 0 Then
    q = q & " where [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  
  If Option4 = True Then
     q = q & " order by [id_cliente]"
  Else
     q = q & " order by [denominacion]"
  End If

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  reg2 = 0
  ttcosto = 0
  espere.Show
  
  While Not rs.EOF
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, a2 where [id_cliente] = " & rs("id_cliente") & " and vta_03.[num_int] = vta_02.[num_int] and vta_03.[id_producto] = a2.[id_producto] "
      c = " and "
  
      
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_vend.ListIndex > 0 Then
         q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
      End If
      
     If c_dep.ListIndex > 0 Then
        q = q & c & " [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
        c = " and "
     End If
  
     If c_marca.ListIndex > 0 Then
       q = q & c & " [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
       c = " and "
     End If
  
     If c_grupo.ListIndex > 0 Then
       q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
       c = " and "
     End If
        
      Set rs2 = New ADODB.Recordset
      rs2.Open q, cn1
      tsi = 0
      tf = 0
      tc = 0
      tcosto = 0
      
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If rs2("vta_03.id_producto") > 1 Then
               costo = rs2("costo")
             Else
               costo = 0
             End If
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
               tcosto = tcosto + (costo * c)
               ttcosto = ttcosto + (costo * c)
               
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
               tcosto = tcosto - (costo * c)
               ttcosto = ttcosto - (costo * c)
             End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      ip = rs("id_cliente")
      dp = rs("denominacion")
      If tf > 0 Then
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00") & Chr(9) & Format$(tcosto, "#####0.00") & Chr(9) & Format$(tsi - tcosto, "#####0.00") & Chr(9) & Format$((tsi / (tcosto + 1)) - 1, "##0%")
        msf1.Refresh
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00") & Chr(9) & Format$(ttcosto, "#####0.00") & Chr(9) & Format$(ttsi - ttcosto, "#####0.00") & Chr(9) & Format$((ttsi / (ttcosto + 1)) - 1, "##0%")
   Set rs = Nothing
   Unload espere
  
   
   
   
End Sub

Sub carga9()
 Dim s(20) As Integer
 
 For i = 0 To 19
   s(i) = 0
 Next i
 
Call armagrid
  'selecciono puntos de venta
  Set rs = New ADODB.Recordset
  q = "select * from vta_06 order by [SUCURSAL]"
  rs.Open q, cn1
  p = 0
  b = 0
 While Not rs.EOF
  If p = 0 Then
     s(0) = rs("SUCURSAL")
     p = p + 1
     b = rs("sucursal")
  End If
  
  If b <> rs("SUCURSAL") Then
     s(p) = rs("SUCURSAL")
     p = p + 1
     b = rs("SUCURSAL")
  End If
  rs.MoveNext
Wend
Set rs = Nothing
  
  
  
  

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  reg2 = 0
  ttcosto = 0
  espere.Show
  
  For i = 0 To 19
     If s(i) <> 0 Then
      
    
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, a2 where [sucursal_ingreso] = " & s(i) & " and vta_03.[num_int] = vta_02.[num_int] and vta_03.[id_producto] = a2.[id_producto] "
      c = " and "
  
      
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_vend.ListIndex > 0 Then
         q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
      End If
      
     If c_dep.ListIndex > 0 Then
        q = q & c & " [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
        c = " and "
     End If
  
     If c_marca.ListIndex > 0 Then
       q = q & c & " [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
       c = " and "
     End If
  
     If c_grupo.ListIndex > 0 Then
       q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
       c = " and "
     End If
        
      Set rs2 = New ADODB.Recordset
      rs2.Open q, cn1
      tsi = 0
      tf = 0
      tc = 0
      tcosto = 0
      
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If rs2("vta_03.id_producto") > 1 Then
               costo = rs2("costo")
             Else
               costo = 0
             End If
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
               tcosto = tcosto + (costo * c)
               ttcosto = ttcosto + (costo * c)
               
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
               tcosto = tcosto - (costo * c)
               ttcosto = ttcosto - (costo * c)
             End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      ip = s(i)
      dp = "Punto de Venta " & s(i)
      'If tf > 0 Then
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00") & Chr(9) & Format$(tcosto, "#####0.00") & Chr(9) & Format$(tsi - tcosto, "#####0.00") & Chr(9) & Format$((tsi / (tcosto + 1)) - 1, "##0%")
        msf1.Refresh
      'End If
     Else
      i = 20
     End If
   Next i
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00") & Chr(9) & Format$(ttcosto, "#####0.00") & Chr(9) & Format$(ttsi - ttcosto, "#####0.00") & Chr(9) & Format$((ttsi / (ttcosto + 1)) - 1, "##0%")
  
   Unload espere
  
   
   
   
End Sub

Sub carga3()
  Call armagrid
  'selecciono clientes
  q = "select * from vta_05  "
  
  If c_prov.ListIndex > 0 Then
    q = q & " where [id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
  
  If Option4 = True Then
     q = q & " order by [id_vendedor]"
  Else
     q = q & " order by [denominacion]"
  End If

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  reg2 = 0
  
  espere.Show
  
  While Not rs.EOF
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, a2 where [id_vendedor] = " & rs("id_vendedor") & " and vta_03.[num_int] = vta_02.[num_int] and vta_03.[id_producto] = a2.[id_producto] "
      c = " and "
  
      
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_prov.ListIndex > 0 Then
         q = q & c & " [Id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
      End If
      
     If c_dep.ListIndex > 0 Then
        q = q & c & " [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
        c = " and "
     End If
  
     If c_marca.ListIndex > 0 Then
       q = q & c & " [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
       c = " and "
     End If
  
     If c_grupo.ListIndex > 0 Then
       q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
       c = " and "
     End If
        
      Set rs2 = New ADODB.Recordset
      rs2.Open q, cn1
      tsi = 0
      tf = 0
      tc = 0
      
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
             End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      ip = rs("id_vendedor")
      dp = rs("denominacion")
      If tf > 0 Then
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00")
        msf1.Refresh
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00")
   Set rs = Nothing
   Unload espere
  
   
   
   
End Sub

Sub carga6()
  Call armagrid
  'selecciono clientes
  q = "select * from A8  "
  
  If c_grupo.ListIndex > 0 Then
    q = q & " where [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
  End If
  
  
  If Option4 = True Then
     q = q & " order by [id_grupo]"
  Else
     q = q & " order by [descripcion]"
  End If

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  ttcosto = 0
  reg2 = 0
  espere.Show
  While Not rs.EOF
    'buscoproductos del grupo
    q = "select * from a2 where [id_grupo] = " & rs("id_grupo")
    Set rs3 = New ADODB.Recordset
    rs3.Open q, cn1
    
     tsi = 0
      tf = 0
      tc = 0
      tcosto = 0
    While Not rs3.EOF
  
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, a2 where [vta_03.id_producto] = " & rs3("id_producto") & " and vta_03.[num_int] = vta_02.[num_int]   and [vta_03.id_producto] = [a2.id_producto]"
      c = " and "
  
      
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_prov.ListIndex > 0 Then
         q = q & c & " [Id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
      End If
      
     If c_dep.ListIndex > 0 Then
        q = q & c & " [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
        c = " and "
     End If
  
     If c_marca.ListIndex > 0 Then
       q = q & c & " [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
       c = " and "
     End If
  
    
        
      Set rs2 = New ADODB.Recordset
     
      rs2.Open q, cn1
     
      
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If rs2("vta_03.id_producto") > 1 Then
               costo = rs2("costo")
             Else
               costo = 0
             End If
             
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
               tcosto = tcosto + (costo * c)
               ttcosto = ttcosto + (costo * c)
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
               tcosto = tcosto - (costo * c)
               ttcosto = ttcosto - (costo * c)
             End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      rs3.MoveNext
      Wend
      Set rs3 = Nothing
      ip = rs("id_grupo")
      dp = rs("descripcion")
      If tf > 0 Then
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00") & Chr(9) & Format$(tcosto, "#####0.00") & Chr(9) & Format$(tsi - tcosto, "#####0.00") & Chr(9) & Format$((tsi / (tcosto + 1)) - 1, "##0%")
        msf1.Refresh
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00") & Chr(9) & Format$(ttcosto, "#####0.00") & Chr(9) & Format$(ttsi - ttcosto, "#####0.00") & Chr(9) & Format$((ttsi / (ttcosto + 1)) - 1, "##0%")
   Set rs = Nothing
   Unload espere
  
   
   
   
End Sub


Sub carga7()
  Call armagrid
  'selecciono departamento
  q = "select * from A9  "
  
  If c_dep.ListIndex > 0 Then
    q = q & " where [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
  End If
  
  
  If Option4 = True Then
     q = q & " order by [id_departamento]"
  Else
     q = q & " order by [descripcion]"
  End If

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  ttcosto = 0
  reg2 = 0
  espere.Show
  While Not rs.EOF
    'buscoproductos del grupo
    q = "select * from a2 where [id_departamento] = " & rs("id_departamento")
    Set rs3 = New ADODB.Recordset
    rs3.Open q, cn1
    
     tsi = 0
      tf = 0
      tc = 0
      tcosto = 0
    While Not rs3.EOF
  
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, a2 where [vta_03.id_producto] = " & rs3("id_producto") & " and vta_03.[num_int] = vta_02.[num_int]   and [vta_03.id_producto] = [a2.id_producto]"
      c = " and "
  
      
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_prov.ListIndex > 0 Then
         q = q & c & " [Id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
      End If
      
     If c_grupo.ListIndex > 0 Then
        q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
        c = " and "
     End If
  
     If c_marca.ListIndex > 0 Then
       q = q & c & " [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
       c = " and "
     End If
  
    
        
      Set rs2 = New ADODB.Recordset
     
      rs2.Open q, cn1
     
      
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If rs2("vta_03.id_producto") > 1 Then
               costo = rs2("costo")
             Else
               costo = 0
             End If
             
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
               tcosto = tcosto + (costo * c)
               ttcosto = ttcosto + (costo * c)
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
               tcosto = tcosto - (costo * c)
               ttcosto = ttcosto - (costo * c)
             End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      rs3.MoveNext
      Wend
      Set rs3 = Nothing
      ip = rs("id_departamento")
      dp = rs("descripcion")
      If tf > 0 Then
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00") & Chr(9) & Format$(tcosto, "#####0.00") & Chr(9) & Format$(tsi - tcosto, "#####0.00") & Chr(9) & Format$((tsi / (tcosto + 1)) - 1, "##0%")
        msf1.Refresh
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00") & Chr(9) & Format$(ttcosto, "#####0.00") & Chr(9) & Format$(ttsi - ttcosto, "#####0.00") & Chr(9) & Format$((ttsi / (ttcosto + 1)) - 1, "##0%")
   Set rs = Nothing
   Unload espere
  
   
   
   
End Sub

Sub carga8()
  Call armagrid
  'selecciono departamento
  q = "select * from A10  "
  
  If c_marca.ListIndex > 0 Then
    q = q & " where [id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
  End If
  
  
  If Option4 = True Then
     q = q & " order by [id_marca]"
  Else
     q = q & " order by [descripcion]"
  End If

  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttsi = 0
  ttf = 0
  ttcosto = 0
  reg2 = 0
  espere.Show
  While Not rs.EOF
    'buscoproductos del grupo
    q = "select * from a2 where [id_marca] = " & rs("id_marca")
    Set rs3 = New ADODB.Recordset
    rs3.Open q, cn1
    
     tsi = 0
      tf = 0
      tc = 0
      tcosto = 0
    While Not rs3.EOF
  
    'busco el producto en las ventas
      reg2 = reg2 + 1
      espere.Label1 = "Espere .... Registro Nro. " & reg2
      espere.Label1.Refresh
      
      q = "select * from vta_02, vta_03, a2 where [vta_03.id_producto] = " & rs3("id_producto") & " and vta_03.[num_int] = vta_02.[num_int]   and [vta_03.id_producto] = [a2.id_producto]"
      c = " and "
  
      
      If IsDate(t_fecha) Then
        q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
      End If
  
      If IsDate(t_fecha2) Then
        q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
  
      If c_prov.ListIndex > 0 Then
         q = q & c & " [Id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
      End If
      
     If c_grupo.ListIndex > 0 Then
        q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
        c = " and "
     End If
  
     If c_dep.ListIndex > 0 Then
       q = q & c & " [id_departamento] = " & c_dep.ItemData(c_dep.ListIndex)
       c = " and "
     End If
  
    
        
      Set rs2 = New ADODB.Recordset
     
      rs2.Open q, cn1
     
      
      While Not rs2.EOF
            Set rs1 = New ADODB.Recordset
            q = "select [venta] from vta_06 where [sucursal] = " & rs2("sucursal_ingreso") & " and [id_tipocomp] = " & rs2("id_tipocomp")
            rs1.Open q, cn1
            If Not rs1.EOF And Not rs1.BOF Then
              v = rs1("venta")
            Else
              v = "N"
            End If
            If v <> "N" Then
             If rs2("letra") = "A" Then
                pf = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                psi = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             Else
                psi = Format((rs2("cantidad_original") * rs2("pu_final")), "#####0.00")
                pf = Format((rs2("cantidad_original") * rs2("vta_03.pu")), "#####0.00")
             End If
             c = rs2("cantidad_original")
             If rs2("vta_03.id_producto") > 1 Then
               costo = rs2("costo")
             Else
               costo = 0
             End If
             
             If v = "S" Then
               tf = tf + pf
               ttf = ttf + pf
               tsi = tsi + psi
               ttsi = ttsi + psi
               tc = tc + c
               tcosto = tcosto + (costo * c)
               ttcosto = ttcosto + (costo * c)
             Else
               tf = tf - pf
               ttf = ttf - pf
               tsi = tsi - psi
               ttsi = ttsi - psi
               tc = tc - c
               tcosto = tcosto - (costo * c)
               ttcosto = ttcosto - (costo * c)
             End If
           
           End If
           Set rs1 = Nothing
        pf = 0
        psi = 0
        c = 0
        reg = reg + 1
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      rs3.MoveNext
      Wend
      Set rs3 = Nothing
      ip = rs("id_marca")
      dp = rs("descripcion")
      If tf > 0 Then
        msf1.AddItem ip & Chr(9) & dp & Chr(9) & tc & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00") & Chr(9) & Format$(tcosto, "#####0.00") & Chr(9) & Format$(tsi - tcosto, "#####0.00") & Chr(9) & Format$((tsi / (tcosto + 1)) - 1, "##0%")
        msf1.Refresh
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttsi, "#####0.00") & Chr(9) & Format$(ttcosto, "#####0.00") & Chr(9) & Format$(ttsi - ttcosto, "#####0.00") & Chr(9) & Format$((ttsi / (ttcosto + 1)) - 1, "##0%")
   Set rs = Nothing
   Unload espere
  
   
   
   
End Sub
Private Sub btnacepta_Click()
J = MsgBox("Este proceso puede demorar y afectar el rendidmiento de otras terminales, ¿desea continuar?", 4)
If J = 6 Then

QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Inf. de Ventas Acum. por prod.(unidades) " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 16, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
  
  
If Option2 = True Then
   Call carga
End If

If Option3 = True Then
    Call carga2
  
End If
 
If Option1 = True Then
    Call carga3
End If


If Option6 = True Then
    Call carga6
End If


If Option7 = True Then
    Call carga7
End If

If Option8 = True Then
    Call carga8
End If

If Option9 = True Then
    Call carga9
End If

End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub c_grupo_LostFocus()
If c_grupo.ListIndex < 0 Then
  c_grupo.ListIndex = 0
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
msf1.Cols = 8
msf1.ColWidth(0) = 700
msf1.ColWidth(1) = 4000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1300
msf1.ColWidth(4) = 1300
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 900

If Option2 = True Then
 msf1.TextMatrix(0, 0) = "Id."
 msf1.TextMatrix(0, 1) = "Producto"
End If
If Option3 = True Then
 msf1.TextMatrix(0, 0) = "Id."
 msf1.TextMatrix(0, 1) = "Cliente"
End If
If Option1 = True Then
 msf1.TextMatrix(0, 0) = "Id."
 msf1.TextMatrix(0, 1) = "Vendedor"
End If
If Option6 = True Then
 msf1.TextMatrix(0, 0) = "Id."
 msf1.TextMatrix(0, 1) = "Grupo"
End If

 msf1.TextMatrix(0, 2) = "Cantidad "
 msf1.TextMatrix(0, 7) = " "
 msf1.TextMatrix(0, 3) = "$ Acum. Final"
 msf1.TextMatrix(0, 4) = "$ Acum. S/ iva"
 msf1.TextMatrix(0, 5) = "Costo Real S/ iva"
 msf1.TextMatrix(0, 6) = "Utilidad Bruta"
 msf1.TextMatrix(0, 7) = "% U.B"
For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 6
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Sub armagrid2()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 700
msf1.ColWidth(1) = 4000
msf1.ColWidth(2) = 1600
msf1.ColWidth(3) = 1600
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(6) = 1000

 msf1.TextMatrix(0, 0) = "Id."
 msf1.TextMatrix(0, 1) = "Cliente"
 msf1.TextMatrix(0, 2) = "Cantidad "
 msf1.TextMatrix(0, 7) = " "



 msf1.TextMatrix(0, 3) = "$ Acum. Final"
 msf1.TextMatrix(0, 4) = "$ Acum. S/ iva"


For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 6
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub
Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call carga_clientes(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0

Call carga_deptos_venta(c_dep)
c_dep.AddItem "<Todos>", 0
c_dep.ListIndex = 0

Call carga_grupos(c_grupo)
c_grupo.AddItem "<Todos>", 0
c_grupo.ListIndex = 0

Call carga_marcas(c_marca)
c_marca.AddItem "<Todos>", 0
c_marca.ListIndex = 0

Option4 = True
Option2 = True

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
    c(6) = 6
    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "INFORME DE VENTAS ACUMULADO POR PRODUCTOS(UNIDADES)", "Vendedor: " & c_vend, "Periodo: " & t_fecha & " : " & t_fecha2, "Cliente: " & c_prov, 90, 7, True, False)
  End If

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
     Load vta_movprodcli
     vta_movprodcli.c_prod.ListIndex = buscaindice(vta_movprodcli.c_prod, Val(msf1.TextMatrix(msf1.Row, 0)))
     If t_fecha <> "" Then
      vta_movprodcli.t_fecha = t_fecha
     End If
     If t_fecha2 <> "" Then
      vta_movprodcli.t_fecha2 = t_fecha2
     End If
     If c_prov.ListIndex > 0 Then
      vta_movprodcli.c_prov.ListIndex = buscaindice(vta_movprodcli.c_prov, c_prov.ItemData(c_prov.ListIndex))
     End If
     vta_movprodcli.Show
    
    Else
     If Val(msf1.TextMatrix(msf1.Row, 7)) > 0 Then
      Load vta_cc_detalle
      vta_cc_detalle.t_numint = Val(msf1.TextMatrix(msf1.Row, 7))
      vta_cc_detalle.Show
     End If
    End If
  
  End If
End If

End Sub



Private Sub t_descprod_GotFocus()
t_descprod = ""
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
