VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_informevta3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE PRODUCTOS  PENDIENTES DE FACTURACION"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detallado por:"
      Height          =   615
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   3615
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto"
         Height          =   195
         Left            =   240
         TabIndex        =   23
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
      Begin VB.ComboBox c_grupo 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox t_descprod 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_prod 
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
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Grupo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Desc. Producto: "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Producto:"
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
      StartOfWeek     =   174915585
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
         Picture         =   "VTA025.frx":0000
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
         Picture         =   "VTA025.frx":0882
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
      Top             =   8445
      Width           =   12120
      _ExtentX        =   21378
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
            TextSave        =   "05/09/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "01:28 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   -360
      TabIndex        =   10
      Top             =   2160
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8916
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4200
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "vta_informevta3"
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
  espere.Show
  espere.Refresh
  'selecciono productos
  q = "select * from a2 where [id_producto] > 1 "
  c = " and "
  
  If c_prod.ListIndex > 0 Then
    q = q & c & " [id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
    c = " and "
  End If
  
  If t_descprod <> "" Then
    q = q & c & " [descripcion] like '%" & t_descprod & "%'"
    c = " and "
  End If
  
  If c_grupo.ListIndex > 0 Then
    q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
    c = " and "
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttp = 0
  ttr = 0
  reg2 = 0
  While Not rs.EOF
    'busco el producto en las ventas
      reg2 = reg2 + 1
      q = "select * from vta_02, vta_03, vta_01 where [id_producto] = " & rs("id_producto") & " and vta_03.[num_int] = vta_02.[num_int] and vta_02.[id_cliente] = vta_01.[id_cliente] and [id_tipocomp] = 45 and [estado] <> 'F'"
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
      tp = 0
      tr = 0
      reg = 0
      Label8 = reg2
      Label8.Refresh
      While Not rs2.EOF
            tp = tp + (rs2("cantidad"))
            ttp = ttp + (rs2("cantidad"))
            reg = reg + 1
            rs2.MoveNext
      Wend
      Set rs2 = Nothing
      ip = rs("id_producto")
      dp = rs("descripcion")
      If tp > 0 Then
       msf1.AddItem ip & Chr(9) & dp & Chr(9) & " " & Chr(9) & Format$(tp, "#####0.00")
      End If
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttp, "#####0.00")
   Set rs = Nothing
  
   Unload espere
   
   
   
End Sub
Sub carga2()
  Call armagrid
  espere.Show
  espere.Refresh
  'selecciono productos
  q = "select * from a2 where [id_producto] > 1 "
  c = " and "
  
  If c_prod.ListIndex > 0 Then
    q = q & c & " [id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
    c = " and "
  End If
  
  If t_descprod <> "" Then
    q = q & c & " [descripcion] like '%" & t_descprod & "%'"
    c = " and "
  End If
  
  If c_grupo.ListIndex > 0 Then
    q = q & c & " [id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
    c = " and "
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  ttp = 0
  ttr = 0
  reg2 = 0
  While Not rs.EOF
    'busco el producto en las ventas
      p1 = 0
      reg2 = reg2 + 1
      q = "select * from vta_02, vta_03, vta_01 where [id_producto] = " & rs("id_producto") & " and vta_03.[num_int] = vta_02.[num_int] and vta_02.[id_cliente] = vta_01.[id_cliente] and [id_tipocomp] = 45  and [cantidad] > 0 and [estado] <> 'F'"
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
        
      q = q & " order by vta_02.[id_cliente]"
      Set rs2 = New ADODB.Recordset
      rs2.Open q, cn1
      tp = 0
      tr = 0
      reg = 0
      Label8 = reg2
      Label8.Refresh
      c1 = 0
      While Not rs2.EOF
         If p1 = 0 Then
           ip = rs("id_producto")
           dp = rs("descripcion")
           msf1.AddItem ip & Chr(9) & dp
           p1 = 1
           tr = 0
         End If
            
         If c1 = 0 Then
           c1 = rs2("vta_02.id_cliente")
           cli = rs2("denominacion")
         End If
         
         If c1 <> rs2("vta_02.id_cliente") Then
            msf1.AddItem "" & Chr(9) & "" & Chr(9) & cli & Chr(9) & Format$(tp, "#####0.00")
            tp = 0
            c1 = rs2("vta_02.id_cliente")
            cli = rs2("denominacion")
         End If
         tp = tp + (rs2("cantidad"))
         tr = tr + (rs2("cantidad"))
         
         ttp = ttp + (rs2("cantidad"))
         reg = reg + 1
         rs2.MoveNext
      Wend
      Set rs2 = Nothing
      If tp > 0 Then
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & cli & Chr(9) & Format$(tp, "#####0.00")
      End If
      If tr > 0 Then
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Total del producto ----> " & Chr(9) & "" & Chr$(9) & Format$(tr, "#####0.00")
      End If
      
      rs.MoveNext
   Wend
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttp, "#####0.00")
   Set rs = Nothing
  
   Unload espere
  
   

End Sub
Private Sub btnacepta_Click()
QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Informe de productos pendientes de Facturacion " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 17, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
If Option4 = True Then
   Call carga
Else
 If Option5 = True Then
    Call carga2
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

Private Sub c_prod_LostFocus()
If c_prod.ListIndex < 0 Then
  c_prod.ListIndex = 0
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
msf1.Cols = 6
msf1.ColWidth(0) = 700
msf1.ColWidth(1) = 4000
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 500
If Option4 = True Then
 msf1.TextMatrix(0, 0) = "Id."
 msf1.TextMatrix(0, 1) = "Producto"
 msf1.TextMatrix(0, 2) = " "
Else
  msf1.TextMatrix(0, 0) = "Id."
  msf1.TextMatrix(0, 1) = "Producto "
  msf1.TextMatrix(0, 2) = "Cliente/Vendedor"
  
End If
 msf1.TextMatrix(0, 3) = "Pendiente"
msf1.TextMatrix(0, 4) = " "
msf1.TextMatrix(0, 5) = ""


For i = 0 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(3) = 9 'der
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

Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0

Call carga_grupos(c_grupo)
c_grupo.AddItem "<Todos>", 0
c_grupo.ListIndex = 0


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
    c(0) = 5
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    
    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "INFORME DE PRODUCTOS PENDIENTES DE FACTURACION(UNIDADES)", "Vendedor: " & c_vend & "   Producto: " & c_prod & "  " & t_descprod, "Periodo: " & t_fecha & " : " & t_fecha2, "Cliente: " & c_prov, 90, 7, True, False)
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    On Error GoTo nada
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

Exit Sub
nada:
  Resume Next
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
