VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form pro_costos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ESTRUCTURA DE COSTOS"
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
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   3015
      Begin VB.TextBox t_dolar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dolar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11033
      _Version        =   393216
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox c_prov 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   9855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Pieza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
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
      Left            =   10080
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "pro010.frx":0000
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
         Picture         =   "pro010.frx":0882
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
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   26458
            MinWidth        =   26458
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "pro_costos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double

Sub carga()
  espere.Show
  espere.Label1 = "Espere... cargando estructura y calculando costos"
  espere.Refresh
  Call armagrid
  ip = c_prov.ItemData(c_prov.ListIndex)
  q = "select * from pro_07 where [id_pieza] = " & c_prov.ItemData(c_prov.ListIndex) & " order by [renglon]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
    If rs("tipo") = "P" Then
      'Armo pieza en estructura
        d = "  ***" & rs("detalle") & "***"
        msf1.AddItem "" & Chr(9) & d
 
       Call buscaestructura(rs("id"))
    Else
      'busco producto en lista de prcios
       Call cargaproducto(rs("id_pieza"), rs("renglon"))
      ' msf1.AddItem rs("renglon") & Chr(9) & rs("detalle") & Chr(9) & rs("cantidad") & Chr(9) & rs("Unidad") & Chr(9) & rs("id") & Chr(9) & rs("tipo")
    End If
    rs.MoveNext
  Wend
  Set rs = Nothing
  
  Unload espere
End Sub
Sub cargaproducto(ByVal p As Long, ByVal r As Integer)
  Set rs3 = New ADODB.Recordset
  q = "select * from pro_07 where [id_pieza] = " & p & " and [renglon] = " & r
  rs3.MaxRecords = 1
  rs3.Open q, cn1
  If Not rs3.EOF And Not rs3.BOF Then
    d = rs3("detalle")
    c = rs3("cantidad")
    u = rs3("unidad")
    idp = rs3("id")
  Else
    d = "Error en la estructura"
    c = 0
    u = " "
  End If
  Set rs3 = Nothing
  
  
  If idp > 1 Then
   Set rs1 = New ADODB.Recordset
   q = "select [fecha_ult_compra], [precio_ult_compra], [denominacion], [moneda]  from a2, a1 where [id_producto] = " & idp & " and a1.[id_proveedor] = [id_proveedor_ult_compra]"
  
  'MsgBox (q)
   rs1.MaxRecords = 1
   rs1.Open q, cn1
 
   If Not rs1.EOF And Not rs1.BOF Then
     ful = rs1("fecha_ult_compra")
     prov = rs1("denominacion")
     If rs1("moneda") = "P" Then
       pc = rs1("precio_ult_compra")
     Else
       pc = rs1("precio_ult_compra") * Val(t_dolar)
     End If
   Else
     ful = "01/01/2000"
     prov = "Error producto dado de baja"
     pc = 0
   End If
   
  Else
     ful = " "
     prov = "Producto Manual"
     pc = 0
  End If
  msf1.AddItem idp & Chr(9) & d & Chr(9) & c & Chr(9) & pc & Chr(9) & Format$(c * pc, "#####0.00") & Chr(9) & u & Chr(9) & ful & Chr(9) & prov
  Set rs1 = Nothing
End Sub
Sub buscaestructura(ByVal ip)
  
  
  q = "select * from pro_07 where [id_pieza] = " & ip & " order by [renglon]"
  Set rs2 = New ADODB.Recordset
  rs2.Open q, cn1
  While Not rs2.EOF

    If rs2("tipo") = "P" Then
         d = "  ***" & rs2("detalle") & "***"
        msf1.AddItem "" & Chr(9) & d
 
      'busco pieza en estructura
       Call buscaestructura(rs2("id_pieza"))
       
    Else
       Call cargaproducto(rs2("id_pieza"), rs2("renglon"))
    End If
    rs2.MoveNext
  Wend
  Set rs2 = Nothing

End Sub

Private Sub btnacepta_Click()
Call carga
If msf1.Rows > 1 Then
  tot = suma_msflexgrid(msf1, 4)
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "-------------------------"
  msf1.AddItem "" & Chr$(9) & "*******  COSTO TOTAL ------------>" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & Format$(tot, "#######0.00")
End If
  
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 8
  msf1.ColWidth(0) = 600
  msf1.ColWidth(1) = 6000
  msf1.ColWidth(2) = 1100
  msf1.ColWidth(3) = 1200
  msf1.ColWidth(4) = 1400
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1200
  msf1.ColWidth(7) = 2500
  msf1.TextMatrix(0, 0) = "Id."
  msf1.TextMatrix(0, 1) = "DETALLE"
  msf1.TextMatrix(0, 2) = "Cantidad"
  msf1.TextMatrix(0, 3) = "Pu"
  msf1.TextMatrix(0, 4) = "Importe"
  msf1.TextMatrix(0, 5) = "Unidad"
  msf1.TextMatrix(0, 6) = "Fec.Actu."
  msf1.TextMatrix(0, 7) = "Proveedor"
  
 For i = 0 To 6
   msf1.ColAlignment(i) = 9 'izq
 Next i
  
 msf1.ColAlignment(1) = 1 'izq
 msf1.ColAlignment(7) = 1 'izq

End Sub







Private Sub c_prov_LostFocus()
If c_prov.ListIndex <= 0 Then
  c_prov.ListIndex = 0
End If
End Sub





Private Sub Form_Load()
t_dolar = Format$(para.cotizacion, "#####0.000")
Call carga_piezas(c_prov)
c_prov.AddItem "<seleccionar Pieza>", 0
c_prov.ListIndex = 0

Call armagrid

Option1 = True


End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = " [F7] Imprime -  [F11] Excel - [ENTER] Modif. Celda "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double
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
      Call imprimegrid(msf1, c(), "PLANILLA COSTOS", "", "Pieza: " & c_prov, " ", 55, 9, True, False, "H")
    End If
         
  End If
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Function busca_pieza_en_estructura(ByVal p As Long, ByVal e As Long) As Boolean
'BUSCA UN PIEZA P DENTRO DE UNA ESTRUCTURA E
'VERDADERO SI LO ENCUENTRA SINO FALSO
Y = False
q = "SELECT * FROM PRO_07 WHERE [ID_pieza] = " & p & " and [tipo] = 'P'"
Set rs = New ADODB.Recordset
rs.Open q, cn1
o = 0
While Not rs.EOF And o = 0
   If rs("ID") = e Then
     'LA ENCONTRO
     Y = True
     o = 1
   Else
     If busca_pieza_en_estructura(rs("id"), e) Then
         Y = True
         o = 1
     End If
   End If
   rs.MoveNext
Wend
busca_pieza_en_estructura = Y
End Function

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.col = 2 Or msf1.col = 3 Then
    If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
      
      d = InputBox$("ingreso de datos", "MODIFICACION TABLA DE COSTOS")
      If Val(d) >= 0 Then
        msf1.TextMatrix(msf1.Row, msf1.col) = Format(Val(d), "#####0.00")
        Call RECALCULA
      End If
    End If
  End If
End If
End Sub
Sub RECALCULA()
J = 1
While J <= msf1.Rows - 1
 If Val(msf1.TextMatrix(J, 0)) > 0 Then
    msf1.TextMatrix(J, 4) = Format$(Val(msf1.TextMatrix(J, 2)) * Val(msf1.TextMatrix(J, 3)), "######0.00")
 End If
 J = J + 1
Wend
msf1.TextMatrix(msf1.Rows - 1, 4) = ""
tot = Format$(suma_msflexgrid(msf1, 4), "######0.00")
msf1.TextMatrix(msf1.Rows - 1, 4) = tot

End Sub
Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub



Private Sub t_dolar_LostFocus()
If Val(t_dolar) < 1 Then
  t_dolar = Format$(para.cotizacion, "#####0.000")
Else
  t_dolar = Format$(Val(t_dolar), "#####0.000")
End If

  

End Sub
