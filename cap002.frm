VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form com_faltantes 
   BackColor       =   &H00E0E0E0&
   Caption         =   "REGISTRO DE FALTANTES"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtros"
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   11535
      Begin VB.TextBox t_desc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   25
         TabIndex        =   15
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   16
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7920
         Picture         =   "cap002.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Desc. Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Basico:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4560
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Provedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
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
         Picture         =   "cap002.frx":0372
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
         Picture         =   "cap002.frx":0BF4
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
      Width           =   11910
      _ExtentX        =   21008
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
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "31/03/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:59 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_faltantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  
  Call armagrid
  t2 = 0
  q = "select * from a5, a6, a1  where [id_tipocomp] = 60 and a5.[num_int] = a6.[num_int]  and [envase] = a1.[id_proveedor]"
  c = " and "
  
  If c_prov.ListIndex > 0 Then
    q = q & c & " [envase] = " & c_prov.ItemData(c_prov.ListIndex)
   
  End If
  If t_fecha <> "" Then
        q = q & c & " datevalue(a6.[fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" Then
        q = q & c & " datevalue(a6.[fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
  If t_basico <> "" Then
    If Len(t_basico) <= 5 Then
       q = q & c & " [id_producto] = " & Val(t_basico)
    Else
       ip = codproddesdebarras(t_basico)
       q = q & c & " [id_producto] = " & ip
    End If
  End If
  
  If t_descprod <> "" Then
       q = q & c & " [descripcion] like %'" & RTrim$(t_descprod) & "%'"
  End If
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
  While Not rs.EOF
       msf1.AddItem rs("a6.fecha") & Chr(9) & rs("id_producto") & Chr(9) & rs("detalle") & Chr(9) & rs("cantidad") & Chr(9) & rs("unidad") & Chr(9) & rs("a1.Id_proveedor") & Chr(9) & rs("denominacion") & Chr(9) & rs("a6.num_int") & Chr(9) & rs("renglon")
       rs.MoveNext
  Wend
  
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
  msf1.Cols = 10
  msf1.ColWidth(0) = 1000
  msf1.ColWidth(1) = 800
  msf1.ColWidth(2) = 3500
  msf1.ColWidth(3) = 1000
  msf1.ColWidth(4) = 600
  msf1.ColWidth(5) = 500
  msf1.ColWidth(6) = 3000
  msf1.ColWidth(7) = 1500
  msf1.ColWidth(8) = 800
  msf1.TextMatrix(0, 0) = "Fecha Ing."
  msf1.TextMatrix(0, 1) = "Basico"
  msf1.TextMatrix(0, 2) = "Producto"
  msf1.TextMatrix(0, 3) = "Cantidad"
  msf1.TextMatrix(0, 4) = "Unidad"
  msf1.TextMatrix(0, 5) = "Id."
  msf1.TextMatrix(0, 6) = "Proveedor"
  msf1.TextMatrix(0, 7) = "Num.Int"
  msf1.TextMatrix(0, 8) = "Reng."
  
  For i = 0 To 3
    msf1.ColAlignment(i) = 1
  Next i
  msf1.ColAlignment(4) = 9
  For i = 5 To 8
    msf1.ColAlignment(i) = 1
  Next i

  
  
End Sub






Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If
End Sub


Private Sub Command5_Click()
com_proveedor.t_id = c_prov.ItemData(c_prov.ListIndex)
com_proveedor.carga
com_proveedor.Show

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
    Call TabEnter2(Me, 2)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call carga_proveedores(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0
Call armagrid
Call barraesag(Me)

Load com_faltantes1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload com_faltantes1
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F8] Elimina  - [F11] Excel - [INS] Agrega  "
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
     
     Call imprimegrid(msf1, c(), "REGISTRO DE FALTANTES", "", "Provedor: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)

    End If
         
  End If
  
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF8 Then
 If Val(msf1.TextMatrix(msf1.Row, 8)) > 0 Then
  J = MsgBox("Confirma Eliminar articulo [" & msf1.TextMatrix(msf1.Row, 3) & "] del Registro de Faltantes", 4)
  If J = 6 Then
    Set cl_prod = New productos
    Call cl_prod.sacafaltante(Val(msf1.TextMatrix(msf1.Row, 8)))
    Call carga
  End If
 End If
End If

If KeyCode = vbKeyInsert Then
  com_faltantes1.t_renglon = 0
  com_faltantes1.Show
 
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub


Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_desc_GotFocus()
t_desc = ""
End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
  
End Sub
