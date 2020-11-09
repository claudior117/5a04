VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form pro_estructura 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ESTRUCTURA DE PRODUCTOS"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10398
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
         BackColor       =   &H00800080&
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
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "pro007.frx":0000
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
         Picture         =   "pro007.frx":0882
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tipo: [B] Basico (Item Lista Precios)    -   [P] pieza (Item estructura producto)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   9135
   End
End
Attribute VB_Name = "pro_estructura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double

Sub carga()
  espere.Show
  espere.Label1 = "Espere... cargand estructura producto"

  Call armagrid
  q = "select * from pro_07 where [id_pieza] = " & c_prov.ItemData(c_prov.ListIndex) & " order by [renglon]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
    'If rs("tipo") = "P" Then
      'busco pieza en estructura
    '   Call buscaestructura(ip)
       
   ' Else
      'busco producto en lista de prcios
       msf1.AddItem rs("renglon") & Chr(9) & rs("detalle") & Chr(9) & rs("cantidad") & Chr(9) & rs("Unidad") & Chr(9) & rs("id") & Chr(9) & rs("tipo")
    'End If
    rs.MoveNext
  Wend
  Unload espere
End Sub
Sub buscaestructura(ip)
  q = "select * from pro_07 where [id_pieza] = " & ip & " order by [renglon]"
  Set rs1 = New ADODB.Recordset
  rs1.Open q, cn1
  While Not rs1.EOF
    If rs1("tipo") = "P" Then
      'busco pieza en estructura
       Call buscaestructura(ip)
       
    Else
      'busco producto en lista de prcios
       msf1.AddItem "" & Chr(9) & "  " & rs1("detalle") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("Unidad") & Chr(9) & rs("id") & Chr(9) & rs("tipo")
    End If
    rs1.MoveNext
  Wend
  

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 6
  msf1.ColWidth(0) = 600
  msf1.ColWidth(1) = 7000
  msf1.ColWidth(2) = 1500
  msf1.ColWidth(3) = 2000
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.TextMatrix(0, 0) = "Renglon"
  msf1.TextMatrix(0, 1) = "DETALLE"
  msf1.TextMatrix(0, 2) = "Cantidad"
  msf1.TextMatrix(0, 3) = "Unidad"
  msf1.TextMatrix(0, 4) = "Id."
  msf1.TextMatrix(0, 5) = "Tipo"
  
End Sub







Private Sub c_prov_LostFocus()
If c_prov.ListIndex <= 0 Then
  c_prov.ListIndex = 0
Else
  Call carga
End If
End Sub





Private Sub Form_Load()

Call carga_piezas(c_prov)
c_prov.AddItem "<seleccionar Pieza>", 0
c_prov.ListIndex = 0

Call armagrid

Option1 = True

Load pro_estructura1
Load pro_estructura2
End Sub
Sub renumera()
For i = 1 To msf1.Rows - 1
  msf1.TextMatrix(i, 0) = i
Next i


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload pro_estructura1
Unload pro_estructura2
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[INS]Agrega prod. -  [F2] Agrega pieza - [F5] Saca reng. -[ENTER] Modif. reng. - [F7] Imprime - [F9] Graba -  [F11] Excel  "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 k = MsgBox("Confirma grabar estructura para esta pieza", 4)
 If k = 6 Then
  If verifica Then
    Call graba
  End If
 End If
End If


If KeyCode = vbKeyInsert Then
  pro_estructura1.t_renglon = ""
  pro_estructura1.t_cantidad = ""
  pro_estructura1.t_unidad = ""
  pro_estructura1.Show
End If


If KeyCode = vbKeyF2 Then
  pro_estructura2.t_renglon = ""
  pro_estructura2.t_cantidad = ""
  pro_estructura2.t_unidad = ""
  pro_estructura2.Show
End If

If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
    Call renumera
  Else
   Call armagrid
   
 End If
End If


If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double

     If Check1 = 0 Then
      c(0) = 10
      c(1) = 0
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      For i = 8 To 14
        c(i) = -1
      Next i
    Else
      
      c(0) = 11
      c(1) = 0
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      c(8) = 8
      
      For i = 9 To 14
        c(i) = -1
      Next i
     End If
     
     If Check2 = 0 Then
        Call imprimegrid(msf1, c(), "ESTADO DE CUENTA", "", "Cliente: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)
     Else
        Call imprimegrid(msf1, c(), "ESTADO DE CUENTA", "", "Cliente: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 50, 9, True, False, "H")
     End If
    End If
         
  End If
  
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub
Function verifica() As Boolean
'esta funcion busca que la pieza para la que secrea la estructura
'no est dentr de otra structura lo que prvocaria un bucle sin fin
'cuando se calcula el costo

v = True
J = 1
While J <= msf1.Rows - 1
  If msf1.TextMatrix(J, 5) = "P" Then 'pieza
    If busca_pieza_en_estructura(Val(msf1.TextMatrix(J, 4)), c_prov.ItemData(c_prov.ListIndex)) Then
       MsgBox ("ERROR!!!! La pieza se encuentra dentro de otra estructura agregada a si misma")
       v = False
       J = msf1.Rows
    End If
  End If
  J = J + 1
Wend
verifica = v
End Function

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
  If msf1.Row > 0 Then
    If msf1.TextMatrix(msf1.Row, 5) = "B" Then
      pro_estructura1.t_renglon = msf1.Row
      pro_estructura1.t_basico = msf1.TextMatrix(msf1.Row, 4)
      pro_estructura1.t_detalle = msf1.TextMatrix(msf1.Row, 1)
      pro_estructura1.t_cantidad = msf1.TextMatrix(msf1.Row, 2)
      pro_estructura1.t_unidad = msf1.TextMatrix(msf1.Row, 3)
      pro_estructura1.Show
    Else
    'mofific pieza
      pro_estructura2.t_renglon = msf1.Row
      pro_estructura2.c_pieza.ListIndex = buscaindice(pro_estructura2.c_pieza, Val(msf1.TextMatrix(msf1.Row, 4)))
      pro_estructura2.t_cantidad = msf1.TextMatrix(msf1.Row, 2)
      pro_estructura2.t_unidad = msf1.TextMatrix(msf1.Row, 3)
      pro_estructura2.Show
    End If
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Sub graba()
 If c_prov.ListIndex > 0 Then
   espere.Show
   espere.Label1 = "Espere...  Guardando estructura de productos"
   If msf1.Rows > 0 Then
      Call borra
      Call graba2
   Else
     J = MsgBox("La estructura para esta pieza esta vacia, confirma eliminar estructura", 4)
     If J = 6 Then
       Call borra
     End If
   End If
   Unload espere
 Else
   MsgBox ("Debe selecionar una pieza para poder grabar estructura")
 End If
End Sub

Sub borra()
Set rs = New ADODB.Recordset
q = "select * from pro_07 where [id_pieza] = " & c_prov.ItemData(c_prov.ListIndex)
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = Nothing

End Sub

Sub graba2()
Set rs = New ADODB.Recordset
q = "select * from pro_07 where [id_pieza] = " & c_prov.ItemData(c_prov.ListIndex)
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
J = 1
While J <= msf1.Rows - 1
  rs.AddNew
    rs("id_pieza") = c_prov.ItemData(c_prov.ListIndex)
    rs("renglon") = Val(msf1.TextMatrix(J, 0))
    rs("id") = Val(msf1.TextMatrix(J, 4))
    rs("tipo") = msf1.TextMatrix(J, 5)
    rs("cantidad") = Format(Val(msf1.TextMatrix(J, 2)), "#####0.00")
    rs("unidad") = msf1.TextMatrix(J, 3)
    rs("detalle") = msf1.TextMatrix(J, 1)
   rs.Update
   J = J + 1
Wend
Set rs = Nothing
MsgBox ("Estructura Guardada")
c_prov.ListIndex = 0
Call armagrid

Set rs = Nothing
End Sub
