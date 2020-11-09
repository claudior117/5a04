VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_listaprecios2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA DE PRECIOS"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Ordenado por"
      Height          =   615
      Left            =   7560
      TabIndex        =   21
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   7560
      TabIndex        =   18
      Top             =   0
      Width           =   2175
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00008000&
         Caption         =   "Encontrados"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox c_prov 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.ComboBox c_marca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.ComboBox c_depto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox t_detalle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5280
         MaxLength       =   13
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800080&
         Caption         =   "Marca"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Departamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Grupo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Cod. Barra"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   5175
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   11655
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   635
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
            TextSave        =   "09/02/2006"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:27 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_listaprecios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga()
 ct = Space$(10)
 List1.Clear
 Call cabecera
 Set rs = New ADODB.Recordset
 q = "select * from a2 "
 c = " where "
 
 If t_basico <> "" Then
   q = q & c & "[id_producto] = " & Val(t_basico)
   c = " and "
 End If
 
 If t_codbarra <> "" Then
   If Len(t_codbarra) = 13 Then
     s = " = "
   Else
     s = " >= "
   End If
   q = q & c & "[cod_barra] " & s & Val(t_codbarra)
   c = " and "
 End If
 
 If t_detalle <> "" Then
   q = q & c & "[descripcion] like  '%" & t_detalle & "%'"
   c = " and "
 End If
 
 If c_grupo.ListIndex > 0 Then
   q = q & c & "[id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
   c = " and "
 End If
 
 If c_depto.ListIndex > 0 Then
   q = q & c & "[id_depto] = " & c_depto.ItemData(c_depto.ListIndex)
   c = " and "
 End If
 
  If c_marca.ListIndex > 0 Then
   q = q & c & "[id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
   c = " and "
 End If
 
  If c_prov.ListIndex > 0 Then
   q = q & c & "[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
   c = " and "
 End If
 
 If Option1 = True Then
   q = q & " order by [id_producto]"
 Else
   q = q & " order by [descripcion]"
 End If
 rs.Open q, cn1
 p = Space$(10)
 c = Space$(4)
 t_encontrados = 0
 While Not rs.EOF
    b = Format$(rs("id_producto"), "00000")
    d = Format$(Left$(rs("descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    RSet p = Format$(rs("precio_final"), "######0.00")
    RSet c = Format$(rs("stock"), "###0")
    
    List1.AddItem b & "  " & d & "  " & p & "  " & c
    t_encontrados = Val(t_encontrados) + 1
    rs.MoveNext
 Wend
List1.SetFocus


End Sub

Sub cabecera()
lc = "-----------------------------------------------------------------------------------------------------------------------"
List1.AddItem "LISTA DE PRECIOS"
List1.AddItem ""
List1.AddItem "Grupo:   " & Format$(c_grupo.ItemData(c_grupo.ListIndex), "0000") & " " & c_grupo
List1.AddItem "Depto:   " & Format$(c_depto.ItemData(c_depto.ListIndex), "0000") & " " & c_depto
List1.AddItem "Marca:   " & Format$(c_marca.ItemData(c_marca.ListIndex), "0000") & " " & c_marca
List1.AddItem ""
List1.AddItem ""
List1.AddItem lc
List1.AddItem "Basico  Descripcion                            P.Final  Stock"
List1.AddItem lc


End Sub

Private Sub c_depto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_prov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If

End Sub

Private Sub Form_Load()
  Call carga_grupos(c_grupo)
  c_grupo.AddItem "<Todos>", 0
  c_grupo.ListIndex = 0
  Call carga_deptos_venta(c_depto)
  c_depto.AddItem "<Todos>", 0
  c_depto.ListIndex = 0
  Call carga_marcas(c_marca)
  c_marca.AddItem "<Todas>", 0
  c_marca.ListIndex = 0
  Call carga_proveedores(c_prov)
  c_prov.AddItem "<Todos>", 0
  c_prov.ListIndex = 0
  Option2 = True
End Sub

  




Private Sub List1_LostFocus()
List1.ListIndex = -1
End Sub


Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_codbarra_GotFocus()
t_codbarra = ""
End Sub

Private Sub t_codbarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_detalle_GotFocus()
t_detalle = ""
End Sub

Private Sub t_detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub
