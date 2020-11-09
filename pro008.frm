VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form pro_estructura1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "INGRESA PRODUCTO A ESTRUCTURA"
   ClientHeight    =   2175
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10440
         MaxLength       =   8
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         MaxLength       =   69
         TabIndex        =   1
         Top             =   720
         Width           =   7935
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   13
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10320
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   8535
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   11880
      _ExtentX        =   20955
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
            TextSave        =   "21/06/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:19 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "pro_estructura1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub Form_Activate()

t_basico.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub t_basico_GotFocus()
t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Sub carga()
If IsNumeric(t_basico) Then
 If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_unidad = "U."
    t_detalle.Enabled = True
    t_detalle.SetFocus
    
 Else
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
 End If
Else
  Call busca("B") 'busca por cod. barra
End If
End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2, g5 where a2.[id_unidad] = g5.[id_unidad]"
If tipo = "I" Then
  q = q & " and [id_producto] = " & Val(t_basico)
Else
  q = q & " and [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  If para.tipoprecioventa = 1 Then
    t_pu = rs("precio_final")
  Else
    t_pu = rs("pu")
  End If
  t_ip = rs("id_producto")
  t_unidad = rs("unidad")
 
 

Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub

Private Sub t_cantidad_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F5] Comvierte Unidades x Envase"
End Sub

Private Sub t_cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
  c = InputBox$("Ingrese Cantidad a Convertir (formula:cantidad x envase)")
  If Val(c) > 0 Then
     t_cantidad = Format$(Val(c) * Val(t_envase), "#####0.00")
  End If
End If
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Call solonum(KeyAscii, 1)

End If
End Sub

Sub cargarenglon(t As String)
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  u = Left$(t_unidad, 8)
  ip = Val(t_basico)
  If t = "A" Then
    r = pro_estructura.msf1.Rows
    pro_estructura.msf1.AddItem r & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr(9) & ip & Chr(9) & "B"
  Else
    r = t_renglon
    pro_estructura.msf1.AddItem r & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr(9) & ip & Chr(9) & "B", r
    pro_estructura.msf1.RemoveItem r + 1
  End If
  para.producto_sel = 0
End Sub
 
  
Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_unidad = ""
t_renglon = ""
End Sub

Private Sub t_cantidad_LostFocus()
If vta_remitos.c_tipocomp.ItemData(vta_remitos.c_tipocomp.ListIndex) = 46 Then
 If Val(t_basico) > 1 Then
  If Val(t_cantidad) > Val(t_tr) Then
    MsgBox ("El cliente tiene " & (t_tr) & " unidades del producto seleccionado sin facturar. La cantidad a devolver no puede superar esa cantidad")
    t_cantidad.SetFocus
  End If
 Else
  tr = ""
 End If
End If

End Sub


Sub pasa()
 If t_renglon = "" Then
   Call cargarenglon("A")
   t_basico.SetFocus
   
  Else
   Call cargarenglon("M")
   Me.Hide
  End If
  Call limpia
  
End Sub

Private Sub t_unidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Call pasa
End If

End Sub
