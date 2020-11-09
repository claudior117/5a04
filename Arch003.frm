VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_solmat1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOLICITUD DE MATERIALES"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de la Requisicion"
      Height          =   615
      Left            =   8640
      TabIndex        =   29
      Top             =   2760
      Width           =   3135
      Begin VB.TextBox t_requerido 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   13
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Requerido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informacion del Producto Seleccionado"
      Height          =   1455
      Left            =   8640
      TabIndex        =   22
      Top             =   1080
      Width           =   3135
      Begin VB.TextBox t_enreq 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   13
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox t_enstock 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   13
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_enoc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   13
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "En Requisicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "En Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "En O.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   14
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_obs 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   5
         Top             =   2520
         Width           =   5895
      End
      Begin VB.TextBox t_fechauso 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox C_obra 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   5895
      End
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox t_cantidad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   21
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id. Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   7
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "Arch003.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Arch003.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   8
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
      TabIndex        =   6
      Top             =   8265
      Width           =   11910
      _ExtentX        =   21008
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
            TextSave        =   "22/12/2005"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:55 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "abm_solmat1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EXISTE As String



Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
j = MsgBox("Confirma Valores para Grabar", 4)
If j = 6 Then
   
   'valida datos
   If c_prod.ListIndex <= 0 Then
      cp = 1
   Else
      cp = c_prod.ItemData(c_prod.ListIndex)
   
   End If
   If Val(t_cantidad) <= 0 Then
      t_cantidad = 1
   End If
   t_cantidad = Format$(Val(t_cantidad), "#####0.00")

   
     t_fecha = Format$(Now(), "dd/mm/yyyy")
      
   If Not IsDate(t_fechauso) Then
     t_fechauso = Format$(Now(), "dd/mm/yyyy")
   Else
     t_fechauso = Format$(t_fechauso, "dd/mm/yyyy")
   End If
   
   
   On Error GoTo ERRORGRABA
   
   Select Case t_funcion
      
   Case "A"
      query = "INSERT INTO a3([id_producto], [detalle], [cantidad_requisicion], [id_obra], [id_usuario], [fecha_requisicion], [hora_requisicion], [fecha_esperado], [estado], [observaciones], [num_int_oc], [fecha_emision_oc], [fecha_prob_entrega], [num_int_comp_compra], [fecha_recepcion], [cant_pedida], [cant_recibida])"
      query = query & " VALUES (" & cp & ", '" & c_prod & "', " & Val(t_cantidad) & ", " & C_obra.ItemData(C_obra.ListIndex) & ", " & para.id_usuario & ", '" & t_fecha & "', '" & Format$(Now, "HH:MM:SS") & "', '" & t_fechauso & "', 'R', '" & t_obs & "', 0, '01/01/2000', '01/01/2000', 0, '01/01/2000', 0, 0" & ")"
      cn1.BeginTrans
      cn1.Execute query
      cn1.CommitTrans
     Set cl_prod = New productos
     Call cl_prod.actualizar(cp, 0, Val(t_cantidad), 0)
     Set cl_prod = Nothing
     
   Case "M"
   
      query = "update a3 set  [id_producto]=" & cp & " , [detalle]='" & c_prod & "' , [cantidad_requisicion]=" & Val(t_cantidad) & " , [id_obra]=" & C_obra.ItemData(C_obra.ListIndex) & " , [id_usuario]=" & para.id_usuario & " , [fecha_requisicion]='" & t_fecha & "' , [hora_requisicion]='" & Format$(Now, "HH:MM:SS") & "' , [fecha_esperado]='" & t_fechauso & "' , [observaciones]='" & t_obs & "'"
      query = query & " where [id_renglon] = " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute query
      cn1.CommitTrans
      
      Set cl_prod = New productos
      Call cl_prod.actualizar(cp, 0, Val(t_cantidad), 0)
      Call cl_prod.actualizar(cp, 0, -Val(t_requerido), 0)
      Set cl_prod = Nothing
      
      
   Case "B"
      query = "DELETE FROM a3 WHERE [id_renglon] = " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute query
      cn1.CommitTrans
   
      Set cl_prod = New productos
      Call cl_prod.actualizar(cp, 0, -Val(t_cantidad), 0)
      Set cl_prod = Nothing
   
   End Select
   
   ABM_SOLMAT.DataGrid1.Refresh
   ABM_SOLMAT.Show
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub



Private Sub c_prod_LostFocus()
If c_prod.ListIndex < 0 Then
  If c_prod = "" Then
    c_prod.ListIndex = 0
  End If
  t_stock = ""
Else
  Set cl_prod = New productos
  cl_prod.cargar (c_prod.ItemData(c_prod.ListIndex))
  If cl_prod.idproducto > 0 Then
    t_enstock = cl_prod.stock
    t_enoc = cl_prod.pedido
    t_enreq = cl_prod.requerido
  Else
   t_enstock = 0
    t_enoc = 0
    t_enreq = 0
  
  End If
  Set cl_prod = Nothing
End If
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  
  c_prod.SetFocus
  C_obra.ListIndex = buscaindice(C_obra, para.id_obraactual)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_obras(C_obra)
Call carga_productos(c_prod)
c_prod.ListIndex = 0

End Sub


Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_cantidad_LostFocus()
If Val(t_cantidad) <= 0 Then
  t_cantidad = "1"
End If
If t_funcion = "M" Then
  If Val(t_cantidad) < Val(t_cantp) Then
    MsgBox ("La cantidad a requerir no puede ser inferior a la cantidad Ingresada")
  Else
    t_cantidad = Format$(Val(t_cantidad), "#####0.00")
  End If
Else
    t_cantidad = Format$(Val(t_cantidad), "#####0.00")
End If
End Sub





Private Sub t_fecha_GotFocus()
If t_fecha = "" Then
  t_fecha = Format$(Now(), "dd/mm/yyyy")
End If
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now(), "dd/mm/yyyy")
Else
    t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If

End Sub

Private Sub t_fechauso_GotFocus()
If t_fechauso = "" Then
  t_fechauso = Format$(Now(), "dd/mm/yyyy")
End If

End Sub

Private Sub t_fechauso_LostFocus()
If Not IsDate(t_fecha) Then
  t_fechauso = Format$(Now(), "dd/mm/yyyy")
Else
  t_fechauso = Format$(t_fechauso, "dd/mm/yyyy")
End If

End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub
