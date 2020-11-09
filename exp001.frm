VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form exp_exporta1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPERACIONES DE EXPORTACION"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3795
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_importe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   14
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox t_fechaf 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox t_fechap 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox c_cli 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   1
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Importe Op.:"
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
         TabIndex        =   19
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Factura:"
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
         TabIndex        =   18
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Embarque:"
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
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
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
         Height          =   495
         Left            =   480
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "id. Operacion"
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
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Detalle:"
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
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9840
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "exp001.frx":0000
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
         Picture         =   "exp001.frx":0882
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
      Top             =   3540
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:40"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "exp_exporta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnacepta_Click()
If verifica Then
  Call graba
Else
  MsgBox ("No se actualizaron los Datos")
End If
End Sub

Function verifica() As Boolean
v = True

If t_descripcion = "" Then
   MsgBox ("La descripcion de la operacion no puede estar en Blanco")
   v = False
End If

Set cl_cli = New Clientes
cl_cli.carga (c_cli.ItemData(c_cli.ListIndex))
If cl_cli.id > 0 Then
   If cl_cli.idtipoiva <> 8 Then
      MsgBox ("El cliente no esta marcado como de Exportacion")
      v = False
   End If
Else
  MsgBox ("Error en la carga del Cliente")
  v = False
End If
Set cl_cli = Nothing

  
If t_fechap <> "" Then
  If Not IsDate(t_fechap) Then
     MsgBox ("Error en la fecha de Pedido")
      v = False
  End If
Else
  MsgBox ("La fecha de pedido no puede estar en Blanco")
  v = False
End If

If t_fechaf <> "" Then
  If Not IsDate(t_fechaf) Then
     MsgBox ("Error en la fecha de factura")
      v = False
  End If
Else
  t_fechaf = t_fechap
End If


verifica = v

End Function

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   
   On Error GoTo ERRORGRABA
    
   Select Case t_funcion
      
   Case "A"
      QUERY = "INSERT INTO exp01([id_cliente], [fecha_embarque], [fecha_fact], [detalle], [importe], [cliente], [estado])"
      QUERY = QUERY & " VALUES (" & c_cli.ItemData(c_cli.ListIndex) & ", '" & t_fechap & "', '" & t_fechaf & "', '" & t_descripcion & "', " & Val(t_importe) & ", '" & Left$(c_cli, 50) & "', 'E')"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   Case "M"
   
      QUERY = "update exp01 set  [fecha_embarque]='" & t_fechap & "', [fecha_fact]='" & t_fechaf & "', [detalle]='" & t_descripcion & "', [importe]=" & Val(t_importe) & ", [cliente]='" & c_cli & "', [id_cliente]=" & c_cli.ItemData(c_cli.ListIndex)
      QUERY = QUERY & " where [num_exp]= " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
      
      
   Case "B"
         
      J = MsgBox("Si elimina la operacion se eliminaran los productos asignados a reintegro. Confirma", 4)
      If J = 6 Then
        q = "select * from exp02 where [num_exp] = " & Val(t_id)
        Set rs = New ADODB.Recordset
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs.EOF
           Set rs2 = New ADODB.Recordset
           q = "select * from a6 where [num_int] = " & rs("num_int_c") & " and [renglon] = " & rs("renglon_c")
           rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
           rs2.MaxRecords = 1
           If Not rs2.EOF And Not rs2.BOF Then
              rs2("exportacion") = rs2("exportacion") - rs("cantidad")
              rs2.Update
           End If
           Set rs2 = Nothing
           
           rs.Delete
           rs.MoveNext
        Wend
        Set rs = Nothing
        
        q = "select * from exp01 where [num_exp] = " & Val(t_id)
        Set rs = New ADODB.Recordset
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.EOF And Not rs.BOF Then
          rs.Delete
        End If
        Set rs = Nothing
     End If
   End Select
   
   
   exp_exporta.Show
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






Private Sub c_cli_LostFocus()
If c_cli.ListIndex < 0 Then
  c_cli.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
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

Call carga_clientes(c_cli)
c_cli.ListIndex = 0
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub



Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub
