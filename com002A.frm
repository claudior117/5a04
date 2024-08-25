VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form com_config_comp1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIGURA COMPROBANTES COMPRAS"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTANTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   9240
      TabIndex        =   37
      Top             =   1320
      Width           =   2655
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1080
         Picture         =   "com002A.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   $"com002A.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Index           =   22
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "¡Atencion!"
      Height          =   1815
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Width           =   9015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   19
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4815
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   9015
      Begin VB.TextBox t_ie 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   9
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox t_sucursal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_formato2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   10
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox t_contab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   8
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox t_venta 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   7
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox t_iva 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   6
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox t_ctacte 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox t_stock 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   4
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox t_copiasA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox t_ultimoA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox t_abrevia 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   0
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label9 
         Caption         =   "[S] Si  [N] No"
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   3960
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Imprime desc. extra"
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
         Left            =   120
         TabIndex        =   41
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Sucursal:"
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
         Left            =   5160
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Formato:"
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
         Index           =   21
         Left            =   120
         TabIndex        =   36
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "[D] Debe   [H] Haber  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Contabilidad:"
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
         Index           =   15
         Left            =   120
         TabIndex        =   33
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "[S] Suma   [R] Resta  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   3240
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Inf. Compra:"
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
         Index           =   14
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "[S] Suma   [R] Resta  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "[D] Debe   [H] Haber  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "[E] Entrada   [S] Salida  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Iva Vta como"
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
         Index           =   13
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra CtaCte como:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Stock como:"
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
         Index           =   11
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cantidad Copias"
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
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Ultimo Num. Usado"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Abreviatura"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id."
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Comprobante"
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
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   12
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "com002A.frx":039B
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "com002A.frx":0C1D
         Style           =   1  'Graphical
         TabIndex        =   13
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
      TabIndex        =   11
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
            TextSave        =   "25/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:13 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_config_comp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private cambio As String


Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   'On Error GoTo ERRORGRABA
    
   Select Case t_funcion
   
   Case "M"
             
      QUERY = "update g2 set  [Descripcion]='" & t_descripcion & "' , [abreviatura]='" & t_abrevia & "' , [ult_num]= " & Val(t_ultimoA) & " , [copias]= " & Val(t_copiasA) & " , [stock]='" & t_stock & "' , [ctacte]='" & t_ctacte & "' , [iva]='" & t_iva & "' , [compra]='" & t_venta & "' , [contabilidad]='" & t_contab & _
      "' , [formato]='" & t_formato2 & "' , [imprime_desc_extra]='" & t_ie & "'"
      QUERY = QUERY & " where [id_tipo_comp] = " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      
      
       QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
       QUERY = QUERY & " VALUES ('Modifica configuracion comprobante:" & t_id & "', " & para.id_usuario & ", 'C', " & Val(t_id) & ", '" & Now & "', '[" & t_abrevia & "', 106, 1)"
       cn1.Execute QUERY
      
      
      
      cn1.CommitTrans
   
   
   End Select
   
   com_config_comp.DataGrid1.Refresh
   com_config_comp.Show
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
    Call TabEnter2(Me, 10)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

End Sub


Private Sub t_contab_GotFocus()
cambio = t_contab
End Sub

Private Sub t_contab_LostFocus()
t_contab = Format$(t_contab, ">@")
If t_contab <> "D" And t_contab <> "H" And t_contab <> "N" Then
   t_contab = cambio
End If

End Sub

Private Sub t_ctacte_GotFocus()
cambio = t_ctacte
End Sub

Private Sub t_ctacte_LostFocus()
t_ctacte = Format$(t_ctacte, ">@")
If t_ctacte <> "D" And t_ctacte <> "H" And t_ctacte <> "N" Then
   t_ctacte = cambio
End If

End Sub

Private Sub t_descripcion_LostFocus()
If t_descripcion = "" Then
  t_descripcion = "Null"
End If
End Sub


Private Sub t_formato2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_formato2_LostFocus()
If t_formato2 = "" Then
  t_formato2 = "1"
End If
End Sub


Private Sub t_ie_LostFocus()
t_ie = Format$(t_ie, ">@")
If t_ie <> "S" And t_ie <> "N" Then
   t_ie = "N"
End If

End Sub

Private Sub t_iva_GotFocus()
cambio = t_iva
End Sub

Private Sub t_iva_LostFocus()
t_iva = Format$(t_iva, ">@")
If t_iva <> "S" And t_iva <> "R" And t_iva <> "N" Then
   t_iva = cambio
End If

End Sub



Private Sub t_stock_GotFocus()
cambio = t_stock
End Sub

Private Sub t_stock_LostFocus()
t_stock = Format$(t_stock, ">@")
If t_stock <> "E" And t_stock <> "S" And t_stock <> "N" Then
   t_stock = cambio
End If

End Sub

Private Sub t_venta_GotFocus()
cambio = t_venta
End Sub

Private Sub t_venta_LostFocus()
t_venta = Format$(t_venta, ">@")
If t_venta <> "S" And t_venta <> "R" And t_venta <> "N" Then
   t_venta = cambio
End If
End Sub


Private Sub Text1_Change()

End Sub

Private Sub Text1_LostFocus()
End Sub
