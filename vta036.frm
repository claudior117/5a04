VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_fact_viaje1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE FLETES"
   ClientHeight    =   2175
   ClientLeft      =   1080
   ClientTop       =   4740
   ClientWidth     =   13890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   13890
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   13815
      Begin VB.TextBox t_cartaporte 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7440
         MaxLength       =   49
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_destino 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6000
         MaxLength       =   49
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_origen 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   4560
         MaxLength       =   49
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_chofer 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2760
         MaxLength       =   49
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox t_costo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_kmts 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   9480
         MaxLength       =   5
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1080
         MaxLength       =   49
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   12720
         MaxLength       =   11
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox c_tasa 
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
         Left            =   11400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_pu 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8520
         MaxLength       =   8
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "C.Porte"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7320
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Destino"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5880
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Origen"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Chofer/camion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Kmts"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   12480
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Tasa Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11160
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Tarifa"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10200
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Toneladas"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "03/03/2011"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:58 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_fact_viaje1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub c_tasa_LostFocus()
If c_tasa.ListIndex < 0 Then
  c_tasa.ListIndex = 0
End If
End Sub



Private Sub Form_Activate()

t_fecha.SetFocus

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
    Call TabEnter2(Me, 10)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

For i = 0 To 9
  c_tasa.AddItem para.tasaiva(i)
Next i
c_tasa.ListIndex = 1
End Sub


Private Sub t_fecha_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Acepta - [ESC] Sale  "

End Sub








Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
 If Not IsDate(t_fecha) Then
   t_fecha = Format$(Now, "dd/mm/yyyy")
 End If
Else
 t_fecha = Format$(Now, "dd/mm/yyyy")
End If
 
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub

Sub cargarenglon(t As String)
  
  f = t_fecha
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  ti = Format$(c_tasa, "####0.00")
  If para.tipoprecioventa = 1 Then
   pu = Format$(Val(t_pu) / (1 + Val(c_tasa) / 100), "#####0.000")
   im = Format$(Val(pu) * Val(cu), "#####0.00")
   puf = Format$(Val(t_pu), "#####0.00")
  Else
   pu = Format$(Val(t_pu), "#####0.000")
   im = Format$(Val(pu) * Val(cu), "#####0.00")
   puf = Format$(Val(t_pu) * (1 + Val(c_tasa) / 100), "#####0.000")
  End If
  
  pt = Format$((Val(puf) * Val(cu)), "######0.00")
  iv = Format$(Val(pt) - Val(im), "######0.00")
  
  If t = "A" Then
    r = vta_fact_viaje.msf1.Rows
    If r < 25 Then
       vta_fact_viaje.msf1.AddItem r & Chr(9) & f & Chr(9) & d & Chr(9) & t_chofer & Chr(9) & t_origen & Chr$(9) & t_destino & Chr(9) & t_cartaporte & Chr(9) & cu & Chr(9) & t_kmts & Chr(9) & pu & Chr(9) & ti & Chr$(9) & im & Chr$(9) & puf & Chr$(9) & iv & Chr$(9) & pt
    End If
  Else
    r = t_renglon
    vta_fact_viaje.msf1.AddItem r & Chr(9) & f & Chr(9) & d & Chr(9) & t_chofer & Chr(9) & t_origen & Chr$(9) & t_destino & Chr(9) & t_cartaporte & Chr(9) & cu & Chr(9) & t_kmts & Chr(9) & pu & Chr(9) & ti & Chr$(9) & im & Chr$(9) & puf & Chr$(9) & iv & Chr$(9) & pt, r
    vta_fact_viaje.msf1.RemoveItem r + 1
  End If
   
  s = 0
  V = 0
  For i = 1 To vta_fact_viaje.msf1.Rows - 1
      r = Val(vta_fact_viaje.msf1.TextMatrix(i, 11))
      s = s + r
      V = V + (r * vta_fact_viaje.msf1.TextMatrix(i, 10) / 100)
  Next i
  vta_fact_viaje.t_subtotal = s
  vta_fact_viaje.t_iva = V
  vta_fact_viaje.sacatotales

  para.producto_sel = 0
End Sub
 
  
Sub limpia()
t_cantidad = ""
t_fecha = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_ip = ""
t_chofer = ""
t_destino = ""
t_cartaporte = ""
t_kmts = ""



End Sub

Private Sub t_importe_GotFocus()
t_importe = Format$(Val(t_cantidad) * Val(t_pu), "#####0.00")
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If t_renglon = "" Then
   Call cargarenglon("A")
   t_fecha.SetFocus
   
  Else
   Call cargarenglon("M")
   Me.Hide
  End If
  Call limpia
  
Else
  Call solonum(KeyAscii, 1)
End If
End Sub

