VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form vta_facturacion_perc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PERCEPCIONES DE VENTA"
   ClientHeight    =   6180
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   16455
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   16455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_totalperc 
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
      Left            =   10320
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   18
      Top             =   5280
      Width           =   2175
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   5775
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   15875
            MinWidth        =   15875
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox t_modulo 
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Caption         =   "Ingreso Percepcion"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   16095
      Begin VB.TextBox t_cod 
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
         Left            =   13800
         Locked          =   -1  'True
         MaxLength       =   21
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox t_cuenta 
         Enabled         =   0   'False
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
         Left            =   12000
         MaxLength       =   21
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox t_base 
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
         Left            =   5760
         MaxLength       =   21
         TabIndex        =   13
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox t_alicuota 
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
         Left            =   8160
         MaxLength       =   21
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox t_importe 
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
         Left            =   9720
         MaxLength       =   21
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox c_perc 
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
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox t_renglon 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Afip"
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
         Left            =   13800
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Base Imponible"
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
         Left            =   5640
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Alicuota"
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
         Left            =   7800
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Cuenta"
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
         Left            =   12000
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Importe"
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
         Left            =   9600
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Percepcion/Retencion"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Renglon"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   6376
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Total Percepciones"
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
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   5280
      Width           =   2415
   End
End
Attribute VB_Name = "vta_facturacion_perc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creaci�n impl�cita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Sub limpia()
   Call armagrid
  
  
End Sub



Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 700
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 4500
msf1.ColWidth(3) = 2500
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 2500
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1500
msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Perc."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Base Imp."
msf1.TextMatrix(0, 4) = "Alicuota"
msf1.TextMatrix(0, 5) = "Importe"
msf1.TextMatrix(0, 6) = "Cod."
msf1.TextMatrix(0, 7) = "Cuenta"

End Sub


 


Private Sub c_perc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  
  t_base.SetFocus
End If

If KeyAscii = 27 Then
  Frame1.Visible = False
End If

End Sub

Private Sub c_perc_LostFocus()
Call buscaperc
End Sub

Private Sub Form_Activate()
Frame1.Visible = False
t_base = vta_facturacion.t_subtotal
Call carga_percepciones_venta(c_perc)
msf1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
  
End Select
End Sub

Sub buscaperc()
  q = "select * from I_01 where [id_impuesto] = " & c_perc.ItemData(c_perc.ListIndex)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    t_alicuota = rs("tasa_i1")
    t_cuenta = rs("id_cuenta_i1")
    t_cod = rs("id_otrostributos") 'codigo afip
  Else
     t_alicuota = 0
     t_cuenta = "110302"
     t_cod = 99
  End If
  Set rs = Nothing
End Sub
Sub cargarenglon()
  
  ip = c_perc.ItemData(c_perc.ListIndex)
  d = c_perc
  im = Format$(t_importe, "#####0.00")
  b = Format$(t_base, "#####0.00")
  a = Format$(t_alicuota, "#####0.00")
  cta = t_cuenta
  
  If EXISTE = "N" Then
    r = msf1.Rows
    msf1.AddItem r & Chr(9) & Format$(ip, "000") & Chr(9) & d & Chr(9) & b & Chr(9) & a & Chr(9) & im & Chr(9) & t_cod & Chr(9) & cta
    Else
    r = Val(t_renglon)
    msf1.AddItem r & Chr(9) & Format$(ip, "000") & Chr(9) & d & Chr(9) & b & Chr(9) & a & Chr(9) & im & Chr(9) & t_cod & Chr(9) & cta, r
    msf1.RemoveItem r + 1
  End If
  
  Call calculatotal
  Frame1.Visible = False
  msf1.SetFocus
    
    
End Sub
Sub calculatotal()
'busco perc
If msf1.Rows > 1 Then
  t = 0
  For i = 1 To msf1.Rows - 1
    t = t + Val(msf1.TextMatrix(i, 5))
  Next i
  t_totalperc = Format$(t, "######0.00")
Else
  t_totalperc = Format$(0, "######0.00")
End If

End Sub
Private Sub Form_Load()
Call armagrid
Call barraesag(Me)

End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Sale"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False
Frame1.Visible = False
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF9 Then
  'If msf1.Rows > 1 Then
     Call sale
  'End If
   
End If

If KeyCode = vbKeyInsert Then
   EXISTE = "N"
   Frame1.Visible = True
   t_renglon = ""
   c_perc.ListIndex = 0
   t_importe = ""
   c_perc.SetFocus
End If
End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
   EXISTE = "S"
    t_renglon = msf1.Row
    c_perc.ListIndex = buscaindice(c_perc, msf1.TextMatrix(msf1.Row, 1))
    t_importe = msf1.TextMatrix(msf1.Row, 5)
    t_base = msf1.TextMatrix(msf1.Row, 3)
    t_alicuota = msf1.TextMatrix(msf1.Row, 4)
    t_cod = msf1.TextMatrix(msf1.Row, 6)
    t_cuenta = msf1.TextMatrix(msf1.Row, 7)
    
    Frame1.Visible = True
  End If
End If

If KeyAscii = 27 Then
   Call sale
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub

Sub sale()
  vta_facturacion.sacatotales
  vta_facturacion.t_perc.SetFocus
  Me.Hide

End Sub



Private Sub t_alicuota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  
  t_importe.SetFocus
End If

If KeyAscii = 27 Then
  Frame1.Visible = False
End If
End Sub

Private Sub t_alicuota_LostFocus()
t_importe = Format$(Val(t_base) * Val(t_alicuota) / 100, "#######0.00")
End Sub

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_alicuota.SetFocus
End If
If KeyAscii = 27 Then
  Frame1.Visible = False
End If
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Val(t_importe) > 0 Then
    Call cargarenglon
  End If
End If
If KeyAscii = 27 Then
  Frame1.Visible = False
End If


End Sub

Private Sub t_totalperc_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Call sale
End If

End Sub