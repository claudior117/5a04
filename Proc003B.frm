VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ABM_COMP_COMPRA2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INGRESO DE PERCEPCIONES/RETENCIONES"
   ClientHeight    =   6180
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   13425
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   5775
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   15875
            MinWidth        =   15875
            TextSave        =   ""
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
      TabIndex        =   9
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
      Width           =   12975
      Begin VB.ComboBox c_concepto 
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
         Left            =   8040
         TabIndex        =   5
         Top             =   840
         Width           =   4575
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
         Left            =   5640
         MaxLength       =   21
         TabIndex        =   4
         Top             =   840
         Width           =   2295
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Concepto percepción"
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
         Left            =   7680
         TabIndex        =   10
         Top             =   240
         Width           =   4935
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
         Left            =   5640
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   13095
      _ExtentX        =   23098
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
End
Attribute VB_Name = "ABM_COMP_COMPRA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   T_TOTAL = ""
  
End Sub



Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 700
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 2100
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 3000
msf1.ColWidth(6) = 600
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 2100
msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Perc."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Importe Perc."
msf1.TextMatrix(0, 4) = "Cuenta"
msf1.TextMatrix(0, 5) = "Concepto perc."
msf1.TextMatrix(0, 6) = "Cod."
msf1.TextMatrix(0, 7) = "Tasa"
msf1.TextMatrix(0, 8) = "Base Imp"

End Sub


 
Private Sub c_concepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And c_concepto.ListIndex >= 0 Then
  If Val(t_importe) > 0 Then
    Call cargarenglon
  End If
End If

If KeyAscii = 27 Then
  Frame1.Visible = False
End If

End Sub

Private Sub c_perc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_importe.SetFocus
End If

If KeyAscii = 27 Then
  Frame1.Visible = False
End If

End Sub

Private Sub c_perc_LostFocus()
If c_perc.ListIndex < 0 Then
  c_perc.ListIndex = 0
End If
q = "select * from a12 where [id_percepcion] = " & c_perc.ItemData(c_perc.ListIndex)
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  Select Case rs("impuesto12")
  Case Is = "G"
    Call carga_regimen_retvta(c_concepto, 217)
  Case Is = "I"
    Call carga_regimen_retvta(c_concepto, 767)
  Case Is = "S"
    Call carga_regimen_retvta(c_concepto, 736)
  Case Is = "B"
    Call carga_regimen_retvta(c_concepto, 1)
  Case Else
   c_concepto.clear
   c_concepto.AddItem "<Error en tablas de conceptos>", 0
   c_concepto.ListIndex = 0
  
  End Select
  
  
  
Else
   c_concepto.clear
   c_concepto.AddItem "<Error en tablas de conceptos>", 0
   c_concepto.ListIndex = 0
End If
Set rs = Nothing
End Sub

Private Sub Form_Activate()
Frame1.Visible = False
't_modulo C --> Perc Compra  V --> Ret Venta    P --> Percepcines

Select Case t_modulo

Case Is = "C"
  Call carga_percepciones(c_perc, "P") 'comprs
Case Is = "V"
  Call carga_percepciones(c_perc, "R") 'ventas directs
Case Is = "P"
  Call carga_percepciones(c_perc, "P") 'facturas de venta
Case Is = "S"
  Call carga_percepciones(c_perc, "P") 'Comprobantes varios ventas

Case Else
   Call carga_percepciones(c_perc, "T")
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
  
End Select
End Sub
Sub cargarenglon()
  
  ip = c_perc.ItemData(c_perc.ListIndex)
  d = c_perc
  im = Format$(t_importe, "#####0.00")
  q = "select * from a12 where [id_percepcion] = " & c_perc.ItemData(c_perc.ListIndex)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     cta = rs("id_cuenta")
  Else
     cta = 0
  End If
  Set rs = Nothing
  
  If EXISTE = "N" Then
    r = ABM_COMP_COMPRA2.msf1.Rows
    ABM_COMP_COMPRA2.msf1.AddItem r & Chr(9) & Format$(ip, "000") & Chr(9) & d & Chr(9) & im & Chr(9) & cta & Chr(9) & c_concepto & Chr(9) & c_concepto.ItemData(c_concepto.ListIndex)
   Else
    r = Val(t_renglon)
    ABM_COMP_COMPRA2.msf1.AddItem r & Chr(9) & Format$(ip, "000") & Chr(9) & d & Chr(9) & im & Chr(9) & cta & Chr(9) & c_concepto & Chr(9) & c_concepto.ItemData(c_concepto.ListIndex), r
    ABM_COMP_COMPRA2.msf1.RemoveItem r + 1
  End If
  Frame1.Visible = False
  Select Case t_modulo
  Case Is = "C"
    ABM_COMP_COMPRA.sacatotales
  Case Is = "V"
    If msf1.Rows > 1 Then
      t = 0
      For i = 1 To msf1.Rows - 1
        t = t + Val(msf1.TextMatrix(i, 3))
      Next i
      vta_directa.t_perc = Format$(t, "######0.00")
    Else
      vta_directa.t_perc = Format$(0, "######0.00")
    End If
  Case Is = "L"
   
    
   Case Is = "P"
     vta_facturacion.sacatotales
   
   Case Is = "S"
   If msf1.Rows > 1 Then
      t = 0
      For i = 1 To msf1.Rows - 1
        t = t + Val(msf1.TextMatrix(i, 3))
      Next i
      vta_COMPVARIOS.t_perc = Format$(t, "######0.00")
    Else
      vta_COMPVARIOS.t_perc = Format$(0, "######0.00")
    End If
    
   End Select
  
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
    t_importe = msf1.TextMatrix(msf1.Row, 3)
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
 Select Case t_modulo
 Case Is = "C"
  ABM_COMP_COMPRA.t_perc.SetFocus
 Case Is = "V"
  vta_directa.t_perc.SetFocus
 Case Is = "L"
  vta_liqcereal.t_perc.SetFocus
Case Is = "P"
  vta_facturacion.t_perc.SetFocus
 
Case Is = "S"
   vta_COMPVARIOS.t_perc.Enabled = True
   vta_COMPVARIOS.t_perc.SetFocus
End Select
  Me.Hide

End Sub

Sub sacatotales()

End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Val(t_importe) > 0 Then
     c_concepto.SetFocus
  End If
End If

End Sub

