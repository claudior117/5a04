VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ABM_COMP_COMPRA5 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DETALLE DE IVA"
   ClientHeight    =   5715
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   7695
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Totales"
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
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   7215
      Begin VB.TextBox t_totiva 
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
         Left            =   5280
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox t_totneto 
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
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingreso iVA}"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   7335
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
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox t_iva 
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
         Left            =   5280
         MaxLength       =   14
         TabIndex        =   2
         Top             =   720
         Width           =   1935
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
         Height          =   360
         Left            =   2400
         MaxLength       =   14
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Iva"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Neto Gravado"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Tasa Iva"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4048
      _Version        =   393216
      BackColorBkg    =   12648447
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   5310
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "01/03/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:35 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ABM_COMP_COMPRA5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creaci�n impl�cita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Sub limpia()
  c_tasa.SetFocus
  t_importe = ""
  t_iva = ""
End Sub



Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 4
msf1.ColWidth(0) = 700
msf1.ColWidth(1) = 1700
msf1.ColWidth(2) = 2500
msf1.ColWidth(3) = 2200

msf1.TextMatrix(0, 0) = "id"
msf1.TextMatrix(0, 1) = "Tasa"
msf1.TextMatrix(0, 2) = "Neto Gravado"
msf1.TextMatrix(0, 3) = "Iva"

t_totneto = ""
t_totiva = ""
End Sub


 

Sub cargarenglon()
If Val(t_iva) > 0 Then
    'buscatasa
    i = 1
    While i < msf1.Rows
      If Val(msf1.TextMatrix(i, 0)) = c_tasa.ListIndex Then
        If msf1.Rows <= 2 Then
          Call armagrid
        Else
         msf1.RemoveItem i
         i = msf1.Rows
        End If
      End If
      i = i + 1
    
    Wend
        
    im = Format$(t_importe, "#####0.00")
    ti = Format$(t_iva, "#####0.00")
    msf1.AddItem c_tasa.ListIndex & Chr(9) & Format$(c_tasa, "##.00") & Chr(9) & im & Chr(9) & ti
    Call sacatotales
 End If
  
End Sub

Sub actualiza()
    ABM_COMP_COMPRA.t_iva = Format$(t_totiva, "######0.00")
    ABM_COMP_COMPRA.t_subtotal = Format$(t_totneto, "######0.00")
     ABM_COMP_COMPRA.t_iva.SetFocus
End Sub
Private Sub Form_Activate()
Call limpia
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 2)
End If


If KeyAscii = 27 Then
  Call actualiza
  Me.Hide
End If


End Sub

Private Sub Form_Load()
Call armagrid
Call barraesag(Me)

Call carga_alicuotaiva(c_tasa)
c_tasa.ListIndex = 0

End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Sale"
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
 Call sacatotales
End If



End Sub


Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub

Sub sacatotales()
If msf1.Rows > 1 Then
      tn = 0
      ti = 0
      For i = 1 To msf1.Rows - 1
        tn = tn + Val(msf1.TextMatrix(i, 2))
        ti = ti + Val(msf1.TextMatrix(i, 3))
      Next i
      t_totneto = Format$(tn, "######0.00")
      t_totiva = Format$(ti, "######0.00")
Else
      t_totneto = 0
      t_totiva = 0
End If


End Sub

Private Sub t_importe_LostFocus()
t_iva = Format$(Val(c_tasa) * Val(t_importe) / 100, "####0.00")
End Sub

Private Sub t_iva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call cargarenglon
 Call limpia
End If
End Sub


