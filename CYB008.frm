VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CYB_cc_detalle 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DETALLE DE CAJA"
   ClientHeight    =   7665
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame CUIT 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_op 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox t_fp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox t_idfp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Op."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Concepto Caja"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9960
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CYB008.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CYB008.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   3
      Top             =   7410
      Width           =   12075
      _ExtentX        =   21299
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
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4695
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
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
End
Attribute VB_Name = "CYB_cc_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String


Private Sub btnsale_Click()
Unload Me
End Sub









Private Sub Form_Activate()
Call carga
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 2700
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 2000
msf1.ColWidth(7) = 2000
msf1.TextMatrix(0, 0) = "Num."
msf1.TextMatrix(0, 1) = "Concepto"
msf1.TextMatrix(0, 2) = "Fecha"
msf1.TextMatrix(0, 3) = "Descripcion"
msf1.TextMatrix(0, 4) = "Entradas"
msf1.TextMatrix(0, 5) = "Salidas"
msf1.TextMatrix(0, 6) = "Operacion"
msf1.TextMatrix(0, 7) = "Contrapartida "

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub carga()
Call armagrid
If t_idfp <> "" Then
  F = ""
  Select Case t_op
  Case Is = "A"
    F = " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
  Case Is = "E"
    F = " and datevalue([fecha]) = datevalue('" & t_fecha & "') and [ubicacion] = 'D'"
  Case Is = "S"
    F = " and datevalue([fecha]) = datevalue('" & t_fecha & "') and [ubicacion] = 'H'"
  Case Is = "D"
    F = " and datevalue([fecha]) = datevalue('" & t_fecha & "')"
  Case Is = "T"
    F = " and datevalue([fecha]) <= datevalue('" & t_fecha & "')"
  
  End Select
  
  q = "select * from cyb_05, c_01 , cyb_01 where cyb_05.[id_forma_pago] = " & Val(t_idfp) & F & " and [id_cuenta_contra] = [id_cuenta] and cyb_05.[id_forma_pago] = cyb_01.[id_forma_pago]"
  
  Set rs = New ADODB.Recordset
 
  rs.Open q, cn1
  While Not rs.EOF
     If rs("ubicacion") = "D" Then
        e = Format$(rs("importe"), "######0.00")
        s = ""
     Else
        s = Format$(rs("importe"), "######0.00")
        e = ""
     End If
     msf1.AddItem Format$(rs("num_mov_caja"), "00000") & Chr$(9) & rs("cyb_01.descripcion") & Chr$(9) & Format$(rs("fecha"), "dd/mm/yyyy") & Chr$(9) & rs("cyb_05.descripcion") & Chr$(9) & e & Chr$(9) & s & Chr$(9) & rs("operacion") & Chr$(9) & rs("c_01.descripcion")
     rs.MoveNext
  Wend
  Set rs = Nothing
End If
End Sub

