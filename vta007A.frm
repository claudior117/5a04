VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_recibo1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "COMPROBANTES PENDIENTES DE PAGO"
   ClientHeight    =   6345
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   15765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   15765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   9255
      Begin VB.Label Label1 
         Caption         =   "Todos los saldos y cancelaciones son en $. En caso  de ser comprobantes en U$s se convertiran segun cotizacion original."
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Height          =   855
      Left            =   10080
      TabIndex        =   3
      Top             =   4920
      Width           =   5055
      Begin VB.TextBox T_APAGAR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "TOTAL A PAGAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comprobantes a Aplicar"
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15615
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4335
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   7646
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   5940
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_recibo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creaci�n impl�cita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 27
        
        Me.Hide
End Select
End Sub
Function APAGAR() As Double
  t = 0
  If msf1.Rows > 1 Then
   For i = 1 To msf1.Rows - 1
      t = t + Val(msf1.TextMatrix(i, 9))
   Next i
  End If
  APAGAR = t
End Function

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 1300
msf1.ColWidth(2) = 2500
msf1.ColWidth(3) = 1700
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1700
msf1.ColWidth(6) = 1600
msf1.ColWidth(7) = 500
msf1.ColWidth(8) = 1700
msf1.ColWidth(9) = 1700

msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Comprobante"
msf1.TextMatrix(0, 3) = "Total $"
msf1.TextMatrix(0, 4) = "Num.Int."
msf1.TextMatrix(0, 5) = "Neto"
msf1.TextMatrix(0, 6) = "Total U$s"
msf1.TextMatrix(0, 7) = "Tipo"
msf1.TextMatrix(0, 8) = "Saldo $"
msf1.TextMatrix(0, 9) = "A Aplicar $"
End Sub

Sub carga()
vta_recibo.armagrid
k = 0
r = 1
While k < msf1.Rows
  If msf1.TextMatrix(k, 0) = "**" Then
   F = msf1.TextMatrix(k, 1)
   c = msf1.TextMatrix(k, 2)
   vta_recibo.msf1.AddItem F & Chr(9) & c & Chr(9) & msf1.TextMatrix(k, 3) & Chr(9) & msf1.TextMatrix(k, 4) & Chr(9) & msf1.TextMatrix(k, 5) & Chr(9) & "" & msf1.TextMatrix(k, 6) & Chr(9) & msf1.TextMatrix(k, 7) & Chr(9) & Format$(Val(msf1.TextMatrix(k, 8)) - Val(msf1.TextMatrix(k, 9)), "#####0.00") & Chr(9) & msf1.TextMatrix(k, 9)
   r = r + 1
  End If
  k = k + 1
Wend

   
End Sub

Private Sub Form_Load()
'Call barraesag(Me)
Me.StatusBar1.Panels.item(1) = "Agregar comprobantes a recibo"
End Sub

  


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[Barra]Pago Total-[F2]Pago Parcial-[F3] Cambia Saldo - [F5]Todos Total- [F9] Agrega a Rbo "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
  Call carga
  Me.Hide
End If

If KeyCode = vbKeyF2 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 8)) > 0 Then
         k = InputBox$("Ingrese la cantidad a Cobrar", "COBRO PARCIAL")
         If Val(k) > 0 And Val(k) <= Val(msf1.TextMatrix(msf1.Row, 8)) Then
             msf1.TextMatrix(msf1.Row, 9) = Format$(Val(k), "######0.00")
             msf1.TextMatrix(msf1.Row, 0) = "**"
             'p = Val(msf1.TextMatrix(msf1.Row, 9)) / Val(msf1.TextMatrix(msf1.Row, 9))
             'msf1.TextMatrix(msf1.Row, 11) = Format$(Val(msf1.TextMatrix(msf1.Row, 6)) * p, "######0.00")
          Else
             msf1.TextMatrix(msf1.Row, 9) = ""
             msf1.TextMatrix(msf1.Row, 0) = ""
             
         End If
    End If
  End If
  T_APAGAR = Format$(APAGAR, "######0.00")
End If

If KeyCode = vbKeyF5 Then
  If msf1.Rows > 1 Then
   For i = 1 To msf1.Rows - 1
    If Val(msf1.TextMatrix(i, 8)) > 0 Then
             msf1.TextMatrix(i, 9) = Format$(Val(msf1.TextMatrix(i, 8)), "######0.00")
             msf1.TextMatrix(i, 0) = "**"
    End If
   Next i
   T_APAGAR = Format$(APAGAR, "######0.00")
  End If
End If

If KeyCode = vbKeyF3 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 3)) > 0 Then
         k = InputBox$("Ingrese el saldo pendiente", "CAMBIAR SALDO PENDIENTE")
         If Val(k) > 0 And Val(k) <= Val(msf1.TextMatrix(msf1.Row, 3)) Then
             msf1.TextMatrix(msf1.Row, 8) = Format$(Val(k), "######0.00")
             
         End If
    End If
  End If
  
End If



End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
  'If Val(msf1.TextMatrix(msf1.Row, 8)) > 0 Then
      If msf1.TextMatrix(msf1.Row, 0) = "**" Then
          msf1.TextMatrix(msf1.Row, 0) = ""
          msf1.TextMatrix(msf1.Row, 9) = ""
      Else
         msf1.TextMatrix(msf1.Row, 0) = "**"
         msf1.TextMatrix(msf1.Row, 9) = msf1.TextMatrix(msf1.Row, 8)
      End If
  'End If
  T_APAGAR = APAGAR
End If

End Sub

