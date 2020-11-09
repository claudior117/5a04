VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form op1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4830
   ScaleWidth      =   12180
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Height          =   615
      Left            =   8520
      TabIndex        =   5
      Top             =   3600
      Width           =   3495
      Begin VB.TextBox T_APAGAR 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "TOTAL A PAGAR"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   7695
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"Proc005a.frx":0000
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comprobantes a Aplicar"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5530
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4575
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
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
            TextSave        =   "09:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "op1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Function APAGAR() As Double
  t = 0
  If msf1.Rows > 1 Then
   For i = 1 To msf1.Rows - 1
      t = t + Val(msf1.TextMatrix(i, 10))
   Next i
  End If
  APAGAR = t
End Function

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 13
msf1.ColWidth(0) = 300
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 1900
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 800
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 800
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 1000
msf1.ColWidth(11) = 1000
msf1.ColWidth(12) = 1000


msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Comprobante"
msf1.TextMatrix(0, 3) = "Total $"
msf1.TextMatrix(0, 4) = "Ret.Gan"
msf1.TextMatrix(0, 5) = "Num.Int."
msf1.TextMatrix(0, 6) = "Neto $"
msf1.TextMatrix(0, 7) = "Total U$s"
msf1.TextMatrix(0, 8) = "Cuenta"
msf1.TextMatrix(0, 9) = "Saldo $"
msf1.TextMatrix(0, 10) = "Aplicar $"
msf1.TextMatrix(0, 11) = "Neto Aplicar $"
msf1.TextMatrix(0, 12) = "Ret. IB"

For i = 0 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 10
  msf1.ColAlignment(i) = 9
Next i


End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 27
        
        Me.Hide
End Select
End Sub
Sub carga()
op.armagrid
r = 1
minimo = 9999999.99
cribmax = 0
aliribmax = 0
If msf1.Rows > 1 Then
 For i = 1 To msf1.Rows - 1
  If msf1.TextMatrix(i, 0) = "**" Then
     F = msf1.TextMatrix(i, 1)
     c = msf1.TextMatrix(i, 2)
     rg = msf1.TextMatrix(i, 4)
     rib = msf1.TextMatrix(i, 12)
     op.msf1.AddItem F & Chr(9) & c & Chr(9) & msf1.TextMatrix(i, 9) & Chr(9) & rg & Chr(9) & msf1.TextMatrix(i, 5) & Chr$(9) & msf1.TextMatrix(i, 11) & Chr$(9) & msf1.TextMatrix(i, 7) & Chr$(9) & msf1.TextMatrix(i, 8) & Chr$(9) & msf1.TextMatrix(i, 10) & Chr$(9) & rib
  End If
 Next i
End If
End Sub
Private Sub Form_Load()
Call barraesag(Me)

End Sub

  




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[Barra] Pago Total - [F2] Pago Parcial - [F3] Cambia Saldo -  [F5] Todos Total  - [F9] Agrega a OP "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
  Call carga
  Me.Hide
End If


If KeyCode = vbKeyF2 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 5)) > 0 Then
         k = InputBox$("Ingrese la cantidad a Pagar", "PAGO PARCIAL")
         If Val(k) > 0 And Val(k) <= Val(msf1.TextMatrix(msf1.Row, 9)) Then
             msf1.TextMatrix(msf1.Row, 10) = Format$(Val(k), "######0.00")
             msf1.TextMatrix(msf1.Row, 0) = "**"
             p = Val(msf1.TextMatrix(msf1.Row, 6)) / Val(msf1.TextMatrix(msf1.Row, 3))
             msf1.TextMatrix(msf1.Row, 11) = Format$(Val(msf1.TextMatrix(msf1.Row, 10)) * p, "######0.00")
         Else
             MsgBox ("EL importe a aplicar del comprobante debe ser mayor que 0 y menor que el saldo impago ")
         
         End If
    End If
  End If
  T_APAGAR = Format$(APAGAR, "######0.00")
End If

If KeyCode = vbKeyF3 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 5)) > 0 Then
         k = InputBox$("Ingrese Saldo Impago del Comprobantes antes de emitir O.P.", "SALDO COMPROBANTE")
         If Val(k) > 0 And Val(k) <= Val(msf1.TextMatrix(msf1.Row, 3)) Then
             J = MsgBox("Confirma Cambio de saldo", 4)
             If J = 6 Then
                Call nivel_acceso(2)
                If para.id_grupo_modulo_actual >= 8 Then
                  msf1.TextMatrix(msf1.Row, 9) = Format$(Val(k), "######0.00")
                  Set rs = New ADODB.Recordset
                  q = "select [saldo_impago] from a5 where [num_int] = " & Val(msf1.TextMatrix(msf1.Row, 5))
                  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF And Not rs.BOF Then
                     rs("saldo_impago") = Format(Val(k), "######0.00")
                     rs.Update
                  Else
                     MsgBox ("Comprobante Inexistente")
                  End If
                  Set rs = Nothing
                Else
                  Call sinpermisos
                End If
             End If
           Else
             MsgBox ("EL saldo del comprobante debe ser mayor que 0 y menor que el importe total del comprobante")
         End If
    End If
  End If
  T_APAGAR = Format$(APAGAR, "######0.00")
End If


If KeyCode = vbKeyF5 Then
  If msf1.Rows > 1 Then
   For i = 1 To msf1.Rows - 1
    If Val(msf1.TextMatrix(i, 5)) > 0 Then
             msf1.TextMatrix(i, 10) = Format$(Val(msf1.TextMatrix(i, 9)), "######0.00")
             msf1.TextMatrix(i, 0) = "**"
             p = Val(msf1.TextMatrix(i, 6)) / Val(msf1.TextMatrix(i, 3))
             msf1.TextMatrix(i, 11) = Format$(Val(msf1.TextMatrix(i, 10)) * p, "######0.00")

    End If
   Next i
   T_APAGAR = Format$(APAGAR, "######0.00")
  End If
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
  If Val(msf1.TextMatrix(msf1.Row, 5)) > 0 Then
      If msf1.TextMatrix(msf1.Row, 0) = "**" Then
          msf1.TextMatrix(msf1.Row, 0) = ""
          msf1.TextMatrix(msf1.Row, 10) = ""
          msf1.TextMatrix(msf1.Row, 11) = ""
          
      Else
         msf1.TextMatrix(msf1.Row, 0) = "**"
         msf1.TextMatrix(msf1.Row, 10) = Format$(Val(msf1.TextMatrix(msf1.Row, 9)), "######0.00")
         p = Val(msf1.TextMatrix(msf1.Row, 6)) / Val(msf1.TextMatrix(msf1.Row, 3))
         msf1.TextMatrix(msf1.Row, 11) = Format$(Val(msf1.TextMatrix(msf1.Row, 10)) * p, "######0.00")
      End If
  End If
  T_APAGAR = Format$(APAGAR, "######0.00")
End If

End Sub
