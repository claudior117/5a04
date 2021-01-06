VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "IF Universal - Demo"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Datos Iniciales"
      Height          =   1095
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Text            =   "2"
      Top             =   720
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":005F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cierre X"
      Height          =   1095
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cierre Z"
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ticket"
      Height          =   1095
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Factura B"
      Height          =   1095
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Factura A"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Puerto COM:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo de impresora:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cMODELO As Integer
Dim cPUERTO As Integer
Const cBAUDIOS = bd9600

Private Sub Combo1_Click(Index As Integer)
  cMODELO = CInt(Combo1.Item(0).ItemData(Combo1.Item(0).ListIndex))
End Sub

Private Sub Command2_Click()
  'Si el monto supera 1000 pesos deben enviarse los datos del cliente.
  Call ImprimeFactura(tcFactura_B, False)
End Sub

Private Sub ImprimeFactura(TipoComprobante As TipoDeComprobante, ByVal ImprimeDatosCliente As Boolean)

  Dim Fiscal As Driver
  
  Set Fiscal = New Driver
  
  On Error GoTo DepuraErrores
  
  Fiscal.Modelo = cMODELO
  Fiscal.Puerto = cPUERTO
  Fiscal.Baudios = cBAUDIOS
  
  If Not Fiscal.Inicializar Then
    Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
  End If
  
  Fiscal.CancelarComprobante
    
  If ImprimeDatosCliente Then
    If Not Fiscal.DatosCliente("Abel Miranda", tdCUIT, 20939802593#, riResponsableInscripto, "Haefreingue, 1686, Moron") Then
       Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
    End If
  End If
  
'****** USAR ESTE METODO PARA INFORMAR COMPROBANTES RELACIONADOS EN CASOS DE FACTURAS,NC, ND ********
'  If Not Fiscal.DocumentoDeReferencia2g(tcRemito, "0001-00000001") Then
'     Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
'  End If
  
  If Not Fiscal.AbrirComprobante(TipoComprobante) Then
     Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
  End If
  
  If Not Fiscal.ImprimirItem2g("Item 1", 1, 0.1, 21, 0, IFUniversal.Gravado, "0", 1, "7790001001054", "", IFUniversal.Unidad) Then
     Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
  End If
  
  If Not Fiscal.ImprimirDescuentoGeneral("Descuento General", 0.01) Then
     Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
  End If
  
  If Not Fiscal.ImprimirPago2g("Efectivo", 5, "", IFUniversal.Efectivo, 1, "", "") Then
     Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
  End If
  
  Fiscal.CerrarComprobante
  
  Fiscal.Finalizar
  
  MsgBox ("Comprobante impreso exitosamente")
  Exit Sub

DepuraErrores:
  Fiscal.Finalizar
  MsgBox Fiscal.ErrorDesc
End Sub

Private Sub Command3_Click()
  Call ImprimeFactura(tcTique, False)
End Sub
Private Sub Command1_Click()
  Call ImprimeFactura(tcFactura_A, True)
End Sub

Private Sub Command4_Click()
  Dim Fiscal As Driver
  
  Set Fiscal = New Driver
  
  Fiscal.Modelo = cMODELO
  Fiscal.Puerto = cPUERTO
  Fiscal.Baudios = cBAUDIOS
  
  If Fiscal.Inicializar Then
    Fiscal.CancelarComprobante
      
    If Fiscal.CierreZ Then
      MsgBox ("Cierre realizado exitosamente")
    Else
      MsgBox (Fiscal.ErrorDesc)
    End If
    
    Fiscal.Finalizar
  Else
    MsgBox (Fiscal.ErrorDesc)
  End If
End Sub

Private Sub Command5_Click()
  Dim Fiscal As Driver
  
  Set Fiscal = New Driver
  
  Fiscal.Modelo = cMODELO
  Fiscal.Puerto = cPUERTO
  Fiscal.Baudios = cBAUDIOS
  
  If Fiscal.Inicializar Then
  
    Fiscal.CancelarComprobante
      
    If Fiscal.CierreX Then
      MsgBox ("Cierre realizado exitosamente")
    Else
      MsgBox (Fiscal.ErrorDesc)
    End If
    
    Fiscal.Finalizar
  Else
    MsgBox (Fiscal.ErrorDesc)
  End If
End Sub

Private Sub Command6_Click()
  Dim Fiscal As Driver
  
  Set Fiscal = New Driver
  
  Fiscal.Modelo = cMODELO
  Fiscal.Puerto = cPUERTO
  Fiscal.Baudios = cBAUDIOS
  
  If Fiscal.Inicializar Then
  
    Dim DatosIni As ObtenerDatosDeInicializacionRespuesta
    Set DatosIni = Fiscal.ObtenerDatosDeInicializacion
    If DatosIni.Resultado Then
      PuntoVta = DatosIni.NroPOS
      MsgBox "El punto de venta es: " + CStr(PuntoVta)
    Else
      MsgBox Fiscal.ErrorDesc
    End If
    Fiscal.Finalizar
  Else
    MsgBox (Fiscal.ErrorDesc)
  End If
End Sub

Private Sub Form_Load()
  cPUERTO = 2
End Sub

Private Sub Text1_Change()
  cPUERTO = CInt(Text1.Text)
End Sub
