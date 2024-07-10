VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_facte1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "IMPORTAR FACTURAS ELECTRONICAS DESDE DUPLICADO DIGITAL (R.G. 1361)"
   ClientHeight    =   9480
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   12105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9480
   ScaleWidth      =   12105
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de Importacion"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   11655
      Begin VB.ComboBox c_actividad 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Proceso"
      Height          =   1695
      Left            =   240
      TabIndex        =   11
      Top             =   7440
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "2.   Importar comprobantes seleccionados                         "
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1.   Cargar y Verificar Duplicados de Facturas Electronicas"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLIENTES"
      ForeColor       =   &H00C00000&
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5530
         _Version        =   393216
         BackColorBkg    =   12632256
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SELECCIONE CARPETA DONDE SE ALMACENAN LOS DUPLICADOS"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9735
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ubicacion definitiva de los archivos"
         Height          =   735
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   5055
         Begin VB.TextBox t_camino 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   5055
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   5880
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9840
      TabIndex        =   1
      Top             =   8160
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta031.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta031.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   9225
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_facte1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim t1 As String

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 4000
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 2000
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 3000
msf1.ColWidth(7) = 2500
msf1.ColWidth(8) = 1500


msf1.TextMatrix(0, 0) = " "
msf1.TextMatrix(0, 1) = "Estado "
msf1.TextMatrix(0, 2) = "Archivo Cabecera"
msf1.TextMatrix(0, 3) = "Fecha"
msf1.TextMatrix(0, 4) = "Comprobante"
msf1.TextMatrix(0, 5) = "Cuit"
msf1.TextMatrix(0, 6) = "Cliente"
msf1.TextMatrix(0, 7) = "Observaciones"
msf1.TextMatrix(0, 8) = "Otras_Perc"



End Sub

Function verifica() As Boolean
If t_fecha <> "" And t_fecha2 <> "" Then
  verifica = True
Else
  verifica = False
End If

  
End Function

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_actividad_KeyUp(KeyCode As Integer, Shift As Integer)
If c_actividad.ListIndex < 0 Then
 c_actividad.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
Call armagrid
Dim FileSystem As New FileSystemObject
Dim Folder As Folder
Dim CurrentFile As File
Dim FileColl As Files

Set Folder = FileSystem.GetFolder(t_camino)
'Set Folder = FileSystem.GetFolder("C:\facte")

Set FileColl = Folder.Files
If FileColl.Count > 0 Then
  With msf1
   For Each CurrentFile In FileColl
      If Mid$(CurrentFile.Name, 30, 8) = "CABECERA" Then
       'VERIFICA
       '1) EXISTEN ARCHIVOS DETALLE Y VENTAS
         estado = "OK"
         Detalle = ""
         a0 = t_camino & Mid$(CurrentFile.Name, 1, 29) & "CABECERA.TXT"
         a1 = t_camino & Mid$(CurrentFile.Name, 1, 29) & "DETALLE.TXT"
         a2 = t_camino & Mid$(CurrentFile.Name, 1, 29) & "OTRAS_PERCEP.TXT"
         
         If Not FileSystem.FileExists(a1) Then
           estado = "ERR"
           Detalle = Detalle & " No existe Archivo DETALLE. "
         End If
         
         If Not FileSystem.FileExists(a2) Then
            percep = "N"
         Else
            percep = "S"
         End If
         
         'COMPRUEBO SI EL COMPROBANTE NO EXISTE EN EL SISTEMA
        Open a0 For Input As #1
        Line Input #1, l
        F = Mid$(l, 8, 2) & "/" & Mid$(l, 6, 2) & "/" & Mid$(l, 2, 4)
        tc = Val(Mid$(l, 10, 3))
        Select Case tc
          Case Is = 1
            letra = "A"
            cc = 1
            dc = "Fact."
          Case Is = 2
            letra = "A"
            cc = 2
            dc = "Nd.  "
          Case Is = 3
            letra = "A"
            cc = 3
            dc = "Nc.  "
          Case Is = 6
            letra = "B"
            cc = 1
            dc = "Fact."
          Case Is = 7
            letra = "B"
            cc = 2
            dc = "Nd.  "
          Case Is = 8
            letra = "B"
            cc = 3
            dc = "Nc.  "
           Case Is = 11
            letra = "C"
            cc = 1
            dc = "Fact."
          Case Is = 12
            letra = "C"
            cc = 2
            dc = "Nd.  "
          Case Is = 13
            letra = "C"
            cc = 3
            dc = "Nc.  "
          Case Is = 19
            letra = "E"
            cc = 1
            dc = "Fact."
          Case Is = 20
            letra = "E"
            cc = 2
            dc = "Nd.  "
          Case Is = 21
            letra = "E"
            cc = 3
            dc = "Nc.  "
           Case Is = 201
            letra = "A"
            cc = 30
            dc = "FCE  "
          Case Is = 202
            letra = "A"
            cc = 31
            dc = "NdCE "
          Case Is = 203
            letra = "A"
            cc = 32
            dc = "NcCE "
          Case Is = 206
            letra = "B"
            cc = 30
            dc = "FCE  "
          Case Is = 207
            letra = "B"
            cc = 31
            dc = "NdCE "
          Case Is = 208
            letra = "B"
            cc = 32
            dc = "NcCE "
          
          Case Else
            estado = "ERR"
            Detalle = Detalle & " Comprobante no soportado por el sistema. "
            letra = "X"
            cc = 1
            dc = "Error"
          End Select
          
          
          If tc < 100 Then
              comp = dc & " " & letra & Mid$(l, 13, 4) & "-" & Mid$(l, 17, 8)
              suc = Mid$(l, 13, 4)
              NUM = Mid$(l, 17, 8)
              CUIT = Mid$(l, 38, 11)
          Else
              comp = dc & " " & letra & Mid$(l, 14, 4) & "-" & Mid$(l, 18, 8)
              suc = Mid$(l, 14, 4)
              NUM = Mid$(l, 18, 8)
              CUIT = Mid$(l, 39, 11)
          
          End If
        
        
        Close #1
       
        'busco si el comprobante no fue cargado
        q = "select * from vta_02 where [id_tipocomp] = " & cc & " and [letra] = '" & letra & "' and [sucursal] = " & Val(suc) & " and [num_comp] = " & NUM
        Set rs = New adodb.Recordset
        rs.Open q, cn1
        If Not rs.EOF And Not rs.BOF Then
            estado = "ERR"
            Detalle = Detalle & " Comprobante ya ingresado. "
        End If
        Set rs = Nothing
        
        
        
        
        Set rs = New adodb.Recordset
        q = "select * from vta_01 where [cuit] = '" & CUIT & "'" ' "' or [cuit] = '" & Format$(CUIT, "@@-@@@@@@@@-@") & "'"
        rs.Open q, cn1
        If Not rs.EOF And Not rs.BOF Then
            Codc = rs("id_cliente")
            cli = rs("denominacion")
        Else
            estado = "ERR"
            Detalle = Detalle & " Cliente Inexistente. "
            Codc = 1
            cli = "Error"
        End If
        
        'verifico si pertenece a un periodo valido
        If verificaperiodog(F) = "C" Then
            estado = "ERR"
            Detalle = Detalle & " Periodo de Importacion Cerrado. "
          
        End If
        
        Set rs = Nothing
        
        cliente = Format$(Codc, "00000") & " " & cli
       msf1.AddItem " " & Chr$(9) & estado & Chr$(9) & CurrentFile.Name & Chr$(9) & F & Chr$(9) & comp & Chr$(9) & CUIT & Chr$(9) & cliente & Chr$(9) & Detalle & Chr$(9) & percep       'add item
      End If
  Next
 
End With
End If

Set FileSystem = Nothing
Set Folder = Nothing
Set FileColl = Nothing
Set CurrentFile = Nothing

End Sub

Private Sub Command2_Click()
c = cuenta
If c > 0 Then

  J = MsgBox("Confirma Importar " & c & " Comprobantes", 4)
  If J = 6 Then
     espere.Show
     espere.Refresh
     r = 1
     While r <= msf1.Rows - 1
      If msf1.TextMatrix(r, 0) = "**" Then
         Call graba(r)
      End If
      r = r + 1
    Wend
    Unload espere
    Unload Me
  End If
Else
 MsgBox ("No hay comprobantes seleccionados")
End If


End Sub

Function cuenta() As Integer
  r = 0
  c = 0
  While r <= msf1.Rows - 1
    If msf1.TextMatrix(r, 0) = "**" Then
      c = c + 1
    End If
    r = r + 1
  Wend
  cuenta = c
End Function
Private Sub Dir1_Change()

Call camino
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
Call camino
End Sub

Sub graba(ByVal r As Integer)

  a0 = t_camino & msf1.TextMatrix(r, 2)
  a1 = t_camino & Mid$(msf1.TextMatrix(r, 2), 1, 29) & "DETALLE.TXT"
  a2 = t_camino & Mid$(msf1.TextMatrix(r, 2), 1, 29) & "OTRAS_PERCEP.TXT"
  tp = msf1.TextMatrix(r, 8)
  
  Open a0 For Input As #1
  Line Input #1, l
        F = Mid$(l, 8, 2) & "/" & Mid$(l, 6, 2) & "/" & Mid$(l, 2, 4)
        tc = Val(Mid$(l, 10, 3))
        Select Case tc
          Case Is = 1
            letra = "A"
            cc = 1
            dc = "Fact."
          Case Is = 2
            letra = "A"
            cc = 2
            dc = "Nd.  "
          Case Is = 3
            letra = "A"
            cc = 3
            dc = "Nc.  "
          Case Is = 6
            letra = "B"
            cc = 1
            dc = "Fact."
          Case Is = 7
            letra = "B"
            cc = 2
            dc = "Nd.  "
          Case Is = 8
            letra = "B"
            cc = 3
            dc = "Nc.  "
           Case Is = 11
            letra = "C"
            cc = 1
            dc = "Fact."
          Case Is = 12
            letra = "C"
            cc = 2
            dc = "Nd.  "
          Case Is = 13
            letra = "C"
            cc = 3
            dc = "Nc.  "
          Case Is = 19
            letra = "E"
            cc = 1
            dc = "Fact."
          Case Is = 20
            letra = "E"
            cc = 2
            dc = "Nd.  "
          Case Is = 21
            letra = "E"
            cc = 3
            dc = "Nc.  "
          Case Is = 201
            letra = "A"
            cc = 30
            dc = "FCE  "
          Case Is = 202
            letra = "A"
            cc = 31
            dc = "NdCE "
          Case Is = 203
            letra = "A"
            cc = 32
            dc = "NcCE "
          Case Is = 206
            letra = "B"
            cc = 30
            dc = "FCE  "
          Case Is = 207
            letra = "B"
            cc = 31
            dc = "NdCE "
          Case Is = 208
            letra = "B"
            cc = 32
            dc = "NcCE "
  
            
            
          Case Else
            estado = "ERR"
            Detalle = Detalle & " Comprobante no soportado por el sistema. "
            letra = "X"
            cc = 1
            dc = "Error"
          End Select
          
          
          If tc < 100 Then
           comp = dc & " " & letra & Mid$(l, 13, 4) & "-" & Mid$(l, 17, 8)
           suc = Mid$(l, 13, 4)
           NUM = Mid$(l, 17, 8)
           CUIT = Mid$(l, 38, 11)
            total = Val(Mid$(l, 79, 13) & "." & Mid$(l, 92, 2))
            subtotal = Val(Mid$(l, 109, 13) & "." & Mid$(l, 122, 2))
            'iva = Val(Mid$(l, 124, 13) & "." & Mid$(l, 137, 2)) + Val(Mid$(l, 139, 13) & "." & Mid$(l, 152, 2))
            'nograbado = Val(Mid$(l, 94, 13) & "." & Mid$(l, 107, 2)) + Val(Mid$(l, 139, 13) & "." & Mid$(l, 152, 2))
            iva = Val(Mid$(l, 124, 13) & "." & Mid$(l, 137, 2))
            nograbado = Val(Mid$(l, 94, 13) & "." & Mid$(l, 107, 2))
            cotizacion = Val(Mid$(l, 249, 4) & "." & Mid$(l, 253, 6))
            percib = Val(Mid$(l, 184, 13) & "." & Mid$(l, 197, 2)) 'percepciones ib ba
            moneda = Mid$(l, 246, 3)
            otrasperc = Val(Mid$(l, 139, 13) & "." & Mid$(l, 152, 2))
            
          
          Else
           comp = dc & " " & letra & Mid$(l, 14, 4) & "-" & Mid$(l, 18, 8)
           suc = Mid$(l, 14, 4)
           NUM = Mid$(l, 18, 8)
           CUIT = Mid$(l, 39, 11)
            total = Val(Mid$(l, 80, 13) & "." & Mid$(l, 93, 2))
            subtotal = Val(Mid$(l, 110, 13) & "." & Mid$(l, 123, 2))
            'iva = Val(Mid$(l, 125, 13) & "." & Mid$(l, 138, 2)) + Val(Mid$(l, 140, 13) & "." & Mid$(l, 153, 2))
            'nograbado = Val(Mid$(l, 95, 13) & "." & Mid$(l, 108, 2)) + Val(Mid$(l, 140, 13) & "." & Mid$(l, 153, 2))
            iva = Val(Mid$(l, 125, 13) & "." & Mid$(l, 138, 2))
            nograbado = Val(Mid$(l, 95, 13) & "." & Mid$(l, 108, 2))
            cotizacion = Val(Mid$(l, 250, 4) & "." & Mid$(l, 254, 6))
            percib = Val(Mid$(l, 185, 13) & "." & Mid$(l, 198, 2))
            moneda = Mid$(l, 247, 3)
            otrasperc = Val(Mid$(l, 140, 13) & "." & Mid$(l, 153, 2))
           
          End If
        
        Close #1
       
        
        Set rs = New adodb.Recordset
        q = "select * from vta_01 where [cuit] = '" & CUIT & "'" ' "' or [cuit] = '" & Format$(CUIT, "@@-@@@@@@@@-@") & "'"
        rs.Open q, cn1
        If Not rs.EOF And Not rs.BOF Then
            Codc = rs("id_cliente")
            cli = rs("denominacion")
            Dire = rs("DIRECCION")
            Loca = rs("LOCALIDAD")
        End If
        Set rs = Nothing
  
  
    
 
 
 If percib > 0 Then
   tpercib = (percib * 100) / subtotal
 Else
   tpercib = 0
 End If
 
 'MsgBox (total)
 If cotizacion < 1 Then
   cotizacion = 1
 End If
 
If moneda = "DOL" Then
    totalotramoneda = total * cotizacion
    moneda = "D"
Else
    totalotramoneda = total / cotizacion
    moneda = "P"
End If
 
 
  numint = saca_ultnumero_int_comp("V")
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(suc)
  cl_compvta.actual (cc)
  cl_compvta.letra = letra
  cl_compvta.numcomp = Val(NUM)
  abreviatura = cl_compvta.abreviatura
      
  
  ep = "N"
  cp = "0000-00000000"
  contado = "N"
  cl_compvta.ACTUALIZA_NUMERADOR
      
      
      
      
  Set rs = New adodb.Recordset
  q = "select * from g8 where [id_actividad] = " & c_actividad.ItemData(c_actividad.ListIndex)
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    codact = rs("id_actividad")
    alicuotaib = rs("alicuota_ib")
    cuentaact = rs("cuenta_contable_venta")
  Else
       codact = 0
       alicuotaib = 0
       cuentaact = para.cuenta_ventas
  End If
  Set rs = Nothing
      
        
             
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
       idcli = Codc
      
      cn1.BeginTrans
       
       
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [saldo_impago02])"



QUERY = QUERY & " VALUES (" & numint & ", " & Val(suc) & ", " & Val(NUM) & ", '" & letra & "', " & cc & _
", " & idcli & ", '" & F & "', " & para.id_usuario & ", " & subtotal & ", " & nograbado & ", " & iva & ", " & total & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', 'Dup.Digital" & _
" ', " & cotizacion & ", " & Format(totalotramoneda, "#####0.00") & ", '" & moneda & "', 0, '" & cl_compvta.venta & "', '" & contado & "', " & percib + otrasperc & _
", 0, " & Val(t_perciva) & ", " & codact & ", " & tpercib & ", " & Val(t_alicuotaperciva) & ", 0, '" & F & "', 0, 0, ' ', ' ', ' ', 0, " & Val(suc) & _
", '" & Left$(cli, 50) & "', '" & Left$(Dire, 50) & "', '" & Left$(CUIT, 20) & "', '" & Left$(Loca, 50) & "', " & tiporespiva & ", " & total & ")"

  
            
       cn1.Execute QUERY
      
      Set cl_cli = Nothing
           
'actualizo producots

  Open a1 For Input As #1
  r = 1
  While Not EOF(1)
   Line Input #1, m
   tc = Val(Mid$(l, 1, 3))
   If tc < 100 Then
        cantidad = Val(Mid$(m, 32, 7) & "." & Mid$(m, 39, 5))
        codunidad = Val(Mid$(m, 44, 2))
        pu = Val(Mid$(m, 46, 13) & "." & Mid$(m, 59, 3))
        iva = Val(Mid$(m, 93, 13) & "." & Mid$(m, 106, 2))
        tasaiva = Val(Mid$(m, 108, 3) & "." & Mid$(m, 111, 2))
        importe = pu * cantidad
        texto = Mid$(m, 117, Len(m) - 117)
   Else
        cantidad = Val(Mid$(m, 33, 7) & "." & Mid$(m, 40, 5))
        codunidad = Val(Mid$(m, 45, 2))
        pu = Val(Mid$(m, 47, 13) & "." & Mid$(m, 60, 3))
        iva = Val(Mid$(m, 94, 13) & "." & Mid$(m, 107, 2))
        tasaiva = Val(Mid$(m, 109, 3) & "." & Mid$(m, 112, 2))
        importe = pu * cantidad
        texto = Mid$(m, 118, Len(m) - 118)
   
   
   
   End If
   
   
   Set rs = New adodb.Recordset
   q = "select * from g5 where [cod_afip] = " & codunidad
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
     unidad3 = rs("unidad")
   Else
     unidad3 = "Unidad"
   End If
   Set rs = Nothing
      puf = Format(pu * (1 + (tasaiva / 100)), "######0.00")
   cp = ""
   i = 0
   While i >= 0
      If Mid$(m, 125 + i, 1) <> "-" Then
           cp = cp & Mid$(m, 125 + i, 1)
             If i > 8 Then
               i = -1
             Else
               i = i + 1
             End If
          Else
            i = -1
          End If
          
   Wend

   
   If Val(cp) > 0 Then
     Set rs = New adodb.Recordset
     q = "select * from a2 where [id_producto] = " & Val(cp)
     rs.MaxRecords = 1
     rs.Open q, cn1
     If Not rs.EOF And Not rs.BOF Then
        codprod = rs("id_producto")
        
     Else
       codprod = 1
     End If
     Set rs = Nothing
   Else
     codprod = 1
   End If
     
   QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [pu_final])"
   QUERY = QUERY & " VALUES (" & numint & ", " & r & ", " & codprod & ", '" & Left$(RTrim$(texto), 50) & " ', " & cantidad & ", " & pu & ", " & importe & ", " & tasaiva & ", 0, 0, " & cantidad & ", '" & Left$(unidad3, 8) & "', " & puf & ")"
  
   cn1.Execute QUERY
  
   If Len(RTrim$(texto)) > 50 Then
     'grabo desc extra
     QUERY = "INSERT INTO vta_015([num_int], [RENGLON], [desc_ext], [cant_lineas])"
     QUERY = QUERY & " VALUES (" & numint & ", " & r & ", '" & Left$(Mid$(texto, 51, Len(texto) - 50), 50) & "', 1)"
     cn1.Execute QUERY
   End If
  
  
  
  
  
  'percepciones
  secuencia = 1
  If percib > 0 Then
          'agrego percepcion ibba
          QUERY = "INSERT INTO vta_016([num_int], [secuencia], [id_percepcion], [importe], [id_cuenta], [cod_regimen], [base_imponible], [alicuota])"
          QUERY = QUERY & " VALUES (" & numint & ", " & secuencia & ", " & 2 & ", " & percib & ", 0, 0," & subtotal & ",0)"
          cn1.Execute QUERY
           
          secuencia = secuencia + 1
  End If
  If tp = "S" Then
    'agrego el resto de las percepciones desde el archivo otras_perc.txt
         
       Open a2 For Input As #3
       While Not EOF(3)
       
            Line Input #3, w
             If tc < 100 Then
                tipoperc = Val(Mid$(w, 25, 2))
                totalperc = Val(Mid$(w, 80, 13) & "." & Mid$(l, 94, 2))
             Else
                tipoperc = Val(Mid$(w, 26, 2))
                totalperc = Val(Mid$(w, 81, 13) & "." & Mid$(l, 95, 2))
             End If
             
             Select Case tipoperc
               Case Is = 2  'ibba
                  tpp = 2
               Case Is = 17 'ib salta
                  tpp = 4
               Case Else
                  tpp = 4
             End Select
             
                   
                QUERY = "INSERT INTO vta_016([num_int], [secuencia], [id_percepcion], [importe], [id_cuenta], [cod_regimen], [base_imponible], [alicuota])"
                QUERY = QUERY & " VALUES (" & numint & ", " & secuencia & ", " & tpp & ", " & totalperc & ", 0, 0," & subtotal & ",0)"
                MsgBox (QUERY)
                cn1.Execute QUERY
            
            secuencia = secuencia + 1
    
      Wend
      Close #3
  End If
  
  
  r = r + 1
  
  
  
  
  Wend

  Close #1

  
  
  
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Importar Factura Electronica NI:" & numint & "', " & para.id_usuario & ", 'V', " & numint & ", '" & Now & "', '[" & cc & "] " & letra & " " & Format$(Val(suc), "0000") & "-" & Format$(Val(NUM), "00000000") & "', 11, " & idcli & ")"
  
  cn1.Execute QUERY
  
  cn1.CommitTrans
      
  
  'calculo los totales por iva
  Call verifica_tasa_iva(numint)
  

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Sub camino()
If Dir1 <> "C:\" Then
  t_camino = Dir1 & "\"
Else
  t_camino = Dir1
End If

End Sub



Private Sub Form_Load()
Call camino
Call armagrid
Call carga_actividades(c_actividad)
c_actividad.ListIndex = 0
Call barraesag(Me)
End Sub





Private Sub msf1_GotFocus()
StatusBar1.Panels.item(1) = "[Barra Espaciadora] Selecciona Comprobante - [F2] Agrega Descripcion"
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
  If msf1.TextMatrix(msf1.Row, 0) = "**" Then
     msf1.TextMatrix(msf1.Row, 0) = ""
  Else
     If msf1.TextMatrix(msf1.Row, 1) = "OK" Then
        msf1.TextMatrix(msf1.Row, 0) = "**"
     Else
        MsgBox ("El comprobante tiene error, no se puede importar")
     End If
  End If
End If


If KeyCode = vbKeyF2 Then
  t = InputBox$("Importa Facturas Electronicas", "Agregar Observacion", msf1.TextMatrix(msf1.Row, 7))
  If t <> "" Then
    msf1.TextMatrix(msf1.Row, 7) = t
    
  End If
End If



If KeyCode = vbKeyF9 Then
c = cuenta
If c > 0 Then

  J = MsgBox("Confirma Importar " & c & " Comprobantes", 4)
  If J = 6 Then
     espere.Show
     espere.Refresh
     r = 1
     While r <= msf1.Rows - 1
      If msf1.TextMatrix(r, 0) = "**" Then
         Call graba(r)
      End If
      r = r + 1
    Wend
    Unload espere
    Unload Me
  End If
Else
 MsgBox ("No hay comprobantes seleccionados")
End If

End If

End Sub


Private Sub msf1_LostFocus()
Call barraesag(Me)
End Sub
