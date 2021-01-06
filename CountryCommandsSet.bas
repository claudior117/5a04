Attribute VB_Name = "CountryCommandsSet"
' -----------------------------------------------------------------------------
' Pais: Argentina
' -----------------------------------------------------------------------------

' Tipo de impresoras en el pais
'Public Printers
'Printers = Array("DEMETER", "CERES-VESTA")


' -----------------------------------------------------------------------------
' variables globales para configuracion de comando en ejemplos
' -----------------------------------------------------------------------------

' Set de comandos
Public GET_FIRMWARE_VERSION As String
Public NUM_CAMPO_VERSION As Integer
Public NUM_CAMPO_VERSION_MAYOR As Integer
Public NUM_CAMPO_VERSION_MENOR As Integer

Public X_REPORT As String
Public Z_REPORT As String

Public TICKET_OPEN As String
Public TICKET_ITEM As String
Public TICKET_PAYMENT As String
Public TICKET_CLOSE As String

Public TICKET_NC_OPEN As String
Public TICKET_NC_ITEM As String
Public TICKET_NC_PAYMENT As String
Public TICKET_NC_CLOSE As String

Public DNF_OPEN As String
Public DNF_ITEM As String
Public DNF_CLOSE As String


' Datos variables de los campos suministrados por el software POS (fijos en este ejemplo solo para efectos demostrativos)
Public TICKET_ITEM_FIELDS As String
Public TICKET_PAYMENT_FIELDS As String
Public TICKET_CLOSE_FIELDS As String

Public TICKET_NC_OPEN_FIELDS As String
Public TICKET_NC_ITEM_FIELDS As String
Public TICKET_NC_PAYMENT_FIELDS As String
Public TICKET_NC_CLOSE_FIELDS As String

Public DNF_ITEM_FIELDS As String
Public DNF_CLOSE_FIELDS As String



Public Sub setCommands_init()
 GET_FIRMWARE_VERSION = ""
 NUM_CAMPO_VERSION = 0
 NUM_CAMPO_VERSION_MAYOR = 0
 NUM_CAMPO_VERSION_MENOR = 0

 X_REPORT = ""
 Z_REPORT = ""

 TICKET_OPEN = ""
 TICKET_ITEM = ""
 TICKET_PAYMENT = ""
 TICKET_CLOSE = ""

 TICKET_NC_OPEN = ""
 TICKET_NC_ITEM = ""
 TICKET_NC_PAYMENT = ""
 TICKET_NC_CLOSE = ""

 DNF_OPEN = ""
 DNF_ITEM = ""
 DNF_CLOSE = ""


' Datos variables de los campos suministrados por el software POS (fijos en este ejemplo solo para efectos demostrativos)
 TICKET_ITEM_FIELDS = ""
 TICKET_PAYMENT_FIELDS = ""
 TICKET_CLOSE_FIELDS = ""

 TICKET_NC_OPEN_FIELDS = ""
 TICKET_NC_ITEM_FIELDS = ""
 TICKET_NC_PAYMENT_FIELDS = ""
 TICKET_NC_CLOSE_FIELDS = ""

 DNF_ITEM_FIELDS = ""
 DNF_CLOSE_FIELDS = ""
End Sub



' -----------------------------------------------------------------------------
' Set de comando usados para el ejemplo
' -----------------------------------------------------------------------------
Public Sub SetData(tipoProtocolo As Integer, tipoImpresora As Integer)


    If (tipoProtocolo = 0) Then ' ProtocoloCompatible
        ' -----------------------------------------------------------------------------
        ' Protocolo compatible
        ' -----------------------------------------------------------------------------
        If (tipoImpresora = 0) Then     ' Impresora en 0 en la lista "DEMETER"
            ' Set de comandos
            GET_FIRMWARE_VERSION = "\x2a|N"
            NUM_CAMPO_VERSION = 0
            NUM_CAMPO_VERSION_MAYOR = 0
            NUM_CAMPO_VERSION_MENOR = 0

            X_REPORT = Chr(&H39) & "|X|P"
            Z_REPORT = "\x39|Z"

            TICKET_OPEN = "\x40|C"
            TICKET_ITEM = "\x42"
            TICKET_PAYMENT = "\x44"
            TICKET_CLOSE = "\x45"

            TICKET_NC_OPEN = "\x60"
            TICKET_NC_ITEM = "\x62"
            TICKET_NC_PAYMENT = "\x64"
            TICKET_NC_CLOSE = "\x65"
      
            DNF_OPEN = "\x48"
            DNF_ITEM = "\x49"
            DNF_CLOSE = "\x4A"

            ' Datos variables de los campos suministrados por el software POS (fijos en este ejemplo solo para efectos demostrativos)
            TICKET_ITEM_FIELDS = "|item|20000|1200|2100|M|1|0|0"
            TICKET_PAYMENT_FIELDS = "|Descripcion de pago|1000|T"
            TICKET_CLOSE_FIELDS = "|T"

            TICKET_NC_OPEN_FIELDS = "|M|C|A|1|P|12|I|I|(Nombre Cliente)||CUIT|27141670641|N|(Domicilio Cliente)|(Capital)|(Bs.As.)|(Numero de Factura)"
            TICKET_NC_ITEM_FIELDS = "|Descripcion de item|20000|1200|2100|M|00000|00000000|Linea 1 de descripcion|||0000|000000000000000"
            TICKET_NC_PAYMENT_FIELDS = "|(Nombre del medio pago)|100|T"
            TICKET_NC_CLOSE_FIELDS = "|M|A|FINAL"
      
            DNF_ITEM_FIELDS = "|Texto a imprimir"
            DNF_CLOSE_FIELDS = "|T"

        End If
    
        If (tipoImpresora = 1) Then     ' Impresora en 0 en la lista "CERES-VESTA"
            ' Set de comandos
            '----------------------------------------------------------------
            '       ESTE MODELO NO MANEJA ESTE PROTOCOLO DE COMUNICACIÓN
            '----------------------------------------------------------------
            GET_FIRMWARE_VERSION = ""
            NUM_CAMPO_VERSION = 0
            NUM_CAMPO_VERSION_MAYOR = 0
            NUM_CAMPO_VERSION_MENOR = 0

            X_REPORT = ""
            Z_REPORT = ""

            TICKET_OPEN = ""
            TICKET_ITEM = ""
            TICKET_PAYMENT = ""
            TICKET_CLOSE = ""

            TICKET_NC_OPEN = ""
            TICKET_NC_ITEM = ""
            TICKET_NC_PAYMENT = ""
            TICKET_NC_CLOSE = ""
      
            DNF_OPEN = ""
            DNF_ITEM = ""
            DNF_CLOSE = ""

            ' Datos variables de los campos suministrados por el software POS (fijos en este ejemplo solo para efectos demostrativos)
            TICKET_ITEM_FIELDS = ""
            TICKET_PAYMENT_FIELDS = ""
            TICKET_CLOSE_FIELDS = ""

            TICKET_NC_OPEN_FIELDS = ""
            TICKET_NC_ITEM_FIELDS = ""
            TICKET_NC_PAYMENT_FIELDS = ""
            TICKET_NC_CLOSE_FIELDS = ""
      
            DNF_ITEM_FIELDS = ""
            DNF_CLOSE_FIELDS = ""
        End If
    Else
        ' -----------------------------------------------------------------------------
        ' protocolo extendido
        ' -----------------------------------------------------------------------------
        If (tipoImpresora = 0) Then  ' Impresora en 0 en la lista "DEMETER"
            ' Set de comandos
            GET_FIRMWARE_VERSION = "020A|0000"
            NUM_CAMPO_VERSION = 1
            NUM_CAMPO_VERSION_MAYOR = 3
            NUM_CAMPO_VERSION_MENOR = 4

            X_REPORT = "0802|0C01"
            Z_REPORT = "0801|0C00"

            TICKET_OPEN = "0A01|0000"
            TICKET_ITEM = "0A02|0000"
            TICKET_PAYMENT = "0A05|0000"
            TICKET_CLOSE = "0A06|0013"

            TICKET_NC_OPEN = "0D01|0000"
            TICKET_NC_ITEM = "0D02|0000"
            TICKET_NC_PAYMENT = "0D05|0000"
            TICKET_NC_CLOSE = "0D06|0003"

            DNF_OPEN = "0E01|0000"
            DNF_ITEM = "0E02|0000"
            DNF_CLOSE = "0E06|0001"

            ' Datos variables de los campos suministrados por el software POS (fijos en este ejemplo solo para efectos demostrativos)
            TICKET_ITEM_FIELDS = "|Linea 1 de descripcion |Linea 2 de descripcion |Linea 3 de descripcion |Linea 4 de descripcion |Descripcion ITEM|10000|120000|2100||"
            TICKET_PAYMENT_FIELDS = "|Pago extra '1|Descripcion del pago|1000"
            TICKET_CLOSE_FIELDS = "||||||"        ' 6 campos

            TICKET_NC_OPEN_FIELDS = "|Nombre Comprador '1|Nombre Comprador '2|Domicilio Comprador '1|||T|30614104712|I|||081-0005-0007777"
            TICKET_NC_ITEM_FIELDS = "|Linea 1 de descripcion |Linea 2 de descripcion |Linea 3 de descripcion |Linea 4 de descripcion |Descripcion ITEM|10000|120000|2100||"
            TICKET_NC_PAYMENT_FIELDS = "|Pago extra '1|Descripcion del pago|1000"
            TICKET_NC_CLOSE_FIELDS = "|||||||1"       ' 7 campos

            DNF_ITEM_FIELDS = "|Texto a imprimir"
            DNF_CLOSE_FIELDS = "||||||"           ' 6 campos
        
        End If
    
        If (tipoImpresora = 1) Then ' Impresora en 1 en la lista  "CERES-VESTA"
            ' Set de comandos
            GET_FIRMWARE_VERSION = "020A|0000"
            NUM_CAMPO_VERSION = 1
            NUM_CAMPO_VERSION_MAYOR = 3
            NUM_CAMPO_VERSION_MENOR = 4

            X_REPORT = "0802|0C01"
            Z_REPORT = "0801|0000"
      
            TICKET_OPEN = "0A01|0080"
            TICKET_ITEM = "0A02|0000"
            TICKET_PAYMENT = "0A05|0000"
            TICKET_CLOSE = "0A06|0013"

            TICKET_NC_OPEN = "0A01|4000"
            TICKET_NC_ITEM = "0A02|0000"
            TICKET_NC_PAYMENT = "0A05|0000"
            TICKET_NC_CLOSE = "0A06|0013"

            DNF_OPEN = "0E01|0000"
            DNF_ITEM = "0E02|0000"
            DNF_CLOSE = "0E06|0001"

            ' Datos variables de los campos suministrados por el software POS (fijos en este ejemplo solo para efectos demostrativos)
            TICKET_ITEM_FIELDS = "|Linea 1 de descripcion |Linea 2 de descripcion |Linea 3 de descripcion |Linea 4 de descripcion |Descripcion ITEM|10000|120000|2100|||||1234567890|1|7"
            TICKET_PAYMENT_FIELDS = "|Pago extra '1|Pago extra '2|10|Otra forma de pago|Detalle de cupones|06|1000"
            TICKET_CLOSE_FIELDS = "|||||||"      ' 7 campos

            TICKET_NC_OPEN_FIELDS = ""
            TICKET_NC_ITEM_FIELDS = "|Linea 1 de descripcion |Linea 2 de descripcion |Linea 3 de descripcion |Linea 4 de descripcion |Descripcion ITEM|10000|120000|2100|||||1234567890|1|7"
            TICKET_NC_PAYMENT_FIELDS = "|Pago extra '1|Pago extra '2|10|Otra forma de pago|Detalle de cupones|06|1000"
            TICKET_NC_CLOSE_FIELDS = "|||||||"     ' 7 campos

            DNF_ITEM_FIELDS = "|Texto a imprimir"
            DNF_CLOSE_FIELDS = "||||||"           ' 6 campos
        End If

    End If
    
End Sub





