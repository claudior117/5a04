
  

   	* TipoDeComprobante
  
    tcNo_Fiscal = 0
    tcFactura_A = 1
    tcFactura_B = 2
    tcFactura_C = 3
    tcNota_Debito_A = 4
    tcNota_Debito_B = 5
    tcNota_Debito_C = 6
    tcNota_Credito_A = 7
    tcNota_Credito_B = 8
    tcNota_Credito_C = 9
    tcTicket = 10
  
	* Puertos

    pcCOM1 = 1
    pcCOM2 = 2
    pcCOM3 = 3
    pcCOM4 = 4
    pcCOM5 = 5
    pcCOM6 = 6
    pcCOM7 = 7
    pcCOM8 = 8
    pcCOM9 = 9
  

  * codigos de error

    errNoError = 0
    errControladorNoDisponible = 1
    errComandoInvalido = 2
    errParametroInvalido = 3
    errExcepcion = 4
    errMemoriaFiscal = 5
    errMemoriaTrabajo = 6
    errBateriaBaja = 7
    errComandoDesconocido = 8
    errDesbordamientoTotales = 9
    errMemoriaFiscalLlena = 10
    errMemoriaFiscalCasiLlena = 11
    errFallaImpresora = 13
    errImpresoraFueraLinea = 14
    errFaltaPapelDiario = 15
    errFaltaPapelTicket = 16
    errTapaImpresoraAbierta = 18
    errCajonCerradoOAusente = 19
    errCampoDatosInvalido = 20
  
	* Modelos de impresora

    modHasar715 = 0
    modHasar715v2 = 2
    modHasar615 = 3
    modHasar320 = 4
    modHasarPR4F = 5
    modHasarPR5F = 6
    modHasar950 = 7
    modHasar951 = 8
    modHasar441 = 9
    modHasar321 = 10
    modHasar322 = 11
    modHasar322v2 = 12
    modHasar330 = 13
    modHasar1120 = 14
    modHasarPL8F = 15
    modHasarPL8Fv2 = 16
    modHasarPL23 = 17
    modEpsonTM300AF = 18
    modEpsonTMU220AF = 19
    modEpsonTM2000 = 20
    modEpsonTM2000AFPlus = 21

	* Tipos de documento
	
    tdCUIT = 0
    tdDNI = 1
    tdPasaporte = 2
    tdCedula = 3
  
	* Responsabilidad ante IVA

    riResponsableInscripto = 0
    riMonotributo = 1
    riExento = 3
    riConsumidorFinal = 4
  	
 	
 	
 	TipoComprobante = tcFactura_A
 	
 	Fiscal = CreateObject("IFUniversal.Driver")

	lError = .F.	
    Fiscal:Modelo = modEpsonTMU220AF
    Fiscal:Puerto = 2
    Fiscal:Baudios = 9600

    Fiscal:Inicializar
    
    Fiscal:CancelarComprobante

    * Esto no se envia si la factura es a consumidor final
    If !lError Then
*!*	    function DatosCliente(const aNombre: WideString; aTipoDeDocumento: TipoDeDocumento;
*!*	                          const aDocumento: WideString; aResponsIVA: ResponsabilidadIVA;
*!*	                          const aDireccion: WideString): OLE_CANCELBOOL; dispid 211;
       lError = !Fiscal:DatosCliente("Abel Miranda", tdCUIT, "20939802593", riMonotributo, "Blanco Encalada 1204 5to A")
	EndIf
	
    If !lError Then
      lError = !Fiscal:AbrirComprobante(TipoComprobante)
    EndIf   
       
    If !lError Then
*!*	    function ImprimirItem(const aDescripcion: WideString; aCantidad: Double; aPrecio: Double;
*!*	                          aIVA: Double; aImpuestosInternos: Double): OLE_CANCELBOOL; dispid 207;
	   lError = !Fiscal:ImprimirItem("Item 1", 2, 100, 21, 0)
    EndIf

    If !lError Then
	   lError = !Fiscal:ImprimirItem("Item 1", 2, 100, 21, 0)
    EndIf

    If !lError Then
*!*	    function ImprimirDescuentoGeneral(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; dispid 208;
	   lError = !Fiscal:ImprimirDescuentoGeneral("Item 1", 10)
    EndIf

    If !lError Then
    	lError = !Fiscal:ImprimirPago("Visa", 100)
    EndIf

    If !lError Then
	   lError = !Fiscal:ImprimirPago("Efectivo", 100)
    EndIf

    Fiscal:CerrarComprobante

  If lError Then 
    MessageBox(Fiscal.ErrorDesc)
  Else
    MessageBox("Comprobante impreso exitosamente!")
  EndIf