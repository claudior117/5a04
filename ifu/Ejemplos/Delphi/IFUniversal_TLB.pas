unit IFUniversal_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 24/10/2018 10:57:24 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files (x86)\Javscript Driver Fiscal\ifu.dll (1)
// LIBID: {AF121C69-AB27-444F-9DB9-4260A7CBB41E}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  IFUniversalMajorVersion = 1;
  IFUniversalMinorVersion = 0;

  LIBID_IFUniversal: TGUID = '{AF121C69-AB27-444F-9DB9-4260A7CBB41E}';

  IID_IDriver: TGUID = '{00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}';
  CLASS_Driver: TGUID = '{536413FB-C017-4B59-8923-AE79800E3BB4}';
  IID_IObtenerDatosDeInicializacionRespuesta: TGUID = '{44C8E088-C222-4FC1-94ED-9395F5FE32C2}';
  CLASS_ObtenerDatosDeInicializacionRespuesta: TGUID = '{EF88ACD1-CD97-418F-A01B-B4657E28C6B2}';
  IID_ISubtotalRespuesta: TGUID = '{09BDCB7C-4945-4231-AB0C-628CF69E8561}';
  CLASS_SubtotalRespuesta: TGUID = '{27D2653D-A3D2-4037-A5AD-EF73A64A0C69}';
  IID_ICierreZTotales: TGUID = '{A7973DAB-A411-454D-927E-517037721A21}';
  CLASS_CierreZTotales: TGUID = '{F0C532B6-9FDC-4A80-BEC1-C9A064F5400D}';
  IID_IConsultarCapacidadZetasRespuesta: TGUID = '{1D33F62C-0EA9-44D4-8971-2333F441D7EE}';
  CLASS_ConsultarCapacidadZetasRespuesta: TGUID = '{D9C06EB3-5688-46D9-839F-8F265C41272F}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TipoDeComprobante
type
  TipoDeComprobante = TOleEnum;
const
  tcNo_Fiscal = $00000000;
  tcFactura_A = $00000001;
  tcFactura_B = $00000002;
  tcFactura_C = $00000003;
  tcNota_Debito_A = $00000004;
  tcNota_Debito_B = $00000005;
  tcNota_Debito_C = $00000006;
  tcNota_Credito_A = $00000007;
  tcNota_Credito_B = $00000008;
  tcNota_Credito_C = $00000009;
  tcTique = $0000000A;
  tcRemito = $0000000B;
  tcTiqueNotaCredito = $0000000C;
  tcRemitoX = $0000000D;
  tcReciboX = $0000000E;
  tcReciboA = $0000000F;
  tcReciboB = $00000010;
  tcReciboC = $00000011;

// Constants for enum PuertoCOM
type
  PuertoCOM = TOleEnum;
const
  pcCOM1 = $00000001;
  pcCOM2 = $00000002;
  pcCOM3 = $00000003;
  pcCOM4 = $00000004;
  pcCOM5 = $00000005;
  pcCOM6 = $00000006;
  pcCOM7 = $00000007;
  pcCOM8 = $00000008;
  pcCOM9 = $00000009;

// Constants for enum Baudio
type
  Baudio = TOleEnum;
const
  bd2400 = $00000960;
  bd4800 = $000012C0;
  bd9600 = $00002580;
  bd19200 = $00004B00;
  bd38400 = $00009600;
  bd57600 = $0000E100;
  bd115200 = $0001C200;

// Constants for enum ErrorNro
type
  ErrorNro = TOleEnum;
const
  errNoError = $00000000;
  errControladorNoDisponible = $00000001;
  errComandoInvalido = $00000002;
  errParametroInvalido = $00000003;
  errExcepcion = $00000004;
  errMemoriaFiscal = $00000005;
  errMemoriaTrabajo = $00000006;
  errBateriaBaja = $00000007;
  errComandoDesconocido = $00000008;
  errDesbordamientoTotales = $00000009;
  errMemoriaFiscalLlena = $0000000A;
  errMemoriaFiscalCasiLlena = $0000000B;
  errFallaImpresora = $0000000D;
  errImpresoraFueraLinea = $0000000E;
  errFaltaPapelDiario = $0000000F;
  errFaltaPapelTicket = $00000010;
  errTapaImpresoraAbierta = $00000012;
  errCajonCerradoOAusente = $00000013;
  errCampoDatosInvalido = $00000014;
  errCerrarJornada = $00000015;

// Constants for enum ModeloPrn
type
  ModeloPrn = TOleEnum;
const
  modHasar715 = $00000000;
  modHasar715v2 = $00000002;
  modHasar615 = $00000003;
  modHasar320 = $00000004;
  modHasarPR4F = $00000005;
  modHasarPR5F = $00000006;
  modHasar950 = $00000007;
  modHasar951 = $00000008;
  modHasar441 = $00000009;
  modHasar321 = $0000000A;
  modHasar322 = $0000000B;
  modHasar322v2 = $0000000C;
  modHasar330 = $0000000D;
  modHasar1120 = $0000000E;
  modHasarPL8F = $0000000F;
  modHasarPL8Fv2 = $00000010;
  modHasarPL23 = $00000011;
  modEpsonTM300AF = $00000012;
  modEpsonTMU220AF = $00000013;
  modEpsonTM2000 = $00000014;
  modEpsonTM2000AFPlus = $00000015;
  modEpsonLX300 = $00000016;
  modHasarPT1000F = $00000017;
  modEpsonTMT900FA = $00000018;
  modEpsonTMU220AFII = $00000019;

// Constants for enum TipoDeDocumento
type
  TipoDeDocumento = TOleEnum;
const
  tdCUIT = $00000000;
  tdDNI = $00000001;
  tdPasaporte = $00000002;
  tdCedula = $00000003;
  tdNinguno = $00000004;

// Constants for enum ResponsabilidadIVA
type
  ResponsabilidadIVA = TOleEnum;
const
  riResponsableInscripto = $00000000;
  riMonotributo = $00000001;
  riExento = $00000003;
  riConsumidorFinal = $00000004;
  riNoResponsable = $00000005;
  riNoCategorizado = $00000006;

// Constants for enum TiposTributos
type
  TiposTributos = TOleEnum;
const
  SinImpuesto = $00000000;
  ImpuestosNacionales = $00000001;
  ImpuestosProvinciales = $00000002;
  ImpuestosMunicipales = $00000003;
  ImpuestosInternos = $00000004;
  IIBB = $00000005;
  PercepcionIVA = $00000006;
  PercepcionIIBB = $00000007;
  PercepcionImpuestosMunicipales = $00000008;
  OtrasPercepciones = $00000009;
  ImpuestoInternoItem = $0000000A;
  OtrosTributos = $0000000B;

// Constants for enum CondicionesIVA
type
  CondicionesIVA = TOleEnum;
const
  NoGravado = $00000001;
  Exento = $00000002;
  Gravado = $00000007;

// Constants for enum UnidadesMedida
type
  UnidadesMedida = TOleEnum;
const
  SinDescripcion = $00000000;
  Kilo = $00000001;
  Metro = $00000002;
  MetroCuadrado = $00000003;
  MetroCubico = $00000004;
  Litro = $00000005;
  KWH = $00000006;
  Unidad = $00000007;
  Par = $00000008;
  Docena = $00000009;
  Quilate = $0000000A;
  Millar = $0000000B;
  MegaUInterActAntib = $0000000C;
  UnidadInternaActInmung = $0000000D;
  Gramo = $0000000E;
  Milimetro = $0000000F;
  MilimetroCubico = $00000010;
  Kilometro = $00000011;
  Hectolitro = $00000012;
  MegaUnidadIntActInmung = $00000013;
  Centimetro = $00000014;
  KilogramoActivo = $00000015;
  GramoActivo = $00000016;
  GramoBase = $00000017;
  UIACTHOR = $00000018;
  JuegoPaqueteMazoNaipes = $00000019;
  MUIACTHOR = $0000001A;
  CentimetroCubico = $0000001B;
  UIACTANT = $0000001C;
  Tonelada = $0000001D;
  DecametroCubico = $0000001E;
  HectometroCubico = $0000001F;
  KilometroCubico = $00000020;
  Microgramo = $00000021;
  Nanogramo = $00000022;
  Picogramo = $00000023;
  MUIACTANT = $00000024;
  UIACTIG = $00000025;
  Miligramo = $00000029;
  Mililitro = $0000002F;
  Curie = $00000030;
  Milicurie = $00000031;
  Microcurie = $00000032;
  UInterActHormonal = $00000033;
  MegaUInterActHor = $00000034;
  KilogramoBase = $00000035;
  Gruesa = $00000036;
  MUIACTIG = $00000037;
  KilogramoBruto = $0000003D;
  Pack = $0000003E;
  Horma = $0000003F;
  Donacion = $0000005A;
  Ajustes = $0000005B;
  Anulacion = $00000060;
  SenasAnticipos = $00000061;
  OtrasUnidades = $00000062;
  Bonificacion = $00000063;

// Constants for enum TiposPago
type
  TiposPago = TOleEnum;
const
  Cambio = $00000000;
  CartaDeCreditoDocumentario = $00000001;
  CartaDeCreditoSimple = $00000002;
  Cheque = $00000003;
  ChequeCancelatorios = $00000004;
  CreditoDocumentario = $00000005;
  CuentaCorriente = $00000006;
  Deposito = $00000007;
  Efectivo = $00000008;
  EndosoDeCheque = $00000009;
  FacturaDeCredito = $0000000A;
  GarantiaBancaria = $0000000B;
  Giro = $0000000C;
  LetraDeCambio = $0000000D;
  MedioDePagoDeComercioExterior = $0000000E;
  OrdenDePagoDocumentaria = $0000000F;
  OrdenDePagoSimple = $00000010;
  PagoContraReembolso = $00000011;
  RemesaDocumentaria = $00000012;
  RemesaSimple = $00000013;
  TarjetaDeCredito = $00000014;
  TarjetaDeDebito = $00000015;
  Ticket = $00000016;
  TransferenciaBancaria = $00000017;
  TransferenciaNoBancaria = $00000018;
  OtrosMediosPago = $00000063;

// Constants for enum TipoImpuestoInterno
type
  TipoImpuestoInterno = TOleEnum;
const
  tiFijo = $00000000;
  tiPorcentaje = $00000001;
  tiCoeficiente = $00000002;

// Constants for enum TipoReporteElectronico
type
  TipoReporteElectronico = TOleEnum;
const
  trFecha = $00000000;
  trNroCierre = $00000001;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IDriver = interface;
  IDriverDisp = dispinterface;
  IObtenerDatosDeInicializacionRespuesta = interface;
  IObtenerDatosDeInicializacionRespuestaDisp = dispinterface;
  ISubtotalRespuesta = interface;
  ISubtotalRespuestaDisp = dispinterface;
  ICierreZTotales = interface;
  ICierreZTotalesDisp = dispinterface;
  IConsultarCapacidadZetasRespuesta = interface;
  IConsultarCapacidadZetasRespuestaDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Driver = IDriver;
  ObtenerDatosDeInicializacionRespuesta = IObtenerDatosDeInicializacionRespuesta;
  SubtotalRespuesta = ISubtotalRespuesta;
  CierreZTotales = ICierreZTotales;
  ConsultarCapacidadZetasRespuesta = IConsultarCapacidadZetasRespuesta;


// *********************************************************************//
// Interface: IDriver
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}
// *********************************************************************//
  IDriver = interface(IDispatch)
    ['{00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}']
    function Get_Error: ErrorNro; safecall;
    function Get_ErrorDesc: WideString; safecall;
    function Get_Puerto: Integer; safecall;
    procedure Set_Puerto(Value: Integer); safecall;
    function Get_Baudios: Baudio; safecall;
    procedure Set_Baudios(Value: Baudio); safecall;
    function Get_Modelo: ModeloPrn; safecall;
    procedure Set_Modelo(Value: ModeloPrn); safecall;
    function AbrirComprobante(aTipoDeComprobante: TipoDeComprobante): OLE_CANCELBOOL; safecall;
    function ImprimirItem(const aDescripcion: WideString; aCantidad: Double; aPrecio: Double; 
                          aIVA: Double; aImpuestosInternos: Double): OLE_CANCELBOOL; safecall;
    function ImprimirItem2g(const Descripcion: WideString; Cantidad: Double; Precio: Double; 
                            IVA: Double; ImpuestosInternos: Double; g2CondicionIVA: CondicionesIVA; 
                            g2TipoImpuestoInterno: TipoImpuestoInterno; 
                            g2UnidadReferencia: Integer; const g2CodigoProducto: WideString; 
                            const g2CodigoInterno: WideString; g2UnidadMedida: UnidadesMedida): OLE_CANCELBOOL; safecall;
    function ImprimirDescuentoGeneral(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; safecall;
    function ImprimirPago(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; safecall;
    function ImprimirPago2g(const Descripcion: WideString; Monto: Double; 
                            const g2DescripcionAdicional: WideString; g2CodigoFormaPago: TiposPago; 
                            g2Cuotas: Integer; const g2Cupones: WideString; 
                            const g2Referencia: WideString): OLE_CANCELBOOL; safecall;
    procedure CerrarComprobante; safecall;
    function DatosCliente(const aNombre: WideString; aTipoDeDocumento: TipoDeDocumento; 
                          const aDocumento: WideString; aResponsIVA: ResponsabilidadIVA; 
                          const aDireccion: WideString): OLE_CANCELBOOL; safecall;
    procedure CancelarComprobante; safecall;
    function Inicializar: OLE_CANCELBOOL; safecall;
    function Get_TotalDocFiscales: Double; safecall;
    function CierreX: OLE_CANCELBOOL; safecall;
    function CierreZ: OLE_CANCELBOOL; safecall;
    function ImprimirTextoFiscal(const aTexto: WideString): OLE_CANCELBOOL; safecall;
    function InformarPercepcionGlobal(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; safecall;
    function InformarPercepcionIVA(const aDescripcion: WideString; aMonto: Double; aAlicuota: Double): OLE_CANCELBOOL; safecall;
    function Get_CbteEsFiscal: OLE_CANCELBOOL; safecall;
    function DocumentoDeReferencia(const aDocumento: WideString): OLE_CANCELBOOL; safecall;
    function UltimoComprobante(aTipoComprobante: TipoDeComprobante): Integer; safecall;
    function UltimoComprobanteCancelado: OLE_CANCELBOOL; safecall;
    function Get_ErroresEnExcepciones: OLE_CANCELBOOL; safecall;
    procedure Set_ErroresEnExcepciones(Value: OLE_CANCELBOOL); safecall;
    function Finalizar: OLE_CANCELBOOL; safecall;
    function DNFHFarmacias(const ObraSocial: WideString; const Coseguro1: WideString; 
                           const Coseguro2: WideString; const Coseguro3: WideString; 
                           const NroAfiliado: WideString; const NombreAfiliado: WideString; 
                           const FechaVencimientoCarnet: WideString; 
                           const DomicilioVend1: WideString; const DomicilioVend2: WideString; 
                           const NombreEstablecimiento: WideString; const NroInterno: WideString; 
                           const Nota1: WideString; const Nota2: WideString): OLE_CANCELBOOL; safecall;
    function CortarPapel: OLE_CANCELBOOL; safecall;
    function ImprimirTextoNoFiscal(const texto: WideString): OLE_CANCELBOOL; safecall;
    function ImprimirDescuentoUltimoItem(const Descripcion: WideString; Monto: Double): OLE_CANCELBOOL; safecall;
    function ReporteZFechas(const FechaInicial: WideString; const FechaFinal: WideString; 
                            Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL; safecall;
    function ReporteZNumeros(NroInicio: Integer; NroFin: Integer; Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL; safecall;
    function EspecificarEncabezado(Linea: Integer; const texto: WideString): OLE_CANCELBOOL; safecall;
    function EspecificarPie(Linea: Integer; const texto: WideString): OLE_CANCELBOOL; safecall;
    function CerrarComprobanteNumero(out Numero: Integer): OLE_CANCELBOOL; safecall;
    function Get_Copias: Integer; safecall;
    procedure Set_Copias(Value: Integer); safecall;
    function Get_Depurar: OLE_CANCELBOOL; safecall;
    procedure Set_Depurar(Value: OLE_CANCELBOOL); safecall;
    function ObtenerFechaHora: WideString; safecall;
    procedure AbrirCajon; safecall;
    function ObtenerDatosDeInicializacion: IObtenerDatosDeInicializacionRespuesta; safecall;
    function Subtotal: ISubtotalRespuesta; safecall;
    function ImprimirOtrosTributos(Codigo: TiposTributos; const Descripcion: WideString; 
                                   BaseImponible: Double; Importe: Double; Alicuota: Double): OLE_CANCELBOOL; safecall;
    procedure CargarLicencia(const Licencia: WideString); safecall;
    function Conectar(const DireccionIP: WideString; Puerto: Integer): OLE_CANCELBOOL; safecall;
    function DocumentoDeReferencia2g(TipoComprobante: TipoDeComprobante; const Documento: WideString): OLE_CANCELBOOL; safecall;
    function Get_CierreZTotales: CierreZTotales; safecall;
    function EspecificarFechaHora(const FechaHora: WideString): OLE_CANCELBOOL; safecall;
    function Get_PrecioBase: OLE_CANCELBOOL; safecall;
    procedure Set_PrecioBase(Value: OLE_CANCELBOOL); safecall;
    function CargarTransportista(const RazonSocial: WideString; Cuit: Double; 
                                 const Domicilio: WideString; const NombreChofer: WideString; 
                                 TipoDocumento: TipoDeDocumento; const NumeroDocumento: WideString; 
                                 const Dominio1: WideString; const Dominio2: WideString): OLE_CANCELBOOL; safecall;
    function ImprimirConceptoRecibo(const texto: WideString): OLE_CANCELBOOL; safecall;
    function EspecificarIngresosBrutos(const texto: WideString): OLE_CANCELBOOL; safecall;
    function EspecificarInicioActividades(const texto: WideString): OLE_CANCELBOOL; safecall;
    function ObtenerPrimerBloqueReporteElectronico(const RangoInicial: WideString; 
                                                   const RangoFinal: WideString; 
                                                   const NombreArchivo: WideString; 
                                                   TipoReporte: TipoReporteElectronico): OLE_CANCELBOOL; safecall;
    function ObtenerSiguienteBloqueReporteElectronico: OLE_CANCELBOOL; safecall;
    function ConsultarCapacidadZetas: OLE_CANCELBOOL; safecall;
    property Error: ErrorNro read Get_Error;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property Puerto: Integer read Get_Puerto write Set_Puerto;
    property Baudios: Baudio read Get_Baudios write Set_Baudios;
    property Modelo: ModeloPrn read Get_Modelo write Set_Modelo;
    property TotalDocFiscales: Double read Get_TotalDocFiscales;
    property CbteEsFiscal: OLE_CANCELBOOL read Get_CbteEsFiscal;
    property ErroresEnExcepciones: OLE_CANCELBOOL read Get_ErroresEnExcepciones write Set_ErroresEnExcepciones;
    property Copias: Integer read Get_Copias write Set_Copias;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
    property CierreZTotales: CierreZTotales read Get_CierreZTotales;
    property PrecioBase: OLE_CANCELBOOL read Get_PrecioBase write Set_PrecioBase;
  end;

// *********************************************************************//
// DispIntf:  IDriverDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}
// *********************************************************************//
  IDriverDisp = dispinterface
    ['{00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}']
    property Error: ErrorNro readonly dispid 203;
    property ErrorDesc: WideString readonly dispid 204;
    property Puerto: Integer dispid 201;
    property Baudios: Baudio dispid 205;
    property Modelo: ModeloPrn dispid 206;
    function AbrirComprobante(aTipoDeComprobante: TipoDeComprobante): OLE_CANCELBOOL; dispid 202;
    function ImprimirItem(const aDescripcion: WideString; aCantidad: Double; aPrecio: Double; 
                          aIVA: Double; aImpuestosInternos: Double): OLE_CANCELBOOL; dispid 207;
    function ImprimirItem2g(const Descripcion: WideString; Cantidad: Double; Precio: Double; 
                            IVA: Double; ImpuestosInternos: Double; g2CondicionIVA: CondicionesIVA; 
                            g2TipoImpuestoInterno: TipoImpuestoInterno; 
                            g2UnidadReferencia: Integer; const g2CodigoProducto: WideString; 
                            const g2CodigoInterno: WideString; g2UnidadMedida: UnidadesMedida): OLE_CANCELBOOL; dispid 242;
    function ImprimirDescuentoGeneral(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; dispid 208;
    function ImprimirPago(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; dispid 209;
    function ImprimirPago2g(const Descripcion: WideString; Monto: Double; 
                            const g2DescripcionAdicional: WideString; g2CodigoFormaPago: TiposPago; 
                            g2Cuotas: Integer; const g2Cupones: WideString; 
                            const g2Referencia: WideString): OLE_CANCELBOOL; dispid 243;
    procedure CerrarComprobante; dispid 210;
    function DatosCliente(const aNombre: WideString; aTipoDeDocumento: TipoDeDocumento; 
                          const aDocumento: WideString; aResponsIVA: ResponsabilidadIVA; 
                          const aDireccion: WideString): OLE_CANCELBOOL; dispid 211;
    procedure CancelarComprobante; dispid 212;
    function Inicializar: OLE_CANCELBOOL; dispid 213;
    property TotalDocFiscales: Double readonly dispid 214;
    function CierreX: OLE_CANCELBOOL; dispid 215;
    function CierreZ: OLE_CANCELBOOL; dispid 216;
    function ImprimirTextoFiscal(const aTexto: WideString): OLE_CANCELBOOL; dispid 217;
    function InformarPercepcionGlobal(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL; dispid 218;
    function InformarPercepcionIVA(const aDescripcion: WideString; aMonto: Double; aAlicuota: Double): OLE_CANCELBOOL; dispid 219;
    property CbteEsFiscal: OLE_CANCELBOOL readonly dispid 220;
    function DocumentoDeReferencia(const aDocumento: WideString): OLE_CANCELBOOL; dispid 221;
    function UltimoComprobante(aTipoComprobante: TipoDeComprobante): Integer; dispid 222;
    function UltimoComprobanteCancelado: OLE_CANCELBOOL; dispid 223;
    property ErroresEnExcepciones: OLE_CANCELBOOL dispid 224;
    function Finalizar: OLE_CANCELBOOL; dispid 225;
    function DNFHFarmacias(const ObraSocial: WideString; const Coseguro1: WideString; 
                           const Coseguro2: WideString; const Coseguro3: WideString; 
                           const NroAfiliado: WideString; const NombreAfiliado: WideString; 
                           const FechaVencimientoCarnet: WideString; 
                           const DomicilioVend1: WideString; const DomicilioVend2: WideString; 
                           const NombreEstablecimiento: WideString; const NroInterno: WideString; 
                           const Nota1: WideString; const Nota2: WideString): OLE_CANCELBOOL; dispid 226;
    function CortarPapel: OLE_CANCELBOOL; dispid 227;
    function ImprimirTextoNoFiscal(const texto: WideString): OLE_CANCELBOOL; dispid 228;
    function ImprimirDescuentoUltimoItem(const Descripcion: WideString; Monto: Double): OLE_CANCELBOOL; dispid 229;
    function ReporteZFechas(const FechaInicial: WideString; const FechaFinal: WideString; 
                            Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL; dispid 230;
    function ReporteZNumeros(NroInicio: Integer; NroFin: Integer; Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL; dispid 231;
    function EspecificarEncabezado(Linea: Integer; const texto: WideString): OLE_CANCELBOOL; dispid 232;
    function EspecificarPie(Linea: Integer; const texto: WideString): OLE_CANCELBOOL; dispid 233;
    function CerrarComprobanteNumero(out Numero: Integer): OLE_CANCELBOOL; dispid 234;
    property Copias: Integer dispid 235;
    property Depurar: OLE_CANCELBOOL dispid 236;
    function ObtenerFechaHora: WideString; dispid 237;
    procedure AbrirCajon; dispid 238;
    function ObtenerDatosDeInicializacion: IObtenerDatosDeInicializacionRespuesta; dispid 239;
    function Subtotal: ISubtotalRespuesta; dispid 240;
    function ImprimirOtrosTributos(Codigo: TiposTributos; const Descripcion: WideString; 
                                   BaseImponible: Double; Importe: Double; Alicuota: Double): OLE_CANCELBOOL; dispid 241;
    procedure CargarLicencia(const Licencia: WideString); dispid 244;
    function Conectar(const DireccionIP: WideString; Puerto: Integer): OLE_CANCELBOOL; dispid 245;
    function DocumentoDeReferencia2g(TipoComprobante: TipoDeComprobante; const Documento: WideString): OLE_CANCELBOOL; dispid 246;
    property CierreZTotales: CierreZTotales readonly dispid 247;
    function EspecificarFechaHora(const FechaHora: WideString): OLE_CANCELBOOL; dispid 248;
    property PrecioBase: OLE_CANCELBOOL dispid 249;
    function CargarTransportista(const RazonSocial: WideString; Cuit: Double; 
                                 const Domicilio: WideString; const NombreChofer: WideString; 
                                 TipoDocumento: TipoDeDocumento; const NumeroDocumento: WideString; 
                                 const Dominio1: WideString; const Dominio2: WideString): OLE_CANCELBOOL; dispid 250;
    function ImprimirConceptoRecibo(const texto: WideString): OLE_CANCELBOOL; dispid 251;
    function EspecificarIngresosBrutos(const texto: WideString): OLE_CANCELBOOL; dispid 252;
    function EspecificarInicioActividades(const texto: WideString): OLE_CANCELBOOL; dispid 253;
    function ObtenerPrimerBloqueReporteElectronico(const RangoInicial: WideString; 
                                                   const RangoFinal: WideString; 
                                                   const NombreArchivo: WideString; 
                                                   TipoReporte: TipoReporteElectronico): OLE_CANCELBOOL; dispid 254;
    function ObtenerSiguienteBloqueReporteElectronico: OLE_CANCELBOOL; dispid 255;
    function ConsultarCapacidadZetas: OLE_CANCELBOOL; dispid 256;
  end;

// *********************************************************************//
// Interface: IObtenerDatosDeInicializacionRespuesta
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {44C8E088-C222-4FC1-94ED-9395F5FE32C2}
// *********************************************************************//
  IObtenerDatosDeInicializacionRespuesta = interface(IDispatch)
    ['{44C8E088-C222-4FC1-94ED-9395F5FE32C2}']
    function Get_NroCUIT: WideString; safecall;
    function Get_RazonSocial: WideString; safecall;
    function Get_NroSerie: WideString; safecall;
    function Get_FechaInicializacion: WideString; safecall;
    function Get_NroPOS: WideString; safecall;
    function Get_FechaIniActividades: WideString; safecall;
    function Get_CodIngBrutos: WideString; safecall;
    function Get_RespIVA: WideString; safecall;
    function Get_Resultado: OLE_CANCELBOOL; safecall;
    property NroCUIT: WideString read Get_NroCUIT;
    property RazonSocial: WideString read Get_RazonSocial;
    property NroSerie: WideString read Get_NroSerie;
    property FechaInicializacion: WideString read Get_FechaInicializacion;
    property NroPOS: WideString read Get_NroPOS;
    property FechaIniActividades: WideString read Get_FechaIniActividades;
    property CodIngBrutos: WideString read Get_CodIngBrutos;
    property RespIVA: WideString read Get_RespIVA;
    property Resultado: OLE_CANCELBOOL read Get_Resultado;
  end;

// *********************************************************************//
// DispIntf:  IObtenerDatosDeInicializacionRespuestaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {44C8E088-C222-4FC1-94ED-9395F5FE32C2}
// *********************************************************************//
  IObtenerDatosDeInicializacionRespuestaDisp = dispinterface
    ['{44C8E088-C222-4FC1-94ED-9395F5FE32C2}']
    property NroCUIT: WideString readonly dispid 201;
    property RazonSocial: WideString readonly dispid 202;
    property NroSerie: WideString readonly dispid 203;
    property FechaInicializacion: WideString readonly dispid 204;
    property NroPOS: WideString readonly dispid 205;
    property FechaIniActividades: WideString readonly dispid 206;
    property CodIngBrutos: WideString readonly dispid 207;
    property RespIVA: WideString readonly dispid 208;
    property Resultado: OLE_CANCELBOOL readonly dispid 209;
  end;

// *********************************************************************//
// Interface: ISubtotalRespuesta
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {09BDCB7C-4945-4231-AB0C-628CF69E8561}
// *********************************************************************//
  ISubtotalRespuesta = interface(IDispatch)
    ['{09BDCB7C-4945-4231-AB0C-628CF69E8561}']
    function Get_CantidadItemsVendidos: Double; safecall;
    function Get_MontoVentas: Double; safecall;
    function Get_MontoIVA: Double; safecall;
    function Get_MontoPagado: Double; safecall;
    function Get_MontoIVANoInscripto: Double; safecall;
    function Get_MontoImpuestosInternos: Double; safecall;
    function Get_MontoNeto: Double; safecall;
    function Get_Resultado: OLE_CANCELBOOL; safecall;
    property CantidadItemsVendidos: Double read Get_CantidadItemsVendidos;
    property MontoVentas: Double read Get_MontoVentas;
    property MontoIVA: Double read Get_MontoIVA;
    property MontoPagado: Double read Get_MontoPagado;
    property MontoIVANoInscripto: Double read Get_MontoIVANoInscripto;
    property MontoImpuestosInternos: Double read Get_MontoImpuestosInternos;
    property MontoNeto: Double read Get_MontoNeto;
    property Resultado: OLE_CANCELBOOL read Get_Resultado;
  end;

// *********************************************************************//
// DispIntf:  ISubtotalRespuestaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {09BDCB7C-4945-4231-AB0C-628CF69E8561}
// *********************************************************************//
  ISubtotalRespuestaDisp = dispinterface
    ['{09BDCB7C-4945-4231-AB0C-628CF69E8561}']
    property CantidadItemsVendidos: Double readonly dispid 201;
    property MontoVentas: Double readonly dispid 202;
    property MontoIVA: Double readonly dispid 203;
    property MontoPagado: Double readonly dispid 204;
    property MontoIVANoInscripto: Double readonly dispid 205;
    property MontoImpuestosInternos: Double readonly dispid 206;
    property MontoNeto: Double readonly dispid 208;
    property Resultado: OLE_CANCELBOOL readonly dispid 207;
  end;

// *********************************************************************//
// Interface: ICierreZTotales
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A7973DAB-A411-454D-927E-517037721A21}
// *********************************************************************//
  ICierreZTotales = interface(IDispatch)
    ['{A7973DAB-A411-454D-927E-517037721A21}']
    function Get_FNDTotalVentas: Double; safecall;
    function Get_FNDTotalIVA: Double; safecall;
    function Get_FNDTotalImpuestosInternos: Double; safecall;
    function Get_FNDTotalOtrosTributos: Double; safecall;
    function Get_NCTotalVentas: Double; safecall;
    function Get_NCTotalIVA: Double; safecall;
    function Get_NCTotalImpuestosInternos: Double; safecall;
    function Get_NCTotalOtrosTributos: Double; safecall;
    function Get_NroCierre: Integer; safecall;
    property FNDTotalVentas: Double read Get_FNDTotalVentas;
    property FNDTotalIVA: Double read Get_FNDTotalIVA;
    property FNDTotalImpuestosInternos: Double read Get_FNDTotalImpuestosInternos;
    property FNDTotalOtrosTributos: Double read Get_FNDTotalOtrosTributos;
    property NCTotalVentas: Double read Get_NCTotalVentas;
    property NCTotalIVA: Double read Get_NCTotalIVA;
    property NCTotalImpuestosInternos: Double read Get_NCTotalImpuestosInternos;
    property NCTotalOtrosTributos: Double read Get_NCTotalOtrosTributos;
    property NroCierre: Integer read Get_NroCierre;
  end;

// *********************************************************************//
// DispIntf:  ICierreZTotalesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A7973DAB-A411-454D-927E-517037721A21}
// *********************************************************************//
  ICierreZTotalesDisp = dispinterface
    ['{A7973DAB-A411-454D-927E-517037721A21}']
    property FNDTotalVentas: Double readonly dispid 201;
    property FNDTotalIVA: Double readonly dispid 202;
    property FNDTotalImpuestosInternos: Double readonly dispid 203;
    property FNDTotalOtrosTributos: Double readonly dispid 204;
    property NCTotalVentas: Double readonly dispid 205;
    property NCTotalIVA: Double readonly dispid 206;
    property NCTotalImpuestosInternos: Double readonly dispid 207;
    property NCTotalOtrosTributos: Double readonly dispid 208;
    property NroCierre: Integer readonly dispid 209;
  end;

// *********************************************************************//
// Interface: IConsultarCapacidadZetasRespuesta
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {1D33F62C-0EA9-44D4-8971-2333F441D7EE}
// *********************************************************************//
  IConsultarCapacidadZetasRespuesta = interface(IDispatch)
    ['{1D33F62C-0EA9-44D4-8971-2333F441D7EE}']
    function Get_CantidadDeZetasRemanente: Integer; safecall;
    function Get_UltimaZeta: Integer; safecall;
    function Get_UltimaZetaBajada: Integer; safecall;
    function Get_UltimaZetaBorrable: Integer; safecall;
    property CantidadDeZetasRemanente: Integer read Get_CantidadDeZetasRemanente;
    property UltimaZeta: Integer read Get_UltimaZeta;
    property UltimaZetaBajada: Integer read Get_UltimaZetaBajada;
    property UltimaZetaBorrable: Integer read Get_UltimaZetaBorrable;
  end;

// *********************************************************************//
// DispIntf:  IConsultarCapacidadZetasRespuestaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {1D33F62C-0EA9-44D4-8971-2333F441D7EE}
// *********************************************************************//
  IConsultarCapacidadZetasRespuestaDisp = dispinterface
    ['{1D33F62C-0EA9-44D4-8971-2333F441D7EE}']
    property CantidadDeZetasRemanente: Integer readonly dispid 201;
    property UltimaZeta: Integer readonly dispid 202;
    property UltimaZetaBajada: Integer readonly dispid 203;
    property UltimaZetaBorrable: Integer readonly dispid 204;
  end;

// *********************************************************************//
// The Class CoDriver provides a Create and CreateRemote method to          
// create instances of the default interface IDriver exposed by              
// the CoClass Driver. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDriver = class
    class function Create: IDriver;
    class function CreateRemote(const MachineName: string): IDriver;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TDriver
// Help String      : Driver Object
// Default Interface: IDriver
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TDriverProperties= class;
{$ENDIF}
  TDriver = class(TOleServer)
  private
    FIntf:        IDriver;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TDriverProperties;
    function      GetServerProperties: TDriverProperties;
{$ENDIF}
    function      GetDefaultInterface: IDriver;
  protected
    procedure InitServerData; override;
    function Get_Error: ErrorNro;
    function Get_ErrorDesc: WideString;
    function Get_Puerto: Integer;
    procedure Set_Puerto(Value: Integer);
    function Get_Baudios: Baudio;
    procedure Set_Baudios(Value: Baudio);
    function Get_Modelo: ModeloPrn;
    procedure Set_Modelo(Value: ModeloPrn);
    function Get_TotalDocFiscales: Double;
    function Get_CbteEsFiscal: OLE_CANCELBOOL;
    function Get_ErroresEnExcepciones: OLE_CANCELBOOL;
    procedure Set_ErroresEnExcepciones(Value: OLE_CANCELBOOL);
    function Get_Copias: Integer;
    procedure Set_Copias(Value: Integer);
    function Get_Depurar: OLE_CANCELBOOL;
    procedure Set_Depurar(Value: OLE_CANCELBOOL);
    function Get_CierreZTotales: CierreZTotales;
    function Get_PrecioBase: OLE_CANCELBOOL;
    procedure Set_PrecioBase(Value: OLE_CANCELBOOL);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IDriver);
    procedure Disconnect; override;
    function AbrirComprobante(aTipoDeComprobante: TipoDeComprobante): OLE_CANCELBOOL;
    function ImprimirItem(const aDescripcion: WideString; aCantidad: Double; aPrecio: Double; 
                          aIVA: Double; aImpuestosInternos: Double): OLE_CANCELBOOL;
    function ImprimirItem2g(const Descripcion: WideString; Cantidad: Double; Precio: Double; 
                            IVA: Double; ImpuestosInternos: Double; g2CondicionIVA: CondicionesIVA; 
                            g2TipoImpuestoInterno: TipoImpuestoInterno; 
                            g2UnidadReferencia: Integer; const g2CodigoProducto: WideString; 
                            const g2CodigoInterno: WideString; g2UnidadMedida: UnidadesMedida): OLE_CANCELBOOL;
    function ImprimirDescuentoGeneral(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL;
    function ImprimirPago(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL;
    function ImprimirPago2g(const Descripcion: WideString; Monto: Double; 
                            const g2DescripcionAdicional: WideString; g2CodigoFormaPago: TiposPago; 
                            g2Cuotas: Integer; const g2Cupones: WideString; 
                            const g2Referencia: WideString): OLE_CANCELBOOL;
    procedure CerrarComprobante;
    function DatosCliente(const aNombre: WideString; aTipoDeDocumento: TipoDeDocumento; 
                          const aDocumento: WideString; aResponsIVA: ResponsabilidadIVA; 
                          const aDireccion: WideString): OLE_CANCELBOOL;
    procedure CancelarComprobante;
    function Inicializar: OLE_CANCELBOOL;
    function CierreX: OLE_CANCELBOOL;
    function CierreZ: OLE_CANCELBOOL;
    function ImprimirTextoFiscal(const aTexto: WideString): OLE_CANCELBOOL;
    function InformarPercepcionGlobal(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL;
    function InformarPercepcionIVA(const aDescripcion: WideString; aMonto: Double; aAlicuota: Double): OLE_CANCELBOOL;
    function DocumentoDeReferencia(const aDocumento: WideString): OLE_CANCELBOOL;
    function UltimoComprobante(aTipoComprobante: TipoDeComprobante): Integer;
    function UltimoComprobanteCancelado: OLE_CANCELBOOL;
    function Finalizar: OLE_CANCELBOOL;
    function DNFHFarmacias(const ObraSocial: WideString; const Coseguro1: WideString; 
                           const Coseguro2: WideString; const Coseguro3: WideString; 
                           const NroAfiliado: WideString; const NombreAfiliado: WideString; 
                           const FechaVencimientoCarnet: WideString; 
                           const DomicilioVend1: WideString; const DomicilioVend2: WideString; 
                           const NombreEstablecimiento: WideString; const NroInterno: WideString; 
                           const Nota1: WideString; const Nota2: WideString): OLE_CANCELBOOL;
    function CortarPapel: OLE_CANCELBOOL;
    function ImprimirTextoNoFiscal(const texto: WideString): OLE_CANCELBOOL;
    function ImprimirDescuentoUltimoItem(const Descripcion: WideString; Monto: Double): OLE_CANCELBOOL;
    function ReporteZFechas(const FechaInicial: WideString; const FechaFinal: WideString; 
                            Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL;
    function ReporteZNumeros(NroInicio: Integer; NroFin: Integer; Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL;
    function EspecificarEncabezado(Linea: Integer; const texto: WideString): OLE_CANCELBOOL;
    function EspecificarPie(Linea: Integer; const texto: WideString): OLE_CANCELBOOL;
    function CerrarComprobanteNumero(out Numero: Integer): OLE_CANCELBOOL;
    function ObtenerFechaHora: WideString;
    procedure AbrirCajon;
    function ObtenerDatosDeInicializacion: IObtenerDatosDeInicializacionRespuesta;
    function Subtotal: ISubtotalRespuesta;
    function ImprimirOtrosTributos(Codigo: TiposTributos; const Descripcion: WideString; 
                                   BaseImponible: Double; Importe: Double; Alicuota: Double): OLE_CANCELBOOL;
    procedure CargarLicencia(const Licencia: WideString);
    function Conectar(const DireccionIP: WideString; Puerto: Integer): OLE_CANCELBOOL;
    function DocumentoDeReferencia2g(TipoComprobante: TipoDeComprobante; const Documento: WideString): OLE_CANCELBOOL;
    function EspecificarFechaHora(const FechaHora: WideString): OLE_CANCELBOOL;
    function CargarTransportista(const RazonSocial: WideString; Cuit: Double; 
                                 const Domicilio: WideString; const NombreChofer: WideString; 
                                 TipoDocumento: TipoDeDocumento; const NumeroDocumento: WideString; 
                                 const Dominio1: WideString; const Dominio2: WideString): OLE_CANCELBOOL;
    function ImprimirConceptoRecibo(const texto: WideString): OLE_CANCELBOOL;
    function EspecificarIngresosBrutos(const texto: WideString): OLE_CANCELBOOL;
    function EspecificarInicioActividades(const texto: WideString): OLE_CANCELBOOL;
    function ObtenerPrimerBloqueReporteElectronico(const RangoInicial: WideString; 
                                                   const RangoFinal: WideString; 
                                                   const NombreArchivo: WideString; 
                                                   TipoReporte: TipoReporteElectronico): OLE_CANCELBOOL;
    function ObtenerSiguienteBloqueReporteElectronico: OLE_CANCELBOOL;
    function ConsultarCapacidadZetas: OLE_CANCELBOOL;
    property DefaultInterface: IDriver read GetDefaultInterface;
    property Error: ErrorNro read Get_Error;
    property ErrorDesc: WideString read Get_ErrorDesc;
    property TotalDocFiscales: Double read Get_TotalDocFiscales;
    property CbteEsFiscal: OLE_CANCELBOOL read Get_CbteEsFiscal;
    property CierreZTotales: CierreZTotales read Get_CierreZTotales;
    property Puerto: Integer read Get_Puerto write Set_Puerto;
    property Baudios: Baudio read Get_Baudios write Set_Baudios;
    property Modelo: ModeloPrn read Get_Modelo write Set_Modelo;
    property ErroresEnExcepciones: OLE_CANCELBOOL read Get_ErroresEnExcepciones write Set_ErroresEnExcepciones;
    property Copias: Integer read Get_Copias write Set_Copias;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
    property PrecioBase: OLE_CANCELBOOL read Get_PrecioBase write Set_PrecioBase;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TDriverProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TDriver
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TDriverProperties = class(TPersistent)
  private
    FServer:    TDriver;
    function    GetDefaultInterface: IDriver;
    constructor Create(AServer: TDriver);
  protected
    function Get_Error: ErrorNro;
    function Get_ErrorDesc: WideString;
    function Get_Puerto: Integer;
    procedure Set_Puerto(Value: Integer);
    function Get_Baudios: Baudio;
    procedure Set_Baudios(Value: Baudio);
    function Get_Modelo: ModeloPrn;
    procedure Set_Modelo(Value: ModeloPrn);
    function Get_TotalDocFiscales: Double;
    function Get_CbteEsFiscal: OLE_CANCELBOOL;
    function Get_ErroresEnExcepciones: OLE_CANCELBOOL;
    procedure Set_ErroresEnExcepciones(Value: OLE_CANCELBOOL);
    function Get_Copias: Integer;
    procedure Set_Copias(Value: Integer);
    function Get_Depurar: OLE_CANCELBOOL;
    procedure Set_Depurar(Value: OLE_CANCELBOOL);
    function Get_CierreZTotales: CierreZTotales;
    function Get_PrecioBase: OLE_CANCELBOOL;
    procedure Set_PrecioBase(Value: OLE_CANCELBOOL);
  public
    property DefaultInterface: IDriver read GetDefaultInterface;
  published
    property Puerto: Integer read Get_Puerto write Set_Puerto;
    property Baudios: Baudio read Get_Baudios write Set_Baudios;
    property Modelo: ModeloPrn read Get_Modelo write Set_Modelo;
    property ErroresEnExcepciones: OLE_CANCELBOOL read Get_ErroresEnExcepciones write Set_ErroresEnExcepciones;
    property Copias: Integer read Get_Copias write Set_Copias;
    property Depurar: OLE_CANCELBOOL read Get_Depurar write Set_Depurar;
    property PrecioBase: OLE_CANCELBOOL read Get_PrecioBase write Set_PrecioBase;
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoObtenerDatosDeInicializacionRespuesta provides a Create and CreateRemote method to          
// create instances of the default interface IObtenerDatosDeInicializacionRespuesta exposed by              
// the CoClass ObtenerDatosDeInicializacionRespuesta. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoObtenerDatosDeInicializacionRespuesta = class
    class function Create: IObtenerDatosDeInicializacionRespuesta;
    class function CreateRemote(const MachineName: string): IObtenerDatosDeInicializacionRespuesta;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TObtenerDatosDeInicializacionRespuesta
// Help String      : 
// Default Interface: IObtenerDatosDeInicializacionRespuesta
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TObtenerDatosDeInicializacionRespuestaProperties= class;
{$ENDIF}
  TObtenerDatosDeInicializacionRespuesta = class(TOleServer)
  private
    FIntf:        IObtenerDatosDeInicializacionRespuesta;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TObtenerDatosDeInicializacionRespuestaProperties;
    function      GetServerProperties: TObtenerDatosDeInicializacionRespuestaProperties;
{$ENDIF}
    function      GetDefaultInterface: IObtenerDatosDeInicializacionRespuesta;
  protected
    procedure InitServerData; override;
    function Get_NroCUIT: WideString;
    function Get_RazonSocial: WideString;
    function Get_NroSerie: WideString;
    function Get_FechaInicializacion: WideString;
    function Get_NroPOS: WideString;
    function Get_FechaIniActividades: WideString;
    function Get_CodIngBrutos: WideString;
    function Get_RespIVA: WideString;
    function Get_Resultado: OLE_CANCELBOOL;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IObtenerDatosDeInicializacionRespuesta);
    procedure Disconnect; override;
    property DefaultInterface: IObtenerDatosDeInicializacionRespuesta read GetDefaultInterface;
    property NroCUIT: WideString read Get_NroCUIT;
    property RazonSocial: WideString read Get_RazonSocial;
    property NroSerie: WideString read Get_NroSerie;
    property FechaInicializacion: WideString read Get_FechaInicializacion;
    property NroPOS: WideString read Get_NroPOS;
    property FechaIniActividades: WideString read Get_FechaIniActividades;
    property CodIngBrutos: WideString read Get_CodIngBrutos;
    property RespIVA: WideString read Get_RespIVA;
    property Resultado: OLE_CANCELBOOL read Get_Resultado;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TObtenerDatosDeInicializacionRespuestaProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TObtenerDatosDeInicializacionRespuesta
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TObtenerDatosDeInicializacionRespuestaProperties = class(TPersistent)
  private
    FServer:    TObtenerDatosDeInicializacionRespuesta;
    function    GetDefaultInterface: IObtenerDatosDeInicializacionRespuesta;
    constructor Create(AServer: TObtenerDatosDeInicializacionRespuesta);
  protected
    function Get_NroCUIT: WideString;
    function Get_RazonSocial: WideString;
    function Get_NroSerie: WideString;
    function Get_FechaInicializacion: WideString;
    function Get_NroPOS: WideString;
    function Get_FechaIniActividades: WideString;
    function Get_CodIngBrutos: WideString;
    function Get_RespIVA: WideString;
    function Get_Resultado: OLE_CANCELBOOL;
  public
    property DefaultInterface: IObtenerDatosDeInicializacionRespuesta read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoSubtotalRespuesta provides a Create and CreateRemote method to          
// create instances of the default interface ISubtotalRespuesta exposed by              
// the CoClass SubtotalRespuesta. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoSubtotalRespuesta = class
    class function Create: ISubtotalRespuesta;
    class function CreateRemote(const MachineName: string): ISubtotalRespuesta;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TSubtotalRespuesta
// Help String      : 
// Default Interface: ISubtotalRespuesta
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TSubtotalRespuestaProperties= class;
{$ENDIF}
  TSubtotalRespuesta = class(TOleServer)
  private
    FIntf:        ISubtotalRespuesta;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TSubtotalRespuestaProperties;
    function      GetServerProperties: TSubtotalRespuestaProperties;
{$ENDIF}
    function      GetDefaultInterface: ISubtotalRespuesta;
  protected
    procedure InitServerData; override;
    function Get_CantidadItemsVendidos: Double;
    function Get_MontoVentas: Double;
    function Get_MontoIVA: Double;
    function Get_MontoPagado: Double;
    function Get_MontoIVANoInscripto: Double;
    function Get_MontoImpuestosInternos: Double;
    function Get_MontoNeto: Double;
    function Get_Resultado: OLE_CANCELBOOL;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ISubtotalRespuesta);
    procedure Disconnect; override;
    property DefaultInterface: ISubtotalRespuesta read GetDefaultInterface;
    property CantidadItemsVendidos: Double read Get_CantidadItemsVendidos;
    property MontoVentas: Double read Get_MontoVentas;
    property MontoIVA: Double read Get_MontoIVA;
    property MontoPagado: Double read Get_MontoPagado;
    property MontoIVANoInscripto: Double read Get_MontoIVANoInscripto;
    property MontoImpuestosInternos: Double read Get_MontoImpuestosInternos;
    property MontoNeto: Double read Get_MontoNeto;
    property Resultado: OLE_CANCELBOOL read Get_Resultado;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TSubtotalRespuestaProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TSubtotalRespuesta
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TSubtotalRespuestaProperties = class(TPersistent)
  private
    FServer:    TSubtotalRespuesta;
    function    GetDefaultInterface: ISubtotalRespuesta;
    constructor Create(AServer: TSubtotalRespuesta);
  protected
    function Get_CantidadItemsVendidos: Double;
    function Get_MontoVentas: Double;
    function Get_MontoIVA: Double;
    function Get_MontoPagado: Double;
    function Get_MontoIVANoInscripto: Double;
    function Get_MontoImpuestosInternos: Double;
    function Get_MontoNeto: Double;
    function Get_Resultado: OLE_CANCELBOOL;
  public
    property DefaultInterface: ISubtotalRespuesta read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoCierreZTotales provides a Create and CreateRemote method to          
// create instances of the default interface ICierreZTotales exposed by              
// the CoClass CierreZTotales. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCierreZTotales = class
    class function Create: ICierreZTotales;
    class function CreateRemote(const MachineName: string): ICierreZTotales;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TCierreZTotales
// Help String      : 
// Default Interface: ICierreZTotales
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TCierreZTotalesProperties= class;
{$ENDIF}
  TCierreZTotales = class(TOleServer)
  private
    FIntf:        ICierreZTotales;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TCierreZTotalesProperties;
    function      GetServerProperties: TCierreZTotalesProperties;
{$ENDIF}
    function      GetDefaultInterface: ICierreZTotales;
  protected
    procedure InitServerData; override;
    function Get_FNDTotalVentas: Double;
    function Get_FNDTotalIVA: Double;
    function Get_FNDTotalImpuestosInternos: Double;
    function Get_FNDTotalOtrosTributos: Double;
    function Get_NCTotalVentas: Double;
    function Get_NCTotalIVA: Double;
    function Get_NCTotalImpuestosInternos: Double;
    function Get_NCTotalOtrosTributos: Double;
    function Get_NroCierre: Integer;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ICierreZTotales);
    procedure Disconnect; override;
    property DefaultInterface: ICierreZTotales read GetDefaultInterface;
    property FNDTotalVentas: Double read Get_FNDTotalVentas;
    property FNDTotalIVA: Double read Get_FNDTotalIVA;
    property FNDTotalImpuestosInternos: Double read Get_FNDTotalImpuestosInternos;
    property FNDTotalOtrosTributos: Double read Get_FNDTotalOtrosTributos;
    property NCTotalVentas: Double read Get_NCTotalVentas;
    property NCTotalIVA: Double read Get_NCTotalIVA;
    property NCTotalImpuestosInternos: Double read Get_NCTotalImpuestosInternos;
    property NCTotalOtrosTributos: Double read Get_NCTotalOtrosTributos;
    property NroCierre: Integer read Get_NroCierre;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TCierreZTotalesProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TCierreZTotales
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TCierreZTotalesProperties = class(TPersistent)
  private
    FServer:    TCierreZTotales;
    function    GetDefaultInterface: ICierreZTotales;
    constructor Create(AServer: TCierreZTotales);
  protected
    function Get_FNDTotalVentas: Double;
    function Get_FNDTotalIVA: Double;
    function Get_FNDTotalImpuestosInternos: Double;
    function Get_FNDTotalOtrosTributos: Double;
    function Get_NCTotalVentas: Double;
    function Get_NCTotalIVA: Double;
    function Get_NCTotalImpuestosInternos: Double;
    function Get_NCTotalOtrosTributos: Double;
    function Get_NroCierre: Integer;
  public
    property DefaultInterface: ICierreZTotales read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoConsultarCapacidadZetasRespuesta provides a Create and CreateRemote method to          
// create instances of the default interface IConsultarCapacidadZetasRespuesta exposed by              
// the CoClass ConsultarCapacidadZetasRespuesta. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoConsultarCapacidadZetasRespuesta = class
    class function Create: IConsultarCapacidadZetasRespuesta;
    class function CreateRemote(const MachineName: string): IConsultarCapacidadZetasRespuesta;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TConsultarCapacidadZetasRespuesta
// Help String      : 
// Default Interface: IConsultarCapacidadZetasRespuesta
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TConsultarCapacidadZetasRespuestaProperties= class;
{$ENDIF}
  TConsultarCapacidadZetasRespuesta = class(TOleServer)
  private
    FIntf:        IConsultarCapacidadZetasRespuesta;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TConsultarCapacidadZetasRespuestaProperties;
    function      GetServerProperties: TConsultarCapacidadZetasRespuestaProperties;
{$ENDIF}
    function      GetDefaultInterface: IConsultarCapacidadZetasRespuesta;
  protected
    procedure InitServerData; override;
    function Get_CantidadDeZetasRemanente: Integer;
    function Get_UltimaZeta: Integer;
    function Get_UltimaZetaBajada: Integer;
    function Get_UltimaZetaBorrable: Integer;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IConsultarCapacidadZetasRespuesta);
    procedure Disconnect; override;
    property DefaultInterface: IConsultarCapacidadZetasRespuesta read GetDefaultInterface;
    property CantidadDeZetasRemanente: Integer read Get_CantidadDeZetasRemanente;
    property UltimaZeta: Integer read Get_UltimaZeta;
    property UltimaZetaBajada: Integer read Get_UltimaZetaBajada;
    property UltimaZetaBorrable: Integer read Get_UltimaZetaBorrable;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TConsultarCapacidadZetasRespuestaProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TConsultarCapacidadZetasRespuesta
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TConsultarCapacidadZetasRespuestaProperties = class(TPersistent)
  private
    FServer:    TConsultarCapacidadZetasRespuesta;
    function    GetDefaultInterface: IConsultarCapacidadZetasRespuesta;
    constructor Create(AServer: TConsultarCapacidadZetasRespuesta);
  protected
    function Get_CantidadDeZetasRemanente: Integer;
    function Get_UltimaZeta: Integer;
    function Get_UltimaZetaBajada: Integer;
    function Get_UltimaZetaBorrable: Integer;
  public
    property DefaultInterface: IConsultarCapacidadZetasRespuesta read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoDriver.Create: IDriver;
begin
  Result := CreateComObject(CLASS_Driver) as IDriver;
end;

class function CoDriver.CreateRemote(const MachineName: string): IDriver;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Driver) as IDriver;
end;

procedure TDriver.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{536413FB-C017-4B59-8923-AE79800E3BB4}';
    IntfIID:   '{00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TDriver.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IDriver;
  end;
end;

procedure TDriver.ConnectTo(svrIntf: IDriver);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TDriver.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TDriver.GetDefaultInterface: IDriver;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TDriver.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TDriverProperties.Create(Self);
{$ENDIF}
end;

destructor TDriver.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TDriver.GetServerProperties: TDriverProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TDriver.Get_Error: ErrorNro;
begin
    Result := DefaultInterface.Error;
end;

function TDriver.Get_ErrorDesc: WideString;
begin
    Result := DefaultInterface.ErrorDesc;
end;

function TDriver.Get_Puerto: Integer;
begin
    Result := DefaultInterface.Puerto;
end;

procedure TDriver.Set_Puerto(Value: Integer);
begin
  DefaultInterface.Set_Puerto(Value);
end;

function TDriver.Get_Baudios: Baudio;
begin
    Result := DefaultInterface.Baudios;
end;

procedure TDriver.Set_Baudios(Value: Baudio);
begin
  DefaultInterface.Set_Baudios(Value);
end;

function TDriver.Get_Modelo: ModeloPrn;
begin
    Result := DefaultInterface.Modelo;
end;

procedure TDriver.Set_Modelo(Value: ModeloPrn);
begin
  DefaultInterface.Set_Modelo(Value);
end;

function TDriver.Get_TotalDocFiscales: Double;
begin
    Result := DefaultInterface.TotalDocFiscales;
end;

function TDriver.Get_CbteEsFiscal: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.CbteEsFiscal;
end;

function TDriver.Get_ErroresEnExcepciones: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.ErroresEnExcepciones;
end;

procedure TDriver.Set_ErroresEnExcepciones(Value: OLE_CANCELBOOL);
begin
  DefaultInterface.Set_ErroresEnExcepciones(Value);
end;

function TDriver.Get_Copias: Integer;
begin
    Result := DefaultInterface.Copias;
end;

procedure TDriver.Set_Copias(Value: Integer);
begin
  DefaultInterface.Set_Copias(Value);
end;

function TDriver.Get_Depurar: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.Depurar;
end;

procedure TDriver.Set_Depurar(Value: OLE_CANCELBOOL);
begin
  DefaultInterface.Set_Depurar(Value);
end;

function TDriver.Get_CierreZTotales: CierreZTotales;
begin
    Result := DefaultInterface.CierreZTotales;
end;

function TDriver.Get_PrecioBase: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.PrecioBase;
end;

procedure TDriver.Set_PrecioBase(Value: OLE_CANCELBOOL);
begin
  DefaultInterface.Set_PrecioBase(Value);
end;

function TDriver.AbrirComprobante(aTipoDeComprobante: TipoDeComprobante): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.AbrirComprobante(aTipoDeComprobante);
end;

function TDriver.ImprimirItem(const aDescripcion: WideString; aCantidad: Double; aPrecio: Double; 
                              aIVA: Double; aImpuestosInternos: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirItem(aDescripcion, aCantidad, aPrecio, aIVA, aImpuestosInternos);
end;

function TDriver.ImprimirItem2g(const Descripcion: WideString; Cantidad: Double; Precio: Double; 
                                IVA: Double; ImpuestosInternos: Double; 
                                g2CondicionIVA: CondicionesIVA; 
                                g2TipoImpuestoInterno: TipoImpuestoInterno; 
                                g2UnidadReferencia: Integer; const g2CodigoProducto: WideString; 
                                const g2CodigoInterno: WideString; g2UnidadMedida: UnidadesMedida): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirItem2g(Descripcion, Cantidad, Precio, IVA, ImpuestosInternos, 
                                            g2CondicionIVA, g2TipoImpuestoInterno, 
                                            g2UnidadReferencia, g2CodigoProducto, g2CodigoInterno, 
                                            g2UnidadMedida);
end;

function TDriver.ImprimirDescuentoGeneral(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirDescuentoGeneral(aDescripcion, aMonto);
end;

function TDriver.ImprimirPago(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirPago(aDescripcion, aMonto);
end;

function TDriver.ImprimirPago2g(const Descripcion: WideString; Monto: Double; 
                                const g2DescripcionAdicional: WideString; 
                                g2CodigoFormaPago: TiposPago; g2Cuotas: Integer; 
                                const g2Cupones: WideString; const g2Referencia: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirPago2g(Descripcion, Monto, g2DescripcionAdicional, 
                                            g2CodigoFormaPago, g2Cuotas, g2Cupones, g2Referencia);
end;

procedure TDriver.CerrarComprobante;
begin
  DefaultInterface.CerrarComprobante;
end;

function TDriver.DatosCliente(const aNombre: WideString; aTipoDeDocumento: TipoDeDocumento; 
                              const aDocumento: WideString; aResponsIVA: ResponsabilidadIVA; 
                              const aDireccion: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.DatosCliente(aNombre, aTipoDeDocumento, aDocumento, aResponsIVA, 
                                          aDireccion);
end;

procedure TDriver.CancelarComprobante;
begin
  DefaultInterface.CancelarComprobante;
end;

function TDriver.Inicializar: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.Inicializar;
end;

function TDriver.CierreX: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.CierreX;
end;

function TDriver.CierreZ: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.CierreZ;
end;

function TDriver.ImprimirTextoFiscal(const aTexto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirTextoFiscal(aTexto);
end;

function TDriver.InformarPercepcionGlobal(const aDescripcion: WideString; aMonto: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.InformarPercepcionGlobal(aDescripcion, aMonto);
end;

function TDriver.InformarPercepcionIVA(const aDescripcion: WideString; aMonto: Double; 
                                       aAlicuota: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.InformarPercepcionIVA(aDescripcion, aMonto, aAlicuota);
end;

function TDriver.DocumentoDeReferencia(const aDocumento: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.DocumentoDeReferencia(aDocumento);
end;

function TDriver.UltimoComprobante(aTipoComprobante: TipoDeComprobante): Integer;
begin
  Result := DefaultInterface.UltimoComprobante(aTipoComprobante);
end;

function TDriver.UltimoComprobanteCancelado: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.UltimoComprobanteCancelado;
end;

function TDriver.Finalizar: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.Finalizar;
end;

function TDriver.DNFHFarmacias(const ObraSocial: WideString; const Coseguro1: WideString; 
                               const Coseguro2: WideString; const Coseguro3: WideString; 
                               const NroAfiliado: WideString; const NombreAfiliado: WideString; 
                               const FechaVencimientoCarnet: WideString; 
                               const DomicilioVend1: WideString; const DomicilioVend2: WideString; 
                               const NombreEstablecimiento: WideString; 
                               const NroInterno: WideString; const Nota1: WideString; 
                               const Nota2: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.DNFHFarmacias(ObraSocial, Coseguro1, Coseguro2, Coseguro3, 
                                           NroAfiliado, NombreAfiliado, FechaVencimientoCarnet, 
                                           DomicilioVend1, DomicilioVend2, NombreEstablecimiento, 
                                           NroInterno, Nota1, Nota2);
end;

function TDriver.CortarPapel: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.CortarPapel;
end;

function TDriver.ImprimirTextoNoFiscal(const texto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirTextoNoFiscal(texto);
end;

function TDriver.ImprimirDescuentoUltimoItem(const Descripcion: WideString; Monto: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirDescuentoUltimoItem(Descripcion, Monto);
end;

function TDriver.ReporteZFechas(const FechaInicial: WideString; const FechaFinal: WideString; 
                                Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ReporteZFechas(FechaInicial, FechaFinal, Detallado);
end;

function TDriver.ReporteZNumeros(NroInicio: Integer; NroFin: Integer; Detallado: OLE_CANCELBOOL): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ReporteZNumeros(NroInicio, NroFin, Detallado);
end;

function TDriver.EspecificarEncabezado(Linea: Integer; const texto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.EspecificarEncabezado(Linea, texto);
end;

function TDriver.EspecificarPie(Linea: Integer; const texto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.EspecificarPie(Linea, texto);
end;

function TDriver.CerrarComprobanteNumero(out Numero: Integer): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.CerrarComprobanteNumero(Numero);
end;

function TDriver.ObtenerFechaHora: WideString;
begin
  Result := DefaultInterface.ObtenerFechaHora;
end;

procedure TDriver.AbrirCajon;
begin
  DefaultInterface.AbrirCajon;
end;

function TDriver.ObtenerDatosDeInicializacion: IObtenerDatosDeInicializacionRespuesta;
begin
  Result := DefaultInterface.ObtenerDatosDeInicializacion;
end;

function TDriver.Subtotal: ISubtotalRespuesta;
begin
  Result := DefaultInterface.Subtotal;
end;

function TDriver.ImprimirOtrosTributos(Codigo: TiposTributos; const Descripcion: WideString; 
                                       BaseImponible: Double; Importe: Double; Alicuota: Double): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirOtrosTributos(Codigo, Descripcion, BaseImponible, Importe, 
                                                   Alicuota);
end;

procedure TDriver.CargarLicencia(const Licencia: WideString);
begin
  DefaultInterface.CargarLicencia(Licencia);
end;

function TDriver.Conectar(const DireccionIP: WideString; Puerto: Integer): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.Conectar(DireccionIP, Puerto);
end;

function TDriver.DocumentoDeReferencia2g(TipoComprobante: TipoDeComprobante; 
                                         const Documento: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.DocumentoDeReferencia2g(TipoComprobante, Documento);
end;

function TDriver.EspecificarFechaHora(const FechaHora: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.EspecificarFechaHora(FechaHora);
end;

function TDriver.CargarTransportista(const RazonSocial: WideString; Cuit: Double; 
                                     const Domicilio: WideString; const NombreChofer: WideString; 
                                     TipoDocumento: TipoDeDocumento; 
                                     const NumeroDocumento: WideString; const Dominio1: WideString; 
                                     const Dominio2: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.CargarTransportista(RazonSocial, Cuit, Domicilio, NombreChofer, 
                                                 TipoDocumento, NumeroDocumento, Dominio1, Dominio2);
end;

function TDriver.ImprimirConceptoRecibo(const texto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ImprimirConceptoRecibo(texto);
end;

function TDriver.EspecificarIngresosBrutos(const texto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.EspecificarIngresosBrutos(texto);
end;

function TDriver.EspecificarInicioActividades(const texto: WideString): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.EspecificarInicioActividades(texto);
end;

function TDriver.ObtenerPrimerBloqueReporteElectronico(const RangoInicial: WideString; 
                                                       const RangoFinal: WideString; 
                                                       const NombreArchivo: WideString; 
                                                       TipoReporte: TipoReporteElectronico): OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ObtenerPrimerBloqueReporteElectronico(RangoInicial, RangoFinal, 
                                                                   NombreArchivo, TipoReporte);
end;

function TDriver.ObtenerSiguienteBloqueReporteElectronico: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ObtenerSiguienteBloqueReporteElectronico;
end;

function TDriver.ConsultarCapacidadZetas: OLE_CANCELBOOL;
begin
  Result := DefaultInterface.ConsultarCapacidadZetas;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TDriverProperties.Create(AServer: TDriver);
begin
  inherited Create;
  FServer := AServer;
end;

function TDriverProperties.GetDefaultInterface: IDriver;
begin
  Result := FServer.DefaultInterface;
end;

function TDriverProperties.Get_Error: ErrorNro;
begin
    Result := DefaultInterface.Error;
end;

function TDriverProperties.Get_ErrorDesc: WideString;
begin
    Result := DefaultInterface.ErrorDesc;
end;

function TDriverProperties.Get_Puerto: Integer;
begin
    Result := DefaultInterface.Puerto;
end;

procedure TDriverProperties.Set_Puerto(Value: Integer);
begin
  DefaultInterface.Set_Puerto(Value);
end;

function TDriverProperties.Get_Baudios: Baudio;
begin
    Result := DefaultInterface.Baudios;
end;

procedure TDriverProperties.Set_Baudios(Value: Baudio);
begin
  DefaultInterface.Set_Baudios(Value);
end;

function TDriverProperties.Get_Modelo: ModeloPrn;
begin
    Result := DefaultInterface.Modelo;
end;

procedure TDriverProperties.Set_Modelo(Value: ModeloPrn);
begin
  DefaultInterface.Set_Modelo(Value);
end;

function TDriverProperties.Get_TotalDocFiscales: Double;
begin
    Result := DefaultInterface.TotalDocFiscales;
end;

function TDriverProperties.Get_CbteEsFiscal: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.CbteEsFiscal;
end;

function TDriverProperties.Get_ErroresEnExcepciones: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.ErroresEnExcepciones;
end;

procedure TDriverProperties.Set_ErroresEnExcepciones(Value: OLE_CANCELBOOL);
begin
  DefaultInterface.Set_ErroresEnExcepciones(Value);
end;

function TDriverProperties.Get_Copias: Integer;
begin
    Result := DefaultInterface.Copias;
end;

procedure TDriverProperties.Set_Copias(Value: Integer);
begin
  DefaultInterface.Set_Copias(Value);
end;

function TDriverProperties.Get_Depurar: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.Depurar;
end;

procedure TDriverProperties.Set_Depurar(Value: OLE_CANCELBOOL);
begin
  DefaultInterface.Set_Depurar(Value);
end;

function TDriverProperties.Get_CierreZTotales: CierreZTotales;
begin
    Result := DefaultInterface.CierreZTotales;
end;

function TDriverProperties.Get_PrecioBase: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.PrecioBase;
end;

procedure TDriverProperties.Set_PrecioBase(Value: OLE_CANCELBOOL);
begin
  DefaultInterface.Set_PrecioBase(Value);
end;

{$ENDIF}

class function CoObtenerDatosDeInicializacionRespuesta.Create: IObtenerDatosDeInicializacionRespuesta;
begin
  Result := CreateComObject(CLASS_ObtenerDatosDeInicializacionRespuesta) as IObtenerDatosDeInicializacionRespuesta;
end;

class function CoObtenerDatosDeInicializacionRespuesta.CreateRemote(const MachineName: string): IObtenerDatosDeInicializacionRespuesta;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ObtenerDatosDeInicializacionRespuesta) as IObtenerDatosDeInicializacionRespuesta;
end;

procedure TObtenerDatosDeInicializacionRespuesta.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{EF88ACD1-CD97-418F-A01B-B4657E28C6B2}';
    IntfIID:   '{44C8E088-C222-4FC1-94ED-9395F5FE32C2}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TObtenerDatosDeInicializacionRespuesta.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IObtenerDatosDeInicializacionRespuesta;
  end;
end;

procedure TObtenerDatosDeInicializacionRespuesta.ConnectTo(svrIntf: IObtenerDatosDeInicializacionRespuesta);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TObtenerDatosDeInicializacionRespuesta.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TObtenerDatosDeInicializacionRespuesta.GetDefaultInterface: IObtenerDatosDeInicializacionRespuesta;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TObtenerDatosDeInicializacionRespuesta.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TObtenerDatosDeInicializacionRespuestaProperties.Create(Self);
{$ENDIF}
end;

destructor TObtenerDatosDeInicializacionRespuesta.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TObtenerDatosDeInicializacionRespuesta.GetServerProperties: TObtenerDatosDeInicializacionRespuestaProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TObtenerDatosDeInicializacionRespuesta.Get_NroCUIT: WideString;
begin
    Result := DefaultInterface.NroCUIT;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_RazonSocial: WideString;
begin
    Result := DefaultInterface.RazonSocial;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_NroSerie: WideString;
begin
    Result := DefaultInterface.NroSerie;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_FechaInicializacion: WideString;
begin
    Result := DefaultInterface.FechaInicializacion;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_NroPOS: WideString;
begin
    Result := DefaultInterface.NroPOS;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_FechaIniActividades: WideString;
begin
    Result := DefaultInterface.FechaIniActividades;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_CodIngBrutos: WideString;
begin
    Result := DefaultInterface.CodIngBrutos;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_RespIVA: WideString;
begin
    Result := DefaultInterface.RespIVA;
end;

function TObtenerDatosDeInicializacionRespuesta.Get_Resultado: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.Resultado;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TObtenerDatosDeInicializacionRespuestaProperties.Create(AServer: TObtenerDatosDeInicializacionRespuesta);
begin
  inherited Create;
  FServer := AServer;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.GetDefaultInterface: IObtenerDatosDeInicializacionRespuesta;
begin
  Result := FServer.DefaultInterface;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_NroCUIT: WideString;
begin
    Result := DefaultInterface.NroCUIT;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_RazonSocial: WideString;
begin
    Result := DefaultInterface.RazonSocial;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_NroSerie: WideString;
begin
    Result := DefaultInterface.NroSerie;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_FechaInicializacion: WideString;
begin
    Result := DefaultInterface.FechaInicializacion;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_NroPOS: WideString;
begin
    Result := DefaultInterface.NroPOS;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_FechaIniActividades: WideString;
begin
    Result := DefaultInterface.FechaIniActividades;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_CodIngBrutos: WideString;
begin
    Result := DefaultInterface.CodIngBrutos;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_RespIVA: WideString;
begin
    Result := DefaultInterface.RespIVA;
end;

function TObtenerDatosDeInicializacionRespuestaProperties.Get_Resultado: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.Resultado;
end;

{$ENDIF}

class function CoSubtotalRespuesta.Create: ISubtotalRespuesta;
begin
  Result := CreateComObject(CLASS_SubtotalRespuesta) as ISubtotalRespuesta;
end;

class function CoSubtotalRespuesta.CreateRemote(const MachineName: string): ISubtotalRespuesta;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SubtotalRespuesta) as ISubtotalRespuesta;
end;

procedure TSubtotalRespuesta.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{27D2653D-A3D2-4037-A5AD-EF73A64A0C69}';
    IntfIID:   '{09BDCB7C-4945-4231-AB0C-628CF69E8561}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TSubtotalRespuesta.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ISubtotalRespuesta;
  end;
end;

procedure TSubtotalRespuesta.ConnectTo(svrIntf: ISubtotalRespuesta);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TSubtotalRespuesta.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TSubtotalRespuesta.GetDefaultInterface: ISubtotalRespuesta;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TSubtotalRespuesta.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TSubtotalRespuestaProperties.Create(Self);
{$ENDIF}
end;

destructor TSubtotalRespuesta.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TSubtotalRespuesta.GetServerProperties: TSubtotalRespuestaProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TSubtotalRespuesta.Get_CantidadItemsVendidos: Double;
begin
    Result := DefaultInterface.CantidadItemsVendidos;
end;

function TSubtotalRespuesta.Get_MontoVentas: Double;
begin
    Result := DefaultInterface.MontoVentas;
end;

function TSubtotalRespuesta.Get_MontoIVA: Double;
begin
    Result := DefaultInterface.MontoIVA;
end;

function TSubtotalRespuesta.Get_MontoPagado: Double;
begin
    Result := DefaultInterface.MontoPagado;
end;

function TSubtotalRespuesta.Get_MontoIVANoInscripto: Double;
begin
    Result := DefaultInterface.MontoIVANoInscripto;
end;

function TSubtotalRespuesta.Get_MontoImpuestosInternos: Double;
begin
    Result := DefaultInterface.MontoImpuestosInternos;
end;

function TSubtotalRespuesta.Get_MontoNeto: Double;
begin
    Result := DefaultInterface.MontoNeto;
end;

function TSubtotalRespuesta.Get_Resultado: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.Resultado;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TSubtotalRespuestaProperties.Create(AServer: TSubtotalRespuesta);
begin
  inherited Create;
  FServer := AServer;
end;

function TSubtotalRespuestaProperties.GetDefaultInterface: ISubtotalRespuesta;
begin
  Result := FServer.DefaultInterface;
end;

function TSubtotalRespuestaProperties.Get_CantidadItemsVendidos: Double;
begin
    Result := DefaultInterface.CantidadItemsVendidos;
end;

function TSubtotalRespuestaProperties.Get_MontoVentas: Double;
begin
    Result := DefaultInterface.MontoVentas;
end;

function TSubtotalRespuestaProperties.Get_MontoIVA: Double;
begin
    Result := DefaultInterface.MontoIVA;
end;

function TSubtotalRespuestaProperties.Get_MontoPagado: Double;
begin
    Result := DefaultInterface.MontoPagado;
end;

function TSubtotalRespuestaProperties.Get_MontoIVANoInscripto: Double;
begin
    Result := DefaultInterface.MontoIVANoInscripto;
end;

function TSubtotalRespuestaProperties.Get_MontoImpuestosInternos: Double;
begin
    Result := DefaultInterface.MontoImpuestosInternos;
end;

function TSubtotalRespuestaProperties.Get_MontoNeto: Double;
begin
    Result := DefaultInterface.MontoNeto;
end;

function TSubtotalRespuestaProperties.Get_Resultado: OLE_CANCELBOOL;
begin
    Result := DefaultInterface.Resultado;
end;

{$ENDIF}

class function CoCierreZTotales.Create: ICierreZTotales;
begin
  Result := CreateComObject(CLASS_CierreZTotales) as ICierreZTotales;
end;

class function CoCierreZTotales.CreateRemote(const MachineName: string): ICierreZTotales;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CierreZTotales) as ICierreZTotales;
end;

procedure TCierreZTotales.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{F0C532B6-9FDC-4A80-BEC1-C9A064F5400D}';
    IntfIID:   '{A7973DAB-A411-454D-927E-517037721A21}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCierreZTotales.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ICierreZTotales;
  end;
end;

procedure TCierreZTotales.ConnectTo(svrIntf: ICierreZTotales);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TCierreZTotales.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TCierreZTotales.GetDefaultInterface: ICierreZTotales;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TCierreZTotales.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TCierreZTotalesProperties.Create(Self);
{$ENDIF}
end;

destructor TCierreZTotales.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TCierreZTotales.GetServerProperties: TCierreZTotalesProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TCierreZTotales.Get_FNDTotalVentas: Double;
begin
    Result := DefaultInterface.FNDTotalVentas;
end;

function TCierreZTotales.Get_FNDTotalIVA: Double;
begin
    Result := DefaultInterface.FNDTotalIVA;
end;

function TCierreZTotales.Get_FNDTotalImpuestosInternos: Double;
begin
    Result := DefaultInterface.FNDTotalImpuestosInternos;
end;

function TCierreZTotales.Get_FNDTotalOtrosTributos: Double;
begin
    Result := DefaultInterface.FNDTotalOtrosTributos;
end;

function TCierreZTotales.Get_NCTotalVentas: Double;
begin
    Result := DefaultInterface.NCTotalVentas;
end;

function TCierreZTotales.Get_NCTotalIVA: Double;
begin
    Result := DefaultInterface.NCTotalIVA;
end;

function TCierreZTotales.Get_NCTotalImpuestosInternos: Double;
begin
    Result := DefaultInterface.NCTotalImpuestosInternos;
end;

function TCierreZTotales.Get_NCTotalOtrosTributos: Double;
begin
    Result := DefaultInterface.NCTotalOtrosTributos;
end;

function TCierreZTotales.Get_NroCierre: Integer;
begin
    Result := DefaultInterface.NroCierre;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TCierreZTotalesProperties.Create(AServer: TCierreZTotales);
begin
  inherited Create;
  FServer := AServer;
end;

function TCierreZTotalesProperties.GetDefaultInterface: ICierreZTotales;
begin
  Result := FServer.DefaultInterface;
end;

function TCierreZTotalesProperties.Get_FNDTotalVentas: Double;
begin
    Result := DefaultInterface.FNDTotalVentas;
end;

function TCierreZTotalesProperties.Get_FNDTotalIVA: Double;
begin
    Result := DefaultInterface.FNDTotalIVA;
end;

function TCierreZTotalesProperties.Get_FNDTotalImpuestosInternos: Double;
begin
    Result := DefaultInterface.FNDTotalImpuestosInternos;
end;

function TCierreZTotalesProperties.Get_FNDTotalOtrosTributos: Double;
begin
    Result := DefaultInterface.FNDTotalOtrosTributos;
end;

function TCierreZTotalesProperties.Get_NCTotalVentas: Double;
begin
    Result := DefaultInterface.NCTotalVentas;
end;

function TCierreZTotalesProperties.Get_NCTotalIVA: Double;
begin
    Result := DefaultInterface.NCTotalIVA;
end;

function TCierreZTotalesProperties.Get_NCTotalImpuestosInternos: Double;
begin
    Result := DefaultInterface.NCTotalImpuestosInternos;
end;

function TCierreZTotalesProperties.Get_NCTotalOtrosTributos: Double;
begin
    Result := DefaultInterface.NCTotalOtrosTributos;
end;

function TCierreZTotalesProperties.Get_NroCierre: Integer;
begin
    Result := DefaultInterface.NroCierre;
end;

{$ENDIF}

class function CoConsultarCapacidadZetasRespuesta.Create: IConsultarCapacidadZetasRespuesta;
begin
  Result := CreateComObject(CLASS_ConsultarCapacidadZetasRespuesta) as IConsultarCapacidadZetasRespuesta;
end;

class function CoConsultarCapacidadZetasRespuesta.CreateRemote(const MachineName: string): IConsultarCapacidadZetasRespuesta;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ConsultarCapacidadZetasRespuesta) as IConsultarCapacidadZetasRespuesta;
end;

procedure TConsultarCapacidadZetasRespuesta.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{D9C06EB3-5688-46D9-839F-8F265C41272F}';
    IntfIID:   '{1D33F62C-0EA9-44D4-8971-2333F441D7EE}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TConsultarCapacidadZetasRespuesta.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IConsultarCapacidadZetasRespuesta;
  end;
end;

procedure TConsultarCapacidadZetasRespuesta.ConnectTo(svrIntf: IConsultarCapacidadZetasRespuesta);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TConsultarCapacidadZetasRespuesta.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TConsultarCapacidadZetasRespuesta.GetDefaultInterface: IConsultarCapacidadZetasRespuesta;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TConsultarCapacidadZetasRespuesta.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TConsultarCapacidadZetasRespuestaProperties.Create(Self);
{$ENDIF}
end;

destructor TConsultarCapacidadZetasRespuesta.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TConsultarCapacidadZetasRespuesta.GetServerProperties: TConsultarCapacidadZetasRespuestaProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TConsultarCapacidadZetasRespuesta.Get_CantidadDeZetasRemanente: Integer;
begin
    Result := DefaultInterface.CantidadDeZetasRemanente;
end;

function TConsultarCapacidadZetasRespuesta.Get_UltimaZeta: Integer;
begin
    Result := DefaultInterface.UltimaZeta;
end;

function TConsultarCapacidadZetasRespuesta.Get_UltimaZetaBajada: Integer;
begin
    Result := DefaultInterface.UltimaZetaBajada;
end;

function TConsultarCapacidadZetasRespuesta.Get_UltimaZetaBorrable: Integer;
begin
    Result := DefaultInterface.UltimaZetaBorrable;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TConsultarCapacidadZetasRespuestaProperties.Create(AServer: TConsultarCapacidadZetasRespuesta);
begin
  inherited Create;
  FServer := AServer;
end;

function TConsultarCapacidadZetasRespuestaProperties.GetDefaultInterface: IConsultarCapacidadZetasRespuesta;
begin
  Result := FServer.DefaultInterface;
end;

function TConsultarCapacidadZetasRespuestaProperties.Get_CantidadDeZetasRemanente: Integer;
begin
    Result := DefaultInterface.CantidadDeZetasRemanente;
end;

function TConsultarCapacidadZetasRespuestaProperties.Get_UltimaZeta: Integer;
begin
    Result := DefaultInterface.UltimaZeta;
end;

function TConsultarCapacidadZetasRespuestaProperties.Get_UltimaZetaBajada: Integer;
begin
    Result := DefaultInterface.UltimaZetaBajada;
end;

function TConsultarCapacidadZetasRespuestaProperties.Get_UltimaZetaBorrable: Integer;
begin
    Result := DefaultInterface.UltimaZetaBorrable;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TDriver, TObtenerDatosDeInicializacionRespuesta, TSubtotalRespuesta, TCierreZTotales, 
    TConsultarCapacidadZetasRespuesta]);
end;

end.
