package ifu.wsh  ;

import com4j.*;

/**
 * Dispatch interface for Driver Object
 */
@IID("{00AA0FC3-6850-4F18-BB90-9FE15E32ACBD}")
public interface IDriver extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Error"
   * </p>
   * @return  Returns a value of type ifu.wsh.ErrorNro
   */

  @DISPID(203) //= 0xcb. The runtime will prefer the VTID if present
  @VTID(7)
    ifu.wsh.ErrorNro error();


  /**
   * <p>
   * Getter method for the COM property "ErrorDesc"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(204) //= 0xcc. The runtime will prefer the VTID if present
  @VTID(8)
  java.lang.String errorDesc();


  /**
   * <p>
   * Getter method for the COM property "Puerto"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(9)
  int puerto();


  /**
   * <p>
   * Setter method for the COM property "Puerto"
   * </p>
   * @param value Mandatory int parameter.
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(10)
  void puerto(
    int value);


  /**
   * <p>
   * Getter method for the COM property "Baudios"
   * </p>
   * @return  Returns a value of type ifu.wsh.Baudio
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(11)
    ifu.wsh.Baudio baudios();


  /**
   * <p>
   * Setter method for the COM property "Baudios"
   * </p>
   * @param value Mandatory ifu.wsh.Baudio parameter.
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(12)
  void baudios(
    ifu.wsh.Baudio value);


  /**
   * <p>
   * Getter method for the COM property "Modelo"
   * </p>
   * @return  Returns a value of type ifu.wsh.ModeloPrn
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(13)
    ifu.wsh.ModeloPrn modelo();


  /**
   * <p>
   * Setter method for the COM property "Modelo"
   * </p>
   * @param value Mandatory ifu.wsh.ModeloPrn parameter.
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(14)
  void modelo(
    ifu.wsh.ModeloPrn value);


  /**
   * @param aTipoDeComprobante Mandatory ifu.wsh.TipoDeComprobante parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(15)
  boolean abrirComprobante(
    ifu.wsh.TipoDeComprobante aTipoDeComprobante);


  /**
   * @param aDescripcion Mandatory java.lang.String parameter.
   * @param aCantidad Mandatory double parameter.
   * @param aPrecio Mandatory double parameter.
   * @param aIVA Mandatory double parameter.
   * @param aImpuestosInternos Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(207) //= 0xcf. The runtime will prefer the VTID if present
  @VTID(16)
  boolean imprimirItem(
    java.lang.String aDescripcion,
    double aCantidad,
    double aPrecio,
    double aIVA,
    double aImpuestosInternos);


  /**
   * @param descripcion Mandatory java.lang.String parameter.
   * @param cantidad Mandatory double parameter.
   * @param precio Mandatory double parameter.
   * @param iva Mandatory double parameter.
   * @param impuestosInternos Mandatory double parameter.
   * @param g2CondicionIVA Mandatory ifu.wsh.CondicionesIVA parameter.
   * @param g2TipoImpuestoInterno Mandatory ifu.wsh.TipoImpuestoInterno parameter.
   * @param g2UnidadReferencia Mandatory int parameter.
   * @param g2CodigoProducto Mandatory java.lang.String parameter.
   * @param g2CodigoInterno Mandatory java.lang.String parameter.
   * @param g2UnidadMedida Mandatory ifu.wsh.UnidadesMedida parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(242) //= 0xf2. The runtime will prefer the VTID if present
  @VTID(17)
  boolean imprimirItem2g(
    java.lang.String descripcion,
    double cantidad,
    double precio,
    double iva,
    double impuestosInternos,
    ifu.wsh.CondicionesIVA g2CondicionIVA,
    ifu.wsh.TipoImpuestoInterno g2TipoImpuestoInterno,
    int g2UnidadReferencia,
    java.lang.String g2CodigoProducto,
    java.lang.String g2CodigoInterno,
    ifu.wsh.UnidadesMedida g2UnidadMedida);


  /**
   * @param aDescripcion Mandatory java.lang.String parameter.
   * @param aMonto Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(208) //= 0xd0. The runtime will prefer the VTID if present
  @VTID(18)
  boolean imprimirDescuentoGeneral(
    java.lang.String aDescripcion,
    double aMonto);


  /**
   * @param aDescripcion Mandatory java.lang.String parameter.
   * @param aMonto Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(209) //= 0xd1. The runtime will prefer the VTID if present
  @VTID(19)
  boolean imprimirPago(
    java.lang.String aDescripcion,
    double aMonto);


  /**
   * @param descripcion Mandatory java.lang.String parameter.
   * @param monto Mandatory double parameter.
   * @param g2DescripcionAdicional Mandatory java.lang.String parameter.
   * @param g2CodigoFormaPago Mandatory ifu.wsh.TiposPago parameter.
   * @param g2Cuotas Mandatory int parameter.
   * @param g2Cupones Mandatory java.lang.String parameter.
   * @param g2Referencia Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(243) //= 0xf3. The runtime will prefer the VTID if present
  @VTID(20)
  boolean imprimirPago2g(
    java.lang.String descripcion,
    double monto,
    java.lang.String g2DescripcionAdicional,
    ifu.wsh.TiposPago g2CodigoFormaPago,
    int g2Cuotas,
    java.lang.String g2Cupones,
    java.lang.String g2Referencia);


  /**
   */

  @DISPID(210) //= 0xd2. The runtime will prefer the VTID if present
  @VTID(21)
  void cerrarComprobante();


  /**
   * @param aNombre Mandatory java.lang.String parameter.
   * @param aTipoDeDocumento Mandatory ifu.wsh.TipoDeDocumento parameter.
   * @param aDocumento Mandatory java.lang.String parameter.
   * @param aResponsIVA Mandatory ifu.wsh.ResponsabilidadIVA parameter.
   * @param aDireccion Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(22)
  boolean datosCliente(
    java.lang.String aNombre,
    ifu.wsh.TipoDeDocumento aTipoDeDocumento,
    java.lang.String aDocumento,
    ifu.wsh.ResponsabilidadIVA aResponsIVA,
    java.lang.String aDireccion);


  /**
   */

  @DISPID(212) //= 0xd4. The runtime will prefer the VTID if present
  @VTID(23)
  void cancelarComprobante();


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(213) //= 0xd5. The runtime will prefer the VTID if present
  @VTID(24)
  boolean inicializar();


  /**
   * <p>
   * Getter method for the COM property "TotalDocFiscales"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(214) //= 0xd6. The runtime will prefer the VTID if present
  @VTID(25)
  double totalDocFiscales();


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(215) //= 0xd7. The runtime will prefer the VTID if present
  @VTID(26)
  boolean cierreX();


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(216) //= 0xd8. The runtime will prefer the VTID if present
  @VTID(27)
  boolean cierreZ();


  /**
   * @param aTexto Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(217) //= 0xd9. The runtime will prefer the VTID if present
  @VTID(28)
  boolean imprimirTextoFiscal(
    java.lang.String aTexto);


  /**
   * @param aDescripcion Mandatory java.lang.String parameter.
   * @param aMonto Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(218) //= 0xda. The runtime will prefer the VTID if present
  @VTID(29)
  boolean informarPercepcionGlobal(
    java.lang.String aDescripcion,
    double aMonto);


  /**
   * @param aDescripcion Mandatory java.lang.String parameter.
   * @param aMonto Mandatory double parameter.
   * @param aAlicuota Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(219) //= 0xdb. The runtime will prefer the VTID if present
  @VTID(30)
  boolean informarPercepcionIVA(
    java.lang.String aDescripcion,
    double aMonto,
    double aAlicuota);


  /**
   * <p>
   * Getter method for the COM property "CbteEsFiscal"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(220) //= 0xdc. The runtime will prefer the VTID if present
  @VTID(31)
  boolean cbteEsFiscal();


  /**
   * @param aDocumento Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(221) //= 0xdd. The runtime will prefer the VTID if present
  @VTID(32)
  boolean documentoDeReferencia(
    java.lang.String aDocumento);


  /**
   * @param aTipoComprobante Mandatory ifu.wsh.TipoDeComprobante parameter.
   * @return  Returns a value of type int
   */

  @DISPID(222) //= 0xde. The runtime will prefer the VTID if present
  @VTID(33)
  int ultimoComprobante(
    ifu.wsh.TipoDeComprobante aTipoComprobante);


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(223) //= 0xdf. The runtime will prefer the VTID if present
  @VTID(34)
  boolean ultimoComprobanteCancelado();


  /**
   * <p>
   * Getter method for the COM property "ErroresEnExcepciones"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(224) //= 0xe0. The runtime will prefer the VTID if present
  @VTID(35)
  boolean erroresEnExcepciones();


  /**
   * <p>
   * Setter method for the COM property "ErroresEnExcepciones"
   * </p>
   * @param value Mandatory boolean parameter.
   */

  @DISPID(224) //= 0xe0. The runtime will prefer the VTID if present
  @VTID(36)
  void erroresEnExcepciones(
    boolean value);


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(225) //= 0xe1. The runtime will prefer the VTID if present
  @VTID(37)
  boolean finalizar();


  /**
   * @param obraSocial Mandatory java.lang.String parameter.
   * @param coseguro1 Mandatory java.lang.String parameter.
   * @param coseguro2 Mandatory java.lang.String parameter.
   * @param coseguro3 Mandatory java.lang.String parameter.
   * @param nroAfiliado Mandatory java.lang.String parameter.
   * @param nombreAfiliado Mandatory java.lang.String parameter.
   * @param fechaVencimientoCarnet Mandatory java.lang.String parameter.
   * @param domicilioVend1 Mandatory java.lang.String parameter.
   * @param domicilioVend2 Mandatory java.lang.String parameter.
   * @param nombreEstablecimiento Mandatory java.lang.String parameter.
   * @param nroInterno Mandatory java.lang.String parameter.
   * @param nota1 Mandatory java.lang.String parameter.
   * @param nota2 Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(226) //= 0xe2. The runtime will prefer the VTID if present
  @VTID(38)
  boolean dnfhFarmacias(
    java.lang.String obraSocial,
    java.lang.String coseguro1,
    java.lang.String coseguro2,
    java.lang.String coseguro3,
    java.lang.String nroAfiliado,
    java.lang.String nombreAfiliado,
    java.lang.String fechaVencimientoCarnet,
    java.lang.String domicilioVend1,
    java.lang.String domicilioVend2,
    java.lang.String nombreEstablecimiento,
    java.lang.String nroInterno,
    java.lang.String nota1,
    java.lang.String nota2);


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(227) //= 0xe3. The runtime will prefer the VTID if present
  @VTID(39)
  boolean cortarPapel();


  /**
   * @param texto Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(228) //= 0xe4. The runtime will prefer the VTID if present
  @VTID(40)
  boolean imprimirTextoNoFiscal(
    java.lang.String texto);


  /**
   * @param descripcion Mandatory java.lang.String parameter.
   * @param monto Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(229) //= 0xe5. The runtime will prefer the VTID if present
  @VTID(41)
  boolean imprimirDescuentoUltimoItem(
    java.lang.String descripcion,
    double monto);


  /**
   * @param fechaInicial Mandatory java.lang.String parameter.
   * @param fechaFinal Mandatory java.lang.String parameter.
   * @param detallado Mandatory boolean parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(230) //= 0xe6. The runtime will prefer the VTID if present
  @VTID(42)
  boolean reporteZFechas(
    java.lang.String fechaInicial,
    java.lang.String fechaFinal,
    boolean detallado);


  /**
   * @param nroInicio Mandatory int parameter.
   * @param nroFin Mandatory int parameter.
   * @param detallado Mandatory boolean parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(231) //= 0xe7. The runtime will prefer the VTID if present
  @VTID(43)
  boolean reporteZNumeros(
    int nroInicio,
    int nroFin,
    boolean detallado);


  /**
   * @param linea Mandatory int parameter.
   * @param texto Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(232) //= 0xe8. The runtime will prefer the VTID if present
  @VTID(44)
  boolean especificarEncabezado(
    int linea,
    java.lang.String texto);


  /**
   * @param linea Mandatory int parameter.
   * @param texto Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(233) //= 0xe9. The runtime will prefer the VTID if present
  @VTID(45)
  boolean especificarPie(
    int linea,
    java.lang.String texto);


  /**
   * @param numero Mandatory Holder<Integer> parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(234) //= 0xea. The runtime will prefer the VTID if present
  @VTID(46)
  boolean cerrarComprobanteNumero(
    Holder<Integer> numero);


  /**
   * <p>
   * Getter method for the COM property "Copias"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(47)
  int copias();


  /**
   * <p>
   * Setter method for the COM property "Copias"
   * </p>
   * @param value Mandatory int parameter.
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(48)
  void copias(
    int value);


  /**
   * <p>
   * Getter method for the COM property "Depurar"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(236) //= 0xec. The runtime will prefer the VTID if present
  @VTID(49)
  boolean depurar();


  /**
   * <p>
   * Setter method for the COM property "Depurar"
   * </p>
   * @param value Mandatory boolean parameter.
   */

  @DISPID(236) //= 0xec. The runtime will prefer the VTID if present
  @VTID(50)
  void depurar(
    boolean value);


  /**
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(237) //= 0xed. The runtime will prefer the VTID if present
  @VTID(51)
  java.lang.String obtenerFechaHora();


  /**
   */

  @DISPID(238) //= 0xee. The runtime will prefer the VTID if present
  @VTID(52)
  void abrirCajon();


  /**
   * @return  Returns a value of type ifu.wsh.IObtenerDatosDeInicializacionRespuesta
   */

  @DISPID(239) //= 0xef. The runtime will prefer the VTID if present
  @VTID(53)
    ifu.wsh.IObtenerDatosDeInicializacionRespuesta obtenerDatosDeInicializacion();


  /**
   * @return  Returns a value of type ifu.wsh.ISubtotalRespuesta
   */

  @DISPID(240) //= 0xf0. The runtime will prefer the VTID if present
  @VTID(54)
    ifu.wsh.ISubtotalRespuesta subtotal();


  /**
   * @param codigo Mandatory ifu.wsh.TiposTributos parameter.
   * @param descripcion Mandatory java.lang.String parameter.
   * @param baseImponible Mandatory double parameter.
   * @param importe Mandatory double parameter.
   * @param alicuota Mandatory double parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(241) //= 0xf1. The runtime will prefer the VTID if present
  @VTID(55)
  boolean imprimirOtrosTributos(
    ifu.wsh.TiposTributos codigo,
    java.lang.String descripcion,
    double baseImponible,
    double importe,
    double alicuota);


  /**
   * @param licencia Mandatory java.lang.String parameter.
   */

  @DISPID(244) //= 0xf4. The runtime will prefer the VTID if present
  @VTID(56)
  void cargarLicencia(
    java.lang.String licencia);


  /**
   * @param direccionIP Mandatory java.lang.String parameter.
   * @param puerto Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(245) //= 0xf5. The runtime will prefer the VTID if present
  @VTID(57)
  boolean conectar(
    java.lang.String direccionIP,
    int puerto);


  /**
   * @param tipoComprobante Mandatory ifu.wsh.TipoDeComprobante parameter.
   * @param documento Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(246) //= 0xf6. The runtime will prefer the VTID if present
  @VTID(58)
  boolean documentoDeReferencia2g(
    ifu.wsh.TipoDeComprobante tipoComprobante,
    java.lang.String documento);


  /**
   * <p>
   * Getter method for the COM property "CierreZTotales"
   * </p>
   * @return  Returns a value of type ifu.wsh.ICierreZTotales
   */

  @DISPID(247) //= 0xf7. The runtime will prefer the VTID if present
  @VTID(59)
    ifu.wsh.ICierreZTotales cierreZTotales();


  /**
   * @param fechaHora Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(248) //= 0xf8. The runtime will prefer the VTID if present
  @VTID(60)
  boolean especificarFechaHora(
    java.lang.String fechaHora);


  /**
   * <p>
   * Getter method for the COM property "PrecioBase"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(249) //= 0xf9. The runtime will prefer the VTID if present
  @VTID(61)
  boolean precioBase();


  /**
   * <p>
   * Setter method for the COM property "PrecioBase"
   * </p>
   * @param value Mandatory boolean parameter.
   */

  @DISPID(249) //= 0xf9. The runtime will prefer the VTID if present
  @VTID(62)
  void precioBase(
    boolean value);


  /**
   * @param razonSocial Mandatory java.lang.String parameter.
   * @param cuit Mandatory double parameter.
   * @param domicilio Mandatory java.lang.String parameter.
   * @param nombreChofer Mandatory java.lang.String parameter.
   * @param tipoDocumento Mandatory ifu.wsh.TipoDeDocumento parameter.
   * @param numeroDocumento Mandatory java.lang.String parameter.
   * @param dominio1 Mandatory java.lang.String parameter.
   * @param dominio2 Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(250) //= 0xfa. The runtime will prefer the VTID if present
  @VTID(63)
  boolean cargarTransportista(
    java.lang.String razonSocial,
    double cuit,
    java.lang.String domicilio,
    java.lang.String nombreChofer,
    ifu.wsh.TipoDeDocumento tipoDocumento,
    java.lang.String numeroDocumento,
    java.lang.String dominio1,
    java.lang.String dominio2);


  // Properties:
}
