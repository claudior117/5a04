package ifu.wsh  ;

import com4j.*;

@IID("{09BDCB7C-4945-4231-AB0C-628CF69E8561}")
public interface ISubtotalRespuesta extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "CantidadItemsVendidos"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(7)
  double cantidadItemsVendidos();


  /**
   * <p>
   * Getter method for the COM property "MontoVentas"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(8)
  double montoVentas();


  /**
   * <p>
   * Getter method for the COM property "MontoIVA"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(203) //= 0xcb. The runtime will prefer the VTID if present
  @VTID(9)
  double montoIVA();


  /**
   * <p>
   * Getter method for the COM property "MontoPagado"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(204) //= 0xcc. The runtime will prefer the VTID if present
  @VTID(10)
  double montoPagado();


  /**
   * <p>
   * Getter method for the COM property "MontoIVANoInscripto"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(11)
  double montoIVANoInscripto();


  /**
   * <p>
   * Getter method for the COM property "MontoImpuestosInternos"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(12)
  double montoImpuestosInternos();


  /**
   * <p>
   * Getter method for the COM property "Resultado"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(207) //= 0xcf. The runtime will prefer the VTID if present
  @VTID(13)
  boolean resultado();


  // Properties:
}
