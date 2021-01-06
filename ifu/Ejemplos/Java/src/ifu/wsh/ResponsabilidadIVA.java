package ifu.wsh  ;

import com4j.*;

/**
 */
public enum ResponsabilidadIVA implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  riResponsableInscripto(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  riMonotributo(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  riExento(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  riConsumidorFinal(4),
  ;

  private final int value;
  ResponsabilidadIVA(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
