package ifu.wsh  ;

import com4j.*;

/**
 */
public enum CondicionesIVA implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  NoGravado(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  Exento(2),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  Gravado(7),
  ;

  private final int value;
  CondicionesIVA(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
