package ifu.wsh  ;

import com4j.*;

/**
 */
public enum PuertoCOM implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  pcCOM1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  pcCOM2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  pcCOM3(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  pcCOM4(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  pcCOM5(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  pcCOM6(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  pcCOM7(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  pcCOM8(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  pcCOM9(9),
  ;

  private final int value;
  PuertoCOM(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
