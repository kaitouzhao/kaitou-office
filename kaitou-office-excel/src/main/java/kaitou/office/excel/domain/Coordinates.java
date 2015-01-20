package kaitou.office.excel.domain;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static kaitou.office.excel.common.SysCode.CellType;

/**
 * 数据项坐标.
 * User: 赵立伟
 * Date: 2015/1/10
 * Time: 22:55
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Coordinates {

    int[] x() default {};

    int[] y() default {};

    CellType type() default CellType.STRING;
}
