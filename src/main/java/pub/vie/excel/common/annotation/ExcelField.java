package pub.vie.excel.common.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Descrption :
 * @Author: zoujie
 * @Date: 2020-4-15
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    int colIndex();

    String dataFormat() default "yyyy-MM-dd";

}
