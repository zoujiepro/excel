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
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelEntity {

    String classPathSource() default "";

    int sheetAt() default 0;

    int skip() default 0;

    int limitRow() default -1;

}
