package cn.jackrams.POIUtils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import cn.jackrams.POIUtils.emuns.TypeEnum;
import cn.jackrams.POIUtils.emuns.ViewType;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface PoiField {
  String title() default "";
  /**
   * This is the method Perfix of Field Getter
   * We Know The Method Perfix String Of The Boolean Tyoe is "is"
   * @return
   */
 String getterMethodPrefix() default"get";
  String value() default"";
  String formate() default "";
  TypeEnum type() default TypeEnum.String;
  ViewType viewType() default ViewType.String;
  
}
