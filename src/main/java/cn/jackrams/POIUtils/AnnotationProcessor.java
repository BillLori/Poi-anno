package cn.jackrams.POIUtils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import cn.jackrams.POIUtils.annotations.PoiField;
import cn.jackrams.POIUtils.emuns.TypeEnum;
import cn.jackrams.POIUtils.emuns.ViewType;

public class AnnotationProcessor {
	private static final Log log=LogFactory.getLog(AnnotationProcessor.class);
     public void fieldAnnotationProcessor(Class<? extends Object> clazz,Map<String,Map<String,Object>> fieldProperties,List<String> fieldNames) throws Exception{
    	   Field[] fields = clazz.getDeclaredFields();
    	   for (Field field : fields) {
    		   String fieldName = field.getName();
    		   Map<String, Object> map =new HashMap<String, Object>();
    		   if(field.isAnnotationPresent(PoiField.class)){
    			   fieldNames.add(fieldName);
    			   PoiField poiField = field.getAnnotation(PoiField.class);
    			   TypeEnum type = poiField.type();
    			   String title = poiField.title();
    			   String formate = poiField.formate();
    			   String value = poiField.value();
    			   String methodPrefix = poiField.getterMethodPrefix();
    			    ViewType viewType = poiField.viewType();
    			    String name=methodPrefix+fieldName.substring(0, 1).toUpperCase()
    						  + fieldName.substring(1);
    			    Method method = null;
    			     try {
						 method = clazz.getMethod(name, new Class[]{});
					} catch (NoSuchMethodException e) {
						// TODO Auto-generated catch block
						//throw ;
						try{
							if(methodPrefix.equals("is")){
								methodPrefix="get";
								 name=methodPrefix+fieldName.substring(0, 1).toUpperCase()
			    						  + fieldName.substring(1);
							}else{
								methodPrefix="is";
								 name=methodPrefix+fieldName.substring(0, 1).toUpperCase()
			    						  + fieldName.substring(1);
								
								 method = clazz.getMethod(name, new Class[]{});
							}
							
						}catch(Exception e1){
							log.error("The Perfix of getter of the field "+fieldName+" is no found", e1);
							throw  e1;
						}
						
						//e.printStackTrace();
					} catch (SecurityException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
    			  map.put("method", method);
    	          map.put("type", type);
    	          map.put("formate", formate);
    	          map.put("viewType", viewType);
    	          map.put("title", title);
    	          map.put("value", value);
    	       //   map.put("methodPrefix", methodPrefix);
    	          fieldProperties.put(fieldName, map);

    		   }
    	}
    	   
    	 
     } 
}
