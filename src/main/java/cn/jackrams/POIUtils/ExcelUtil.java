package cn.jackrams.POIUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.opc.internal.ZipHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.jackrams.POIUtils.emuns.TypeEnum;
import cn.jackrams.POIUtils.emuns.ViewType;


public class ExcelUtil<T> {
	private static final String XML_ENCODING = "UTF-8";
	private List<String> fieldNames=new ArrayList<String>();
    private	Map<String,Map<String,Object>> fieldProperties=new LinkedHashMap<String, Map<String,Object>>();
    private AnnotationProcessor processor=new AnnotationProcessor();
	//private Map<String,> 
    private static final Log log=LogFactory.getLog(ExcelUtil.class);
  public ExcelUtil(T t) {
	  if(log.isInfoEnabled()){
		  log.info("Start Annotation Processor");
	  }
   Class<? extends Object> clazz = t.getClass();
	
   try {
	processor.fieldAnnotationProcessor(clazz, fieldProperties, fieldNames);
} catch (Exception e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
}
   if(log.isInfoEnabled()){
		  log.info("End Annotation Processor");
	  }
   System.out.println(fieldNames);
	System.out.println("Properties"+fieldProperties);
  }
  
  public void exportExcle(Collection<T> data,OutputStream stream,
		  File templateFile,Map<String, XSSFCellStyle> cellStyles,String sheetName){
	  if(log.isInfoEnabled()){
			 log.info("Start Excel Export");
		 }
	  XSSFWorkbook wb = new XSSFWorkbook();
	  if(cellStyles==null || cellStyles.isEmpty()){
    	   cellStyles=createStyles(wb);
       }
      try{
      XSSFSheet sheet = wb.createSheet(sheetName);

    
      //name of the zip entry holding sheet data, e.g. /xl/worksheets/sheet1.xml
      String sheetRef = sheet.getPackagePart().getPartName().getName();

      //save the template
      FileOutputStream os = new FileOutputStream(templateFile);
      wb.write(os);
      os.close();

      //Step 2. Generate XML file.
      File tmp = File.createTempFile("sheet", ".xml");
      Writer fw = new OutputStreamWriter(new FileOutputStream(tmp), XML_ENCODING);
      generateData(fw, cellStyles,data);
      fw.close();

      //Step 3. Substitute the template entry with the generated data
    //  FileOutputStream out = new FileOutputStream("big-grid.xlsx");
      substitute(templateFile, tmp, sheetRef.substring(1), stream);
      stream.close();
      } catch(Exception e){
    	  log.error("Stop Excel Export，Something Error",e);
    	  
    	  }
     
  }
  
  /**
   * 
   * @param out
   * @param styles
   * @param data
   * @throws Exception
   */
  private  void generateData(Writer out, Map<String, XSSFCellStyle> styles,Collection<T> data) throws Exception {

      Calendar calendar = Calendar.getInstance();

      SpreadsheetWriter sw = new SpreadsheetWriter(out);
      sw.beginSheet();

      //insert header row
      sw.insertRow(0);
      int styleIndex = styles.get("header").getIndex();
       for(int i=0;i<fieldNames.size();i++){
    	   String fieldName = fieldNames.get(i);
    	   Map<String, Object> fieldPropertie = fieldProperties.get(fieldName);
    	   String title = (String) fieldPropertie.get("title");
    	   sw.createCell(i, title, styleIndex);
       }

      sw.endRow();

      int rownum=0;
      Object[] args=new Object[]{};
      for (T t : data) {
    	  rownum++;
		sw.insertRow(rownum);
		for (int i = 0; i <fieldNames.size(); i++) {
			String fieldName = fieldNames.get(i);
			Map<String, Object> properties = fieldProperties.get(fieldName);
			Method method = (Method) properties.get("method");
			Object value = method.invoke(t, args);
			TypeEnum type=(TypeEnum) properties.get("type");
		    ViewType viewType=	(ViewType) properties.get("viewType");
		      if(value!=null){
			if(TypeEnum.Boolean==type){
				Boolean boolValue=(Boolean) value;
				if(boolValue)
				{
					 sw.createCell(i,"是" );
				}else{
					sw.createCell(i, "否");
				}			
				
			}else if(viewType==ViewType.Money){
				try{
				double doubleValue = Double.parseDouble(value.toString());
				
		
				sw.createCell(i, doubleValue, styles.get("currency").getIndex());
				}catch (Exception e) {
                 sw.createCell(i, value.toString());
				}
			}
			else if(TypeEnum.Date==type){
				Date date=(Date) value;
			  calendar.setTime(date);
			   sw.createCell(i, calendar, styles.get("date").getIndex());
				
			}else if(TypeEnum.Double==type||TypeEnum.Integer==type||TypeEnum.Float==type){
				  if(viewType==ViewType.Percent){
					   
					  sw.createCell(i, Double.parseDouble(value.toString()), styles.get("percent").getIndex());
				  }else{
					  sw.createCell(i,Double.parseDouble( value.toString()));
				  }
			}else if(type==TypeEnum.ByteArray || viewType==ViewType.Image){
				sw.createCell(i, " ");
			}else{
				sw.createCell(i, value.toString());
			}
			
		      }
		      //当为空时
		      else{
		    	  sw.createCell(i, " ");
		      }
            			
		}
		
		sw.endRow();
		
	}
      sw.endSheet();
  }
  
  
  private  Map<String, XSSFCellStyle> createStyles(XSSFWorkbook wb){
      Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
      XSSFDataFormat fmt = wb.createDataFormat();

      XSSFCellStyle style1 = wb.createCellStyle();
      style1.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
      style1.setDataFormat(fmt.getFormat("0.0%"));
      styles.put("percent", style1);

      XSSFCellStyle style2 = wb.createCellStyle();
      style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
      style2.setDataFormat(fmt.getFormat("0.0X"));
      styles.put("coeff", style2);

      XSSFCellStyle style3 = wb.createCellStyle();
      style3.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
      style3.setDataFormat(fmt.getFormat("￥#,##0.00"));
      styles.put("currency", style3);

      XSSFCellStyle style4 = wb.createCellStyle();
      style4.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
      style4.setDataFormat(fmt.getFormat("yyyy-mm-dd"));
      styles.put("date", style4);

      XSSFCellStyle style5 = wb.createCellStyle();
      XSSFFont headerFont = wb.createFont();
      headerFont.setBold(true);
      style5.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
      style5.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
      style5.setFont(headerFont);
      styles.put("header", style5);

      return styles;
  }

  
  private  void substitute(File zipfile, File tmpfile, String entry, OutputStream out) throws IOException {
      ZipFile zip = ZipHelper.openZipFile(zipfile);
      try {
          ZipOutputStream zos = new ZipOutputStream(out);
  
          Enumeration<? extends ZipEntry> en = zip.entries();
          while (en.hasMoreElements()) {
              ZipEntry ze = en.nextElement();
              if(!ze.getName().equals(entry)){
                  zos.putNextEntry(new ZipEntry(ze.getName()));
                  InputStream is = zip.getInputStream(ze);
                  copyStream(is, zos);
                  is.close();
              }
          }
          zos.putNextEntry(new ZipEntry(entry));
          InputStream is = new FileInputStream(tmpfile);
          copyStream(is, zos);
          is.close();
  
          zos.close();
      } finally {
          zip.close();
      }
  }

  private  void copyStream(InputStream in, OutputStream out) throws IOException {
      byte[] chunk = new byte[1024];
      int count;
      while ((count = in.read(chunk)) >=0 ) {
        out.write(chunk,0,count);
      }
  }

  /**
   * 
   * 
   *
   */
  public static class SpreadsheetWriter {
      private final Writer _out;
      private int _rownum;

      public SpreadsheetWriter(Writer out){
          _out = out;
      }

      public void beginSheet() throws IOException {
          _out.write("<?xml version=\"1.0\" encoding=\""+XML_ENCODING+"\"?>" +
                  "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" );
          _out.write("<sheetData>\n");
      }

      public void endSheet() throws IOException {
          _out.write("</sheetData>");
          _out.write("</worksheet>");
      }

      /**
       * Insert a new row
       *
       * @param rownum 0-based row number
       */
      public void insertRow(int rownum) throws IOException {
          _out.write("<row r=\""+(rownum+1)+"\">\n");
          this._rownum = rownum;
      }

      /**
       * Insert row end marker
       */
      public void endRow() throws IOException {
          _out.write("</row>\n");
      }

      public void createCell(int columnIndex, String value, int styleIndex) throws IOException {
          String ref = new CellReference(_rownum, columnIndex).formatAsString();
          _out.write("<c r=\""+ref+"\" t=\"inlineStr\"");
          if(styleIndex != -1) _out.write(" s=\""+styleIndex+"\"");
          _out.write(">");
          _out.write("<is><t>"+value+"</t></is>");
          _out.write("</c>");
      }

      public void createCell(int columnIndex, String value) throws IOException {
          createCell(columnIndex, value, -1);
      }

      public void createCell(int columnIndex, double value, int styleIndex) throws IOException {
          String ref = new CellReference(_rownum, columnIndex).formatAsString();
          _out.write("<c r=\""+ref+"\" t=\"n\"");
          if(styleIndex != -1) _out.write(" s=\""+styleIndex+"\"");
          _out.write(">");
          _out.write("<v>"+value+"</v>");
          _out.write("</c>");
      }

      public void createCell(int columnIndex, double value) throws IOException {
          createCell(columnIndex, value, -1);
      }

      public void createCell(int columnIndex, Calendar value, int styleIndex) throws IOException {
          createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
      }
  }
  
}
