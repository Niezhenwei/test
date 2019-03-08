package com.chinamobile.zj.nb.information.utils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.Format;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import ognl.Ognl;
import ognl.OgnlException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;

public class POIUtil
{
	
	/**数据库里向Excle表格导入*/
	public static Workbook poiDataExport(Map<String,String> map , List<?> list)
			throws OgnlException, ParseException {
		//创建一个Excle表格
		 Workbook wk = new HSSFWorkbook();
		 Sheet sheet = wk.createSheet();
		 //设置表头
         Row row = sheet.createRow(0);
         Set<String> keys = map.keySet();
		 int num = 0;
		 for (String key : keys)
		 {
			 Cell cell = row.createCell(num++);
			 cell.setCellValue(key);
		 }
		 for(int i=0;i< list.size();i++){
		 	num=0;
		    row = sheet.createRow(i+1);
		    for(String key0:map.keySet()){
		 		Cell cell = row.createCell(num++);
		 		String data = Ognl.getValue(map.get(key0), list.get(i))+"";
		 		if(data.contains("CST")){
					SimpleDateFormat sdf1 = new SimpleDateFormat ("EEE MMM dd HH:mm:ss Z yyyy", Locale.UK);
					Date date = sdf1.parse(data);
					SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
					String sDate=sdf.format(date);
					data = sDate;
				}
				cell.setCellValue(data);
			 }
		}
		return wk;
	}
	
	/**Excle表格向数据库中导入
	 * 需要一个输入流，和一个导入的对象类型
	 * @throws OgnlException 
	 * @throws IOException 
	 * @throws IllegalAccessException 
	 * @throws InstantiationException 
	 * @throws ParseException */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static List poiDataImport(InputStream is , Class<?> clazz)
            throws Exception {
        Workbook wk = WorkbookFactory.create(is);
		Sheet sheet = wk.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();
		List lists = new ArrayList();
		Field[] fields = clazz.getDeclaredFields();
		for(int i=1;i<=rowCount;i++){
			Object obj = clazz.newInstance();
			Row row = sheet.getRow(i);
			for(int j=0;j<fields.length;j++){
				String string = null;
				if(row != null && row.getCell(j)!=null){
                    Cell cell = row.getCell(j);
                    string = getXCellVal(cell);
                }
                if(string !=null && string.length() > 0) {
                    Object value = getValue(fields[j].getType(), string);
                    if(value!=null){
                        fields[j].setAccessible(true);
                        fields[j].set(obj,value);
                    }
                }
			}
            if(isObject(obj,clazz)){
                lists.add(obj);
            }
		}
		return lists;
	}
	
	/**
	 * 导入的类型处理
	 * */
	private static <T> T getValue(Class<?> clazz,String str) throws ParseException
	{   
		if(clazz==Integer.class){
			return  (T)  Integer.valueOf(str);
		}else if(clazz==Long.class){
			return  (T) Long.valueOf(str);
		}
		else if(clazz==Date.class){
			SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		     return (T) sdf.parse(str);
		} else if(clazz == Double.class) {
            return (T) Double.valueOf(str);
        } else{
			return (T) str;

		}
	}

	/**
     * 读取表格数据的类型进行格式修
     * */
	private static String getXCellVal(Cell cell) {
	    String val = null;
        Format fmt = new SimpleDateFormat();
        Format df = new DecimalFormat();
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					val = fmt.format(cell.getDateCellValue()); //日期型
				} else {
                    String format = df.format(cell.getNumericCellValue());//数字型
                    val = format.replaceAll(",","");
                }
				break;
			case Cell.CELL_TYPE_STRING: //文本类型
				val = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN: //布尔型
				val = String.valueOf(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_BLANK: //空白
				val = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_ERROR: //错误
				val = "错误";
				break;
			case Cell.CELL_TYPE_FORMULA: //公式
				try {
					val = String.valueOf(cell.getStringCellValue());
				} catch (IllegalStateException e) {
					val = String.valueOf(cell.getNumericCellValue());
				}
				break;
			default:
				val = cell.getRichStringCellValue() == null ? null : cell.getRichStringCellValue().toString();
		}
		return val;
	}

	/**
     * 判断是否是一个有效的对象
     * */
	private static boolean isObject(Object obj, Class clazz) throws IllegalAccessException {
        Field[] fields = clazz.getDeclaredFields();

        for (Field field : fields) {
            field.setAccessible(true);
            if(field.get(obj) != null
                    && (field.get(obj).toString()).length() > 0){
                return true;
            }
        }
	    return false;
    }

}	
