package com.ckl.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel读写工具类需要的jar:
 * poi-3.10.1-20140818.jar
 * poi-ooxml-3.10.1-20140818.jar
 * poi-ooxml-schemas-3.10.1-20140818.jar
 * xmlbeans-2.6.0.jar
 * dom4j-1.6.1.jar
 */
public class POIUtil {
	private final static String xls = "xls";
	private final static String xlsx = "xlsx";
	
	/**
	 * 读入excel文件，解析后返回
	 * @param file
	 * @throws IOException 
	 */
	public static List<String[]> readExcel(File file) throws IOException{
		//检查文件
		checkFile(file);
    	//获得Workbook工作薄对象
    	Workbook workbook = getWorkBook(file);
    	//创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
    	List<String[]> list = new ArrayList<String[]>();
    	if(workbook != null){
    		for(int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
    			//获得当前sheet工作表
        		Sheet sheet = workbook.getSheetAt(sheetNum);
        		if(sheet == null){
        			continue;
        		}
        		//获得当前sheet的开始行
        		int firstRowNum  = sheet.getFirstRowNum();
        		//获得当前sheet的结束行
        		int lastRowNum = sheet.getLastRowNum();
        		//循环所有行
        		for(int rowNum = firstRowNum; rowNum <= lastRowNum;rowNum++){
        			//获得当前行
        			Row row = sheet.getRow(rowNum);
        			if(row == null){
        				continue;
        			}
        			//获得当前行的开始列
        			int firstCellNum = row.getFirstCellNum();
        			if (firstCellNum < 0) {
        				continue;
        			}
        			//获得当前行的列数
        			int lastCellNum = row.getLastCellNum();//row.getPhysicalNumberOfCells();
        			String[] cells = new String[lastCellNum/*row.getPhysicalNumberOfCells()*/];
        			//循环当前行
        			for(int cellNum = firstCellNum; cellNum < lastCellNum;cellNum++){
        				Cell cell = row.getCell(cellNum);
                        cells[cellNum] = getCellValue(cell);
        			}
        			list.add(cells);
        		}
    		}
    		//workbook.close();
    	}
		return list;
    }
	
	public static void checkFile(File file) throws IOException{
		//判断文件是否存在
    	if(null == file){
    		System.err.println("文件不存在！");
    		throw new FileNotFoundException("文件不存在！");
    	}
		//获得文件名
    	String fileName = file.getName();
    	//判断文件是否是excel文件
    	if(!fileName.endsWith(xls) && !fileName.endsWith(xlsx)){
    		System.err.println(fileName + "不是excel文件");
    		throw new IOException(fileName + "不是excel文件");
    	}
	}
	
	public static Workbook getWorkBook(File file) {
		//获得文件名
    	String fileName = file.getName();
    	//创建Workbook工作薄对象，表示整个excel
		Workbook workbook = null;
		try {
			//获取excel文件的io流
			InputStream is = new FileInputStream(file);
			//根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
			if(fileName.endsWith(xls)){
				//2003
				workbook = new HSSFWorkbook(is);
			}else if(fileName.endsWith(xlsx)){
				//2007
				workbook = new XSSFWorkbook(is);
			}
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
		return workbook;
	}
	
	public static String getCellValue(Cell cell){
		String cellValue = null;
		if(cell == null){
			return cellValue;
		}
		//把数字当成String来读，避免出现1读成1.0的情况
//		if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
//			cell.setCellType(Cell.CELL_TYPE_STRING); //日期属于数字，会把日期解析成数字
//		}
		//判断数据的类型
        switch (cell.getCellType()){
	        case Cell.CELL_TYPE_NUMERIC: //数字
	        	if (isDateCell(cell)) {
	        		Date date = cell.getDateCellValue();
	        		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	        		cellValue = sdf.format(date);
	        	} else {
	        		//cellValue = String.valueOf(cell.getNumericCellValue());
	        		cell.setCellType(Cell.CELL_TYPE_STRING);
	        		cellValue = String.valueOf(cell.getStringCellValue());
	        	}
	            break;
	        case Cell.CELL_TYPE_STRING: //字符串
	            cellValue = String.valueOf(cell.getStringCellValue());
	            break;
	        case Cell.CELL_TYPE_BOOLEAN: //Boolean
	            cellValue = String.valueOf(cell.getBooleanCellValue());
	            break;
	        case Cell.CELL_TYPE_FORMULA: //公式
	            cellValue = String.valueOf(cell.getCellFormula());
	            break;
	        case Cell.CELL_TYPE_BLANK: //空值 
	            cellValue = "";
	            break;
	        case Cell.CELL_TYPE_ERROR: //故障
	            cellValue = "非法字符";
	            break;
	        default:
	            cellValue = "未知类型";
	            break;
        }
		return cellValue;
	}
	
	/**
	 * 判断cell类型是否为日期型
	 * @param Cell cell
	 * @return  true 是日期类型  false  否，不是日期类型
	 */
	private static boolean isDateCell(Cell cell) {
		if (cell == null)
			return false;
		boolean isDate = false;
		double d = cell.getNumericCellValue();
		if (DateUtil.isValidExcelDate(d)) {
			CellStyle style = cell.getCellStyle();
			if (style == null)
				return false;
			int i = style.getDataFormat();
			String f = style.getDataFormatString();
			f =  f.replaceAll("[\"|\']","").replaceAll("[年|月|日|时|分|秒|毫秒|微秒]", "");
			isDate = DateUtil.isADateFormat(i, f);
		}
		return isDate;
	}
}
