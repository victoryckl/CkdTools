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
 * excel��д��������Ҫ��jar:
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
	 * ����excel�ļ��������󷵻�
	 * @param file
	 * @throws IOException 
	 */
	public static List<String[]> readExcel(File file) throws IOException{
		//����ļ�
		checkFile(file);
    	//���Workbook����������
    	Workbook workbook = getWorkBook(file);
    	//�������ض��󣬰�ÿ���е�ֵ��Ϊһ�����飬��������Ϊһ�����Ϸ���
    	List<String[]> list = new ArrayList<String[]>();
    	if(workbook != null){
    		for(int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
    			//��õ�ǰsheet������
        		Sheet sheet = workbook.getSheetAt(sheetNum);
        		if(sheet == null){
        			continue;
        		}
        		//��õ�ǰsheet�Ŀ�ʼ��
        		int firstRowNum  = sheet.getFirstRowNum();
        		//��õ�ǰsheet�Ľ�����
        		int lastRowNum = sheet.getLastRowNum();
        		//ѭ��������
        		for(int rowNum = firstRowNum; rowNum <= lastRowNum;rowNum++){
        			//��õ�ǰ��
        			Row row = sheet.getRow(rowNum);
        			if(row == null){
        				continue;
        			}
        			//��õ�ǰ�еĿ�ʼ��
        			int firstCellNum = row.getFirstCellNum();
        			if (firstCellNum < 0) {
        				continue;
        			}
        			//��õ�ǰ�е�����
        			int lastCellNum = row.getLastCellNum();//row.getPhysicalNumberOfCells();
        			String[] cells = new String[lastCellNum/*row.getPhysicalNumberOfCells()*/];
        			//ѭ����ǰ��
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
		//�ж��ļ��Ƿ����
    	if(null == file){
    		System.err.println("�ļ������ڣ�");
    		throw new FileNotFoundException("�ļ������ڣ�");
    	}
		//����ļ���
    	String fileName = file.getName();
    	//�ж��ļ��Ƿ���excel�ļ�
    	if(!fileName.endsWith(xls) && !fileName.endsWith(xlsx)){
    		System.err.println(fileName + "����excel�ļ�");
    		throw new IOException(fileName + "����excel�ļ�");
    	}
	}
	
	public static Workbook getWorkBook(File file) {
		//����ļ���
    	String fileName = file.getName();
    	//����Workbook���������󣬱�ʾ����excel
		Workbook workbook = null;
		try {
			//��ȡexcel�ļ���io��
			InputStream is = new FileInputStream(file);
			//�����ļ���׺����ͬ(xls��xlsx)��ò�ͬ��Workbookʵ�������
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
		//�����ֵ���String�������������1����1.0�����
//		if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
//			cell.setCellType(Cell.CELL_TYPE_STRING); //�����������֣�������ڽ���������
//		}
		//�ж����ݵ�����
        switch (cell.getCellType()){
	        case Cell.CELL_TYPE_NUMERIC: //����
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
	        case Cell.CELL_TYPE_STRING: //�ַ���
	            cellValue = String.valueOf(cell.getStringCellValue());
	            break;
	        case Cell.CELL_TYPE_BOOLEAN: //Boolean
	            cellValue = String.valueOf(cell.getBooleanCellValue());
	            break;
	        case Cell.CELL_TYPE_FORMULA: //��ʽ
	            cellValue = String.valueOf(cell.getCellFormula());
	            break;
	        case Cell.CELL_TYPE_BLANK: //��ֵ 
	            cellValue = "";
	            break;
	        case Cell.CELL_TYPE_ERROR: //����
	            cellValue = "�Ƿ��ַ�";
	            break;
	        default:
	            cellValue = "δ֪����";
	            break;
        }
		return cellValue;
	}
	
	/**
	 * �ж�cell�����Ƿ�Ϊ������
	 * @param Cell cell
	 * @return  true ����������  false  �񣬲�����������
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
			f =  f.replaceAll("[\"|\']","").replaceAll("[��|��|��|ʱ|��|��|����|΢��]", "");
			isDate = DateUtil.isADateFormat(i, f);
		}
		return isDate;
	}
}
