package com.example.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;

public class AnalysisExcel {

    public static void main(String[] args) throws Exception{
        toInExcel();
    }
    static void toInExcel() throws Exception{
        String filepath = "D:/Demo.xlsx";
        InputStream is = new FileInputStream(filepath);
        //InputStream is = file.getInputStream();

        //获取文件后缀(2003)xls|(2007)xlsx
        //String fileName = file.getOriginalFilename();
        //String suffix = fileName.substring(fileName.lastIndexOf("."));

        //2003版本
        //Workbook workbook = new HSSFWorkbook(is);
        //2007版本
        //Workbook workbook = new XSSFWorkbook(is);
        //推荐使用poi-ooxml中的WorkbookFactory.create(is)来创建Workbook,因为HSSFWorkbook和XSSFWorkbook都实现了Workbook接口
        Workbook wbc = WorkbookFactory.create(is);
        Sheet sheet = wbc.getSheetAt(0);
        int rowNum = sheet.getLastRowNum();
        Row row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        System.out.println("rowNum:"+rowNum+"   colNum:"+colNum);
        //从1开始，跳过表头的标题
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            for (int j = 0; j < colNum; j++){
                Object obj = getCellFormatValue(row.getCell(j));
                System.out.println("obj["+i+"*"+j+"]======================"+obj);
            }
        }
    }

    static Object getCellFormatValue(Cell cell) {
        Object cellvalue = "";
        if (cell != null) {
            /*_NONE(-1),NUMERIC(0),STRING(1),FORMULA(2),BLANK(3),BOOLEAN(4),ERROR(5);*/
            switch (cell.getCellTypeEnum()) {
                case _NONE: {
                    cellvalue = "_NONE";
                    break;
                }
                case NUMERIC: {
                    cellvalue = cell.getNumericCellValue();
                    break;
                }
                case STRING: {
                    cellvalue = cell.getStringCellValue();
                    break;
                }
                case FORMULA: {
                    cellvalue = cell.getCellFormula();
                    break;
                }
                case BLANK: {
                    cellvalue = "";
                    break;
                }
                case BOOLEAN: {
                    cellvalue = cell.getBooleanCellValue();
                    break;
                }
                case ERROR: {
                    cellvalue = cell.getErrorCellValue();
                    break;
                }
                default: {
                    cellvalue = "";
                    break;
                }
            }
        }
        return cellvalue;
    }
}
