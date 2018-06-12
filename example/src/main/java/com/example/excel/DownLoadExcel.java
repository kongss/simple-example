package com.example.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DownLoadExcel {
    public static void main(String[] args) throws Exception{
        Map<String, Object> map = new HashMap<>();
        Map<String, Object> map1 = new HashMap<>();
        for (int i=1;i<=10;i++){
            map.put("K"+i,"V"+i);
        }
        for (int i=1;i<=10;i++){
            map1.put("K"+i,"S"+i);
        }
        List<Map<String, Object>> lis = new ArrayList<>();
        lis.add(map);
        lis.add(map1);
        System.out.println(lis);
        //Excel文件头
        String[] headNameArr = new String[]{"列1","列2","列3","列4","列5","列6","列7","列8","列9","列10"};
        String[] fieldNameArr = new String[]{"K1","K2","K3","K4","K5","K6","K7","K8","K9","K10"};
        //设置每列宽度
        Integer[] integers = {30,30,30,30,30,30,30,30,30,30};
        exportSXSSFExcel(lis,1000,headNameArr,integers,fieldNameArr);
    }

    public static void exportSXSSFExcel(List<Map<String, Object>> lis,Integer maxRows,String[] headNameArr,Integer[] integers,String[] fieldNameArr)throws Exception{
        SXSSFWorkbook workbook = new SXSSFWorkbook(maxRows);
        SXSSFSheet sheet = workbook.createSheet();
        int rowIndex = 1;
        Row row;
        Cell cell;

        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        //导出文件头
        row = sheet.createRow(0);
        for(int i=0; i<headNameArr.length; i++){
            cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(headNameArr[i]);
        }

        int listSize = lis.size();
        for (int i=0;i<listSize;i++){
            row = sheet.createRow(rowIndex++);
            Map<String, Object> map = lis.get(i);
            for(int j=0;j<fieldNameArr.length;j++){
                String value = String.valueOf(map.get(fieldNameArr[j]));
                cell = row.createCell(j);
                cell.setCellValue(value);
            }
        }
        FileOutputStream os = new FileOutputStream("D:\\Demo.xlsx");
        workbook.write(os);
        os.close();
        System.out.println("导出结束");
    }



}
