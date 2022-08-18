package com.example.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class WriteExcelDemo 
{
    public static void main( String[] args )
    {
        System.out.println("Hello World!");
        
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Employee Data");

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] { "ID", "NAME", "SALARY" });
        data.put("2", new Object[] { 1, "Amit", 20000 });
        data.put("3", new Object[] { 2, "Jai", 30000 });

        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }
        
        try {
            FileOutputStream out = new FileOutputStream(new File("emp.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Excel written successfully..");
        } catch (Exception e) {
            e.printStackTrace();
        }
        
    }
}
