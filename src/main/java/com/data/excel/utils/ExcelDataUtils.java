package com.data.excel.utils;

import com.data.excel.bean.Student;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author ：clt
 * @Date ：Created in 21:16 2019/12/31
 */
public class ExcelDataUtils {

    private ExcelDataUtils(){}

    public static int count = 0;

    public static List<Object> getExcelData(String filePath) throws IOException {
        List<Object> list = new ArrayList<>(10);
        XSSFWorkbook wookbook = new XSSFWorkbook(new FileInputStream(filePath));
        for (Sheet sheet : wookbook) {
            Iterator<Row> it = sheet.rowIterator();
            while (it.hasNext()){
                Row row = it.next();
                Student student = new Student();
                for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                    student.setProperty(i+1,row.getCell(i).toString());
                }
                if ("学号".equals(student.getStuNum())){
                    continue;
                }
                if ("陈留涛".equals(student.getName())){
                    count++;
                }
                list.add(student);
            }
        }
        return list;
    }

    public static void recursionFile(String filePath) throws IOException {
        File file = new File(filePath);
        if (file.getPath().endsWith("xlsx")){
            System.out.println(getExcelData(filePath));
        } else {
            String[] files = file.list();
            if (files != null){
                for (int i = 0; i <  files.length; i++) {
                    recursionFile(files[i]);
                }
            }
        }
    }

    public static void main(String[] args) throws IOException {
        String filePath ="C:/Users/Mrchen/Desktop/新建文件夹 (2)";
        recursionFile(filePath);
        System.out.println(count);
    }
}
