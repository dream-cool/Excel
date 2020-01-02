package com.example.springbootdemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author clt
 * @create 2020/1/2 15:59
 */
public class Test1 {

    public static void main(String[] args) throws IOException {
        String filePath = "C:/Users/mrchen/Desktop/问题用户名称和手机号.xls";
        HSSFWorkbook wookbook = new HSSFWorkbook(new FileInputStream(filePath));
        Writer writer = new FileWriter(new File("C:/Users/mrchen/Desktop/data.txt"),true);
        Reader reader = new FileReader(new File("C:/Users/mrchen/Desktop/新建文本文档 (2).txt"));
        BufferedReader br = new BufferedReader(reader);
        BufferedWriter bw = new BufferedWriter(writer);
        Sheet sheet = (Sheet) wookbook.getSheetAt(0);
        Iterator<Row> it = sheet.rowIterator();
        List<String> list = new ArrayList<String>();
        String str = null;
        while ((str = br.readLine()) != null ){
            if (str.length() > 11){
                list.add(str.substring(0,11));
            }
        }
        while (it.hasNext()) {
            Row row = it.next();
            String phone = row.getCell(1).toString();
            if (!list.contains(phone)){
                bw.write(phone);
                bw.newLine();
                bw.flush();
            }
        }
    }
}
