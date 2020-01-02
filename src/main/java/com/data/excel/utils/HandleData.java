package com.data.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;

/**
 * @author clt
 * @create 2020/1/2 13:39
 */
public class HandleData {
    public static int count = 0;

    public static List<Object> getWorkerExcelData(String filePath, String dirPath) throws IOException {
        List<Object> list = new ArrayList<Object>(10);
        Set<String> set = new HashSet<String>();
        HSSFWorkbook wookbook = new HSSFWorkbook(new FileInputStream(filePath));
        Sheet sheet = (Sheet) wookbook.getSheetAt(0);
        Iterator<Row> it = sheet.rowIterator();
        Writer writer = new FileWriter(new File("C:/Users/mrchen/Desktop/worker.txt"),true);
        BufferedWriter bw = new BufferedWriter(writer);
        while (it.hasNext()) {
            Row row = it.next();
            String phone = row.getCell(1).toString();
            File file = new File(dirPath);
            String workerPath = null;
            try {
                workerPath = dirPath;
                HSSFWorkbook workerExcel = new HSSFWorkbook(new FileInputStream(workerPath));
                Iterator<Row> workerIt = workerExcel.getSheetAt(0).rowIterator();
                int count = 0;
                while (workerIt.hasNext()) {
                    Row worker = workerIt.next();
                    String workerPhone = null;
                    if (worker.getCell(4) != null) {
                        workerPhone = worker.getCell(4).toString();
                        if (worker.getCell(4).toString().length() == 15) {
                            workerPhone = new BigDecimal(workerPhone).toString();
                        }
                    } else if (worker.getCell(4) == null && count == 0) {
                        set.add(workerPath + ":" + worker.getCell(0) + ":");
                    }
                    count++;
                    if (phone.equals(workerPhone)) {
                        bw.write(phone+"------"+worker.getCell(12));
                        bw.newLine();
                        bw.flush();
                        System.out.println(phone + "----" + workerPath);
                        break;
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println(workerPath);
            }
        }
        System.out.println(set);
        return null;
    }


    public static List<Object> getCommitteeMemberExcelData(String filePath, String dirPath) throws IOException {
        List<Object> list = new ArrayList<Object>(10);
        Set<String> set = new HashSet<String>();
        HSSFWorkbook wookbook = new HSSFWorkbook(new FileInputStream(filePath));
        Sheet sheet = (Sheet) wookbook.getSheetAt(0);
        Iterator<Row> it = sheet.rowIterator();
        File fi = new File("C:/Users/mrchen/Desktop/workerdata.txt");
        Writer writer = new FileWriter(fi,true);
        BufferedWriter bw=new BufferedWriter(writer);
        while (it.hasNext()) {
            Row row = it.next();
            String phone = row.getCell(1).toString();
            File file = new File(dirPath);
            String workerPath = null;
            try {
                workerPath = dirPath;
                XSSFWorkbook workerExcel = new XSSFWorkbook(new FileInputStream(workerPath));
                Iterator<Row> workerIt = workerExcel.getSheetAt(1).rowIterator();
                int count = 0;
                while (workerIt.hasNext()) {
                    Row worker = workerIt.next();
                    String workerPhone = null;
                    if (worker.getCell(2) != null) {
                        workerPhone = worker.getCell(2).toString();
                        if (worker.getCell(2).toString().length() == 15) {
                            workerPhone = new BigDecimal(workerPhone).toString();
                        }
                    } else if (worker.getCell(2) == null && count == 0) {
                        set.add(workerPath + ":" + worker.getCell(0) + ":");
                    }
                    count++;
                    if (phone.equals(workerPhone)) {
                        bw.write(workerPhone+"----"+worker.getCell(27));
                        bw.newLine();
                        bw.flush();
                        System.out.println(phone + "----" + workerPath);
                        break;
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println(workerPath);
            }
        }
        System.out.println(set);
        return null;
    }

    public static void recursionFile(String dirPath) throws IOException {
        File file = new File(dirPath);
        if (file.getPath().endsWith("xls")) {
        } else {
            String[] files = file.list();
            if (files != null) {
                for (int i = 0; i < files.length; i++) {
                    recursionFile(dirPath + "/" + files[i]);
                }
            }
        }
    }

    public static void getRepeatData() throws IOException {
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

    public static void main(String[] args) throws IOException {
        String filePath = "C:/Users/mrchen/Desktop/问题用户名称和手机号.xls";
        String workerDirPath = "C:/Users/mrchen/Desktop/工作人员汇总表.xls";
//        String dirPath = "C:/Users/mrchen/Desktop/合并-修订.xlsx";
//        getWorkerExcelData(filePath, workerDirPath);
        getCommitteeMemberExcelData(filePath,workerDirPath);
    }
}
