package com.data.excel.utils;

import com.data.excel.bean.Student;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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

    /**
     * @param filePath excel文件路径
     * @return Excel文件的行数据的对象结果集
     *  从filePath中的Excel表格中，
     *  将所有行的数据转行成相应的Student对象
     */
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

    /**
     * @param filePath 文件路径（需要寻找的Excel文件）
     * @param dirPath  文件夹路径
     * @return 结果（能够匹配到的结果集）
     * 从filePath的Excel表格中，找到用户的姓名和与之对应的机构名称
     * 然后再从dirPath的文件夹中，递归寻找到对应机构名称的用户信息Excel文件
     * 再遍历对应机构下的用户信息的所有行，如果能找到与之匹配的信息，
     * 则将结果放入list中。
     */
    public static List<Object> getExcelData(String filePath,String dirPath) throws IOException {
        List<Object> list = new ArrayList<>(10);
        XSSFWorkbook wookbook = new XSSFWorkbook(new FileInputStream(filePath));
        for (Sheet sheet : wookbook) {
            Iterator<Row> it = sheet.rowIterator();
            while (it.hasNext()){
                Row row = it.next();
                String id = row.getCell(0).toString();
                String userName = row.getCell(1).toString();
                String category = row.getCell(6).toString();
                String orgName = row.getCell(5).toString();
                if ("1-l".equals(category)){
                    String orgFilePath = recursionFile(orgName,dirPath);
                    if (orgFilePath != null){
                        HSSFWorkbook orgExcel = new HSSFWorkbook(new FileInputStream(orgFilePath));
                        final Iterator<Row> orgIt = orgExcel.getSheetAt(0).rowIterator();
                        orgIt.next();
                        while (orgIt.hasNext()){
                            final Row orgRow = orgIt.next();
                            if (orgRow.getCell(3) == null){
                                break;
                            }
                            if (userName!=null && userName.equals(orgRow.getCell(3).toString())){
                                list.add(id+":"+userName);
                                break;
                            }
                        }
                    }
                }
            }
        }
        return list;
    }

    /**
     * @param filePath 文件路径
     * 如果给定的文件路径是文件夹的话则递归遍历文件夹
     * 将文件夹下的所有Excel表格数据转换成对应的Student对象
     */
    public static void recursionFile(String filePath) throws IOException {
        File file = new File(filePath);
        if (file.getPath().endsWith("xlsx")){
            System.out.println(getExcelData(filePath));
        } else {
            String[] files = file.list();
            if (files != null){
                for (int i = 0; i <  files.length; i++) {
                    recursionFile(filePath+"/"+files[i]);
                }
            }
        }
    }


    /**
     * @param orgName 机构名称
     * @param filePath 文件路径
     * @return 对应的机构名称的文件
     * 从给定的filePath文件路径中递归遍历其所有文件
     * 然后逐一比较目标文件是否是所寻找文件
     * 是的话，则返回结果
     */
    public static String recursionFile(String orgName,String filePath) throws IOException {
        File file = new File(filePath);
        if (file.getPath().endsWith("xlsx")||file.getPath().endsWith("xls")){
            if (file.getPath().contains(orgName)&&file.getPath().contains("委员")){
                return file.getPath();
            }
        } else {
            String[] files = file.list();
            if (files != null){
                for (int i = 0; i <  files.length; i++) {
                    String result = recursionFile(orgName,filePath+"/"+files[i]);
                    if (result != null){
                        return result;
                    }
                }
            }
        }
        return null;
    }

    public static void main(String[] args) throws IOException {
//        String filePath ="C:/Users/Mrchen/Desktop/新建文件夹 (2)";
        String filePath ="C:/Users/Mrchen/Desktop/无状态用户.xls";
        String dirPath ="C:/Users/Mrchen/Desktop/现场返回的数据";
        System.out.println(getExcelData(filePath,dirPath));
    }

}
