package com.ixycoder.diffexcel.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelUtils {
    public static void main(String[] args) throws IOException {
        Workbook workbook = readExcel("/tmp/test0.xlsx");
        Workbook workbook1 = readExcel("/tmp/test1.xlsx");

        diffWorkbook(workbook, workbook1);

        workbook1.close();
        workbook.close();
    }

    public static void diffWorkbook(Workbook workbook0, Workbook workbook1) {
        //获得sheet的数量(sheet的index是从0开始的)
        int sheetCount0 = workbook0.getNumberOfSheets();
        int sheetCount1 = workbook1.getNumberOfSheets();
        System.out.println("左边" + sheetCount0 + "个sheet,右边" + sheetCount1 + "个sheet");
        int sheetCount = Math.min(sheetCount0, sheetCount1);
        if (sheetCount0 != sheetCount1) {
            System.out.println("只扫描前"+ sheetCount + "个sheet,忽略超过部分!");
        }
        //遍历Sheet
        for(int i = 0;i < sheetCount;i++){
            diffSheet(i, workbook0.getSheetAt(i), workbook1.getSheetAt(i));
        }
    }

    public static void diffSheet(int idx, Sheet sheet0, Sheet sheet1) {
        //得到每个Sheet的行数,此工作表中包含的最后一行(Row的index是从0开始的)
        int rowCount0 = sheet0.getLastRowNum();
        int rowCount1 = sheet1.getLastRowNum();
        System.out.println("第" + idx + "个sheet(从0开始),左边总计" + rowCount0 + "行,右边总计" + rowCount1 + "行");
        int rowCount = Math.min(rowCount0, rowCount1);
        if (rowCount0 != rowCount1) {
            System.out.println("只扫描前"+ rowCount + "行,忽略超过部分!");
        }
        //遍历Row
        for(int j = 0 ;j <= rowCount;j++) {
            //得到Row
            Row row0 = sheet0.getRow(j);
            Row row1 = sheet1.getRow(j);

            if (row0 == null) {
                if (row1 != null) {
                    System.out.println("第" + idx + "sheet,第" + j + "行(从0开始) 左边为空,右边不为空!");
                }
                continue;
            }

            if (row1 == null) {
                System.out.println("第" + idx + "sheet,第" + j + "行(从0开始) 左边不为空,右边为空!");
                continue;
            }
            diffRow(j, row0, row1);
        }
    }

    public static void diffRow(int idx, Row row0, Row row1) {
        int cellCount0 = row0.getLastCellNum();
        int cellCount1 = row1.getLastCellNum();
        int num = Math.min(cellCount0, cellCount1);
        System.out.println("第" + idx + "行(从0开始) 左边总计" + cellCount0 + "列,右边总计" + cellCount1 + "列");
        if (cellCount0 > cellCount1) {
            System.out.println("第"+idx+"行(从0开始) 左边比较大,超过部分不比较!");
        }
        if (cellCount0 < cellCount1) {
            System.out.println("第"+idx+"行(从0开始) 右边比较大,超过部分不比较!");
        }
        for (int n = 0; n < num; n ++) {
            if (!diffCell(row0.getCell(n), row1.getCell(n))) {
                System.out.println("第"+idx+"行,第"+n+"个单元格(从0开始): " + row0.getCell(n) + " : " + row1.getCell(n));
            }
        }

    }

    public static boolean diffCell(Cell cell0, Cell cell1) {
        return getCellFormatValue(cell0).equals(getCellFormatValue(cell1));
    }


    /**
     * 根据文件地址，解析后缀返回不同的Workbook对象
     * @param filePath 文件地址
     * @return Excel文档对象Workbook
     */
    public static Workbook readExcel(String filePath){

        if(filePath == null || filePath.equals("")){
            return null;
        }
        //得到文件后缀
        String suffix = filePath.substring(filePath.lastIndexOf("."));
//        System.out.println(suffix);
        try {
            InputStream is = new FileInputStream(filePath);
            if(".xls".equals(suffix)){
//                System.out.println("文件类型是.xls");
                return new HSSFWorkbook(is);
            }
            if(".xlsx".equals(suffix)){
//                System.out.println("文件类型是.xlsx");
                return new XSSFWorkbook(is);
            }
            return null;

        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("文件没有找到");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("发生io异常");
        }
        return null;
    }

    public static Object getCellFormatValue(Cell cell){
        Object cellValue;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                //空值单元格
                case BLANK: {
                    cellValue = "";
                    break;
                }
                //数值型单元格 getNumericCellValue()以数字形式获取单元格的值。
                case NUMERIC: {
                    //判断cell是否为日期格式
                    if(DateUtil.isCellDateFormatted(cell)){
                        //转换为日期格式YYYY-mm-dd
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        cellValue = dateFormat.format(date);
                    }else{
                        //数字
                        cellValue = cell.getNumericCellValue();
                    }
                    break;
                }
                //公式型单元格getCellFormula()返回单元格的公式
                case FORMULA: {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                //字符串单元格
                case STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                //布尔值型单元格
                case BOOLEAN: {
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }
}
