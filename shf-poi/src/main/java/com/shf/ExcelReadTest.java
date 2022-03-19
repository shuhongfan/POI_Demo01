package com.shf;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Date;

import static org.apache.poi.ss.usermodel.Cell.*;

public class ExcelReadTest {
    String PATH = "C:\\Users\\shuho\\Documents\\Code\\POI_Demo01\\shf-poi\\src\\main\\java\\com\\";

    @Test
    public void testWrite03() throws Exception {
//        获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "shf狂神观众统计表03.xls");

//        1. 创建一个工作簿。 使用excel能操作这边他都可以操作！
        Workbook workbook = new HSSFWorkbook(inputStream);
//        2.得到表
        Sheet sheet = workbook.getSheetAt(0);
//        3.得到行
        Row row = sheet.getRow(0);
//        4.得到列
        Cell cell = row.getCell(0);

//        字符串类型
        System.out.println(cell.getStringCellValue());

//        数字类型
//        System.out.println(cell.getNumericCellValue());

        inputStream.close();
    }

    @Test
    public void testWrite07() throws Exception {
//        获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "shf狂神观众统计表07.xlsx");

//        1. 创建一个工作簿。 使用excel能操作这边他都可以操作！
        Workbook workbook = new XSSFWorkbook(inputStream);
//        2.得到表
        Sheet sheet = workbook.getSheetAt(0);
//        3.得到行
        Row row = sheet.getRow(0);
//        4.得到列
        Cell cell = row.getCell(1);

//        字符串类型
//        System.out.println(cell.getStringCellValue());

//        数字类型
        System.out.println(cell.getNumericCellValue());

        inputStream.close();
    }

//    @Test
//    public void testCellType() throws Exception {
////        获取文件流
//        FileInputStream inputStream = new FileInputStream(PATH + "明细表");
//
////        1.创建一个工作簿.使用excel能操作的这边他都可以操作
//        Workbook workbook = new HSSFWorkbook(inputStream);
//        Sheet sheet = workbook.getSheetAt(0);
////        获取标题
//        Row rowTitle = sheet.getRow(0);
//        if (rowTitle!=null){
//            int cellCount = rowTitle.getPhysicalNumberOfCells();
//            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                Cell cell = rowTitle.getCell(cellNum);
//                if (cell!=null) {
//                    int cellType = cell.getCellType();
//                    String cellValue = cell.getStringCellValue();
//                    System.out.println(cellValue+" | ");
//                }
//            }
//            System.out.println();
//        }
//
////        获取表中的内容
//        int rowCount = sheet.getPhysicalNumberOfRows();
//        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
//            Row rowData = sheet.getRow(rowNum);
//            if (rowData!=null){
//                int cellCount = rowTitle.getPhysicalNumberOfCells();
//                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                    System.out.println("【"+(rowNum+1)+"-"+(rowNum+1)+"】");
//
//                    Cell cell = rowData.getCell(cellNum);
////                    匹配列的数据类型
//                    if (cell!=null){
//                        int cellType = cell.getCellType();
//                        String cellValue = "";
//
//                        switch (cellType){
//                            case CELL_TYPE_STRING: // 字符串
//                                System.out.println("[String]");
//                                cellValue = cell.getStringCellValue();
//                                break;
//                            case CELL_TYPE_BOOLEAN: // 布尔
//                                System.out.println("[Boolean]");
//                                cellValue = String.valueOf(cell.getBooleanCellValue());
//                                break;
//                            case CELL_TYPE_BLANK: // 空
//                                System.out.println("[Blank]");
//                                cellValue = String.valueOf(cell.getBooleanCellValue());
//                                break;
//                            case CELL_TYPE_NUMERIC: // 数字（日期、普通数字）
//                                System.out.println("[NUMERIC]");
//                                if (HSSFDateUtil.isCellDateFormatted(cell)){
//                                    System.out.println("[日期]");
//                                    Date date = cell.getDateCellValue();
//                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
//                                } else {
////                                    不是日期模式,防止数字过长
//                                    System.out.println("转换为字符串输出");
//                                    cell.setCellType(CELL_TYPE_STRING);
//                                    cellValue = cell.toString();
//                                }
//                                break;
//                            case CELL_TYPE_ERROR: // 空
//                                System.out.println("[数据类型错误]");
//                                break;
//                        }
//                    }
//                }
//            }
//        }
//    }

//    @Test
//    public void testCellType(FileInputStream inputStream) throws Exception {
////        1.创建一个工作簿.使用excel能操作的这边他都可以操作
//        Workbook workbook = new HSSFWorkbook(inputStream);
//        Sheet sheet = workbook.getSheetAt(0);
////        获取标题
//        Row rowTitle = sheet.getRow(0);
//        if (rowTitle!=null){
//            int cellCount = rowTitle.getPhysicalNumberOfCells();
//            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                Cell cell = rowTitle.getCell(cellNum);
//                if (cell!=null) {
//                    int cellType = cell.getCellType();
//                    String cellValue = cell.getStringCellValue();
//                    System.out.println(cellValue+" | ");
//                }
//            }
//            System.out.println();
//        }
//
////        获取表中的内容
//        int rowCount = sheet.getPhysicalNumberOfRows();
//        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
//            Row rowData = sheet.getRow(rowNum);
//            if (rowData!=null){
//                int cellCount = rowTitle.getPhysicalNumberOfCells();
//                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                    System.out.println("【"+(rowNum+1)+"-"+(rowNum+1)+"】");
//
//                    Cell cell = rowData.getCell(cellNum);
////                    匹配列的数据类型
//                    if (cell!=null){
//                        int cellType = cell.getCellType();
//                        String cellValue = "";
//
//                        switch (cellType){
//                            case CELL_TYPE_STRING: // 字符串
//                                System.out.println("[String]");
//                                cellValue = cell.getStringCellValue();
//                                break;
//                            case CELL_TYPE_BOOLEAN: // 布尔
//                                System.out.println("[Boolean]");
//                                cellValue = String.valueOf(cell.getBooleanCellValue());
//                                break;
//                            case CELL_TYPE_BLANK: // 空
//                                System.out.println("[Blank]");
//                                cellValue = String.valueOf(cell.getBooleanCellValue());
//                                break;
//                            case CELL_TYPE_NUMERIC: // 数字（日期、普通数字）
//                                System.out.println("[NUMERIC]");
//                                if (HSSFDateUtil.isCellDateFormatted(cell)){
//                                    System.out.println("[日期]");
//                                    Date date = cell.getDateCellValue();
//                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
//                                } else {
////                                    不是日期模式,防止数字过长
//                                    System.out.println("转换为字符串输出");
//                                    cell.setCellType(CELL_TYPE_STRING);
//                                    cellValue = cell.toString();
//                                }
//                                break;
//                            case CELL_TYPE_ERROR: // 空
//                                System.out.println("[数据类型错误]");
//                                break;
//                        }
//                    }
//                }
//            }
//        }
//    }

//    计算公式
//    @Test
//    public void testFormulate() throws Exception {
//        FileInputStream inputStream = new FileInputStream(PATH + "公式.xlsx");
//        Workbook workbook = new XSSFWorkbook(inputStream);
//        Sheet sheet = workbook.getSheetAt(0);
//
//        Row row = sheet.getRow(4);
//        Cell cell = row.getCell(0);
//
////        拿到计算公式 eval
//        FormulaEvaluator FormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
//
////        输出单元格的内容
//        int cellType = cell.getCellType();
//        switch (cellType){
//            case CELL_TYPE_FORMULA:
//                String formula = cell.getCellFormula();
//                System.out.println(formula);
//
////                计算
//                CellValue evaluate = FormulaEvaluator.evaluate(cell);
//                String cellValue = evaluate.formatAsString();
//                System.out.println(cellValue);
//                break;
//        }
//    }
}
