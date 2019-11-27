package com.atguigu.excelpoi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Date;

public class ExcelReadTest {
    /**
     * 读取03版
     * @throws Exception
     */
    @Test
    public void testRead03() throws Exception {
        //1.创建读取的线程
        InputStream In = new FileInputStream("F:/1java/day07在线教育/在线教育/资源/excel/课程分类列表模板.xls");
        //2.拿到地址中对应的HSSFWorkbook
        Workbook workbook = new HSSFWorkbook(In);
        //3.拿到表实例的操控权,获取第一张表（一个workbook可以有多张表）
        Sheet sheet = workbook.getSheetAt(0);

        //4.拿取表格第一行的内容，就是表的标题
        Row rowtitle = sheet.getRow(0);
        if (rowtitle != null) {
            //看有几列
            int cellCount = rowtitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum <cellCount ; cellNum++) {
                Cell cell = rowtitle.getCell(cellNum);//遍历拿到所有列中的数据

                if (cell!=null){ //要是单元格内容不是空
                    int cellType = cell.getCellType();//判断第一行单元格中内容的类型
                    String cellValue = cell.getStringCellValue();
                    System.out.println(cellValue+"|");
                }
            }
        }
        //读取商品列表数据
        int rowCount = sheet.getPhysicalNumberOfRows();//看有几行
        for (int rowNum = 0; rowNum <rowCount ; rowNum++) {//遍历所有行

            Row rowData = sheet.getRow(rowNum);//拿到所有行中的数据
            if (rowData != null) {

                //通过rowData读取cell
                int cellCount = rowData.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum <cellCount ; cellNum++) {//遍历所有列
                    System.out.println("【" + (rowNum + 1) + "-" + (cellNum + 1) + "】");

                    Cell cell = rowData.getCell(cellNum);//每一行的数据有了，就可以拿到每一个单元格的数据
                    if (cell != null) {
                        int cellType = cell.getCellType();//拿到所有单元格（列）的数据类型

                        //判断单元格数据类型
                        String cellValue = "";
                        switch (cellType) {
                            case HSSFCell.CELL_TYPE_STRING://字符串
                                System.out.print("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;

                            case HSSFCell.CELL_TYPE_BOOLEAN://布尔
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;

                            case HSSFCell.CELL_TYPE_BLANK://空
                                System.out.print("【BLANK】");
                                break;

                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print("【NUMERIC】");
                                //cellValue = String.valueOf(cell.getNumericCellValue());

                                if (HSSFDateUtil.isCellDateFormatted(cell)) {//日期
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                } else {
                                    // 不是日期格式，则防止当数字过长时以科学计数法显示
                                    System.out.print("【转换成字符串】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;

                            case Cell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                        }

                        System.out.println(cellValue);
                    }
                }
            }
        }

        In.close();
    }
    public void testRead03Simple() throws Exception {
        FileInputStream In = new FileInputStream("F:/1java/day07在线教育/在线教育/资源/excel/课程分类列表模板.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(In);

        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(2);
        String cellValue = cell.getStringCellValue();
        System.out.println(cellValue);
    }
}
