package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.regex.*;
public class ExcelCleaner {
    public static void main(String[] args) throws IOException {
        // 打开Excel文件
        FileInputStream fis = new FileInputStream(new File("C:\\Users\\l50011273\\Desktop\\小智语料库向量数据库语料.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // 获取第一个工作表
        // 获取你要清洗的列（例如第一列）
        Row row ;
        Cell cell;
        Pattern pattern = Pattern.compile("\\<.*?\\>"); // 替换[你的正则表达式]为你要使用的正则表达式
        //Pattern pattern = Pattern.compile("&gt");
        Matcher matcher;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
           //for (int j = 0; j < row.getLastCellNum(); j++) {
                cell = row.getCell(0);
                if (cell != null) {
                    String cellValue = cell.toString();

                    if (!cellValue.endsWith("？")){

                        cellValue+="？";

                    }

                    matcher = pattern.matcher(cellValue);
                    cellValue = matcher.replaceAll(""); // 使用正则表达式替换掉匹配的内容
                    cell.setCellValue(cellValue);
                }
            //}
        }
        // 将清洗后的数据写入到新的Excel文件
        FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\l50011273\\Desktop\\xiangliang2.1.xlsx"));
        workbook.write(fos);
        fos.close();
        fis.close();
    }
}
