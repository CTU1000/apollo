/**
 * Alipay.com Inc.
 * Copyright (c) 2004-2018 All Rights Reserved.
 */
package com.ctrip.framework.apollo.portal.sheet;

import com.google.common.base.Joiner;
import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.List;

/**
 *
 * @author lepdou
 * @version $Id: Excel.java, v 0.1 2018年02月28日 下午7:53 lepdou Exp $
 */
public class Excel {
    static NumberFormat formatter = new DecimalFormat("#000000000000.00");
    static Joiner       joiner    = Joiner.on(",");

    public static void main(String[] args) throws Exception {
        //System.out.println(format("0000251899.20000"));
        List<Sheet> sheets = read();
        System.out.println(sheets.size());

        List<String> units = Lists.newArrayList();
        for (Sheet sheet : sheets) {
            units.add(sheet.getUnit());
        }

        List<Double> values = Lists.newArrayList();
        List<List<String>> result = Lists.newArrayList();

        for (int i = 0; i < 70; i++) {
            Double sum = 0d;
            List<String> str = Lists.newArrayList();
            for (Sheet sheet : sheets) {
                double d = sheet.getData(i);
                sum += d;
                str.add(String.valueOf(format(formatter.format(d))));
            }
            values.add(sum);
            result.add(str);
        }

        //for (Double d : values) {
        //    System.out.println(format(formatter.format(d)));
        //}
        System.out.println(joiner.join(units));
        for (List<String> strings : result) {
            System.out.println(joiner.join(strings));
        }
    }

    private static List<Sheet> read() throws Exception {
        List<Sheet> sheets = Lists.newLinkedList();
        XSSFWorkbook workbook = null;

        InputStream is = new FileInputStream("/Users/lepdou/Documents/sheet.xlsm");
        workbook = new XSSFWorkbook(is);

        int sheetCount = workbook.getNumberOfSheets();

        for (int i = 0; i < sheetCount; i++) {
            XSSFSheet sheetDoc = workbook.getSheetAt(i);
            //获取单位
            String unit = sheetDoc.getRow(1).getCell(0).getStringCellValue();

            Sheet sheet = new Sheet(unit, sheetDoc.getSheetName());

            int sumColumnIndex = findSumColumn(sheetDoc);
            if (sumColumnIndex == -1) {
                System.err.println("找不到合计列。" + sheetDoc.getSheetName());
                return null;
            }

            //遍历行
            for (int j = 4; j < 74; j++) {
                XSSFRow row = sheetDoc.getRow(j);
                XSSFCell sumCell = row.getCell(sumColumnIndex);
                double value = sumCell.getNumericCellValue();
                sheet.addData(value);
                //System.out.println(row.getCell(0).getStringCellValue() + " = " + value);
            }

            sheets.add(sheet);
            //break;
            //System.out.println("===============");
        }

        return sheets;
    }

    private static int findSumColumn(XSSFSheet sheetDoc) {
        XSSFRow row = sheetDoc.getRow(2);
        int count = row.getLastCellNum();
        for (int i = 0; i < count; i++) {
            XSSFCell cell = row.getCell(i);
            if ("合计".equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    private static String format(String value) {
        if (value.equals("000000000000.00")) {
            return "0";
        }
        boolean nav = value.charAt(0) == '-';
        int start = 0;
        int end = value.length() - 1;
        while (start < end) {
            if (value.charAt(start) == '0' || value.charAt(start) == '-') {
                start++;
            } else {
                break;
            }
        }
        while (end > start) {
            if (value.charAt(end) == '0') {
                end--;
            } else {
                break;
            }
        }
        String result = value.substring(start, end + 1);
        return nav ? "-" + result : result;
    }

}