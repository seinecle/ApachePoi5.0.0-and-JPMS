/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.codebangusllc.poiooxmljpmsbroken;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

import static org.apache.poi.ss.usermodel.CellType.STRING;

public final class Main {
    public static void main(String[] args) throws Exception {
        System.out.println(Main.class.getModule());

        var wb = new XSSFWorkbook();
        var sheet = wb.createSheet("bob");
        var row = sheet.createRow(0);
        var cell = row.createCell(0, STRING);
        cell.setCellValue("Hello World!");
        try (final FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            wb.write(fos);
        }
        System.out.println("File written");
    }
}
