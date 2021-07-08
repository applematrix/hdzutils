package com.hdz.utils;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

public class ExcelSync {
    private String mSourceFile;
    private String mTargetFile;
    private static final String CONFIG = "init.cfg";
    private static final String SEPERATOR = "\\";

    public ExcelSync(){
        loadConfig();
    }

    public void syncContents() {
        Workbook srcWorkbook = getWorkbook(mSourceFile);
        if (srcWorkbook == null) {
            System.out.println("source file error");
            return;
        }
        Sheet sheet = srcWorkbook.getSheetAt(0);
        if (sheet != null) {
            int rowNum = sheet.getLastRowNum();
            for (int i = 0; i <= rowNum; i++) {
                Row row = sheet.getRow(i);
                int colNum = row.getLastCellNum();
                for (int j = 0; j <= colNum; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        System.out.println(cell.getStringCellValue());
                    }
                }
            }
        }
    }

    private void loadConfig() {
        String workingDir = System.getProperty("user.dir") + SEPERATOR;
        System.out.println("working dir is " + workingDir);
        File configFile = new File(workingDir  + CONFIG);
        System.out.println("working file is " + configFile);
        try {
            Scanner scanner = new Scanner(configFile);
            while (scanner.hasNext()) {
                String config = scanner.nextLine();
                String[] pair = config.split("=");
                String key = pair[0].trim();
                String value = pair[1].trim();
                if (key.equals("src")) {
                    mSourceFile = workingDir + value;
                    System.out.println(mSourceFile);
                } else if (key.equals("target")) {
                    mTargetFile = workingDir + value;
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    private Workbook getWorkbook(String filepath) {
        if (filepath == null) return null;
        try {
            FileInputStream fis = new FileInputStream(new File(filepath));
            return WorkbookFactory.create(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static void main(String[] args) {
        // write your code here
        System.out.println("hello world");
        new ExcelSync().syncContents();
    }
}
