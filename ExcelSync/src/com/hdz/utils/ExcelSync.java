package com.hdz.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

public class ExcelSync {
    private String mSourceFile;
    private String mTargetFile;
    private int mRefCol;
    private String mRefValue;
    private int mFillFromCol;
    private int mKeyCol;
    private static final String CONFIG = "init.cfg";
    private static final String SEPERATOR = "\\";
    private Map<String, ArrayList<String>> mFillContents = new HashMap<>();

    public ExcelSync(){
        loadConfig();
    }

    public void syncContents() {
        pickFillContents();
        syncToTarget();
    }

    private void pickFillContents() {
        System.out.println("picking fill contents...");
        Workbook srcWorkbook = getWorkbook(mSourceFile);
        if (srcWorkbook == null) {
            System.out.println("source file error");
            return;
        }
        Sheet sheet = srcWorkbook.getSheetAt(0);
        if (sheet != null) {
            int rowNum = sheet.getLastRowNum();
            for (int i = 1; i <= rowNum; i++) {
                Row row = sheet.getRow(i);
                int colNum = row.getLastCellNum();

                Cell refCell = row.getCell(mRefCol);
                if (refCell == null) continue;
                if (refCell.getCellType() != CellType.STRING) continue;
                if (mRefValue.equals(refCell.getStringCellValue())) {
                    int fillCol = mFillFromCol;
                    ArrayList<String> lineContent = new ArrayList<>();
                    for (; fillCol < colNum; fillCol++) {
                        Cell pickCell = row.getCell(fillCol);
                        if (pickCell == null) {
                            lineContent.add(null);
                        } else {
                            CellType type = pickCell.getCellType();
                            if (type == CellType.STRING) {
                                lineContent.add(pickCell.getStringCellValue());
                            } else if (type == CellType.NUMERIC) {
                                lineContent.add(String.valueOf(pickCell.getNumericCellValue()));
                            }
                        }
                    }
                    if (!lineContent.isEmpty()) {
                        Cell keyCell = row.getCell(mKeyCol);
                        mFillContents.put(keyCell.getStringCellValue(), lineContent);
                    }
                }
            }
        }
    }

    private void syncToTarget() {
        System.out.println("sync to target...");
        try (FileOutputStream fos = new FileOutputStream(new File(mTargetFile))){
            FileInputStream fis = new FileInputStream(new File(mTargetFile));
            Workbook target = WorkbookFactory.create(fis);
            if (target == null) return;
            Sheet sheet = target.getSheetAt(0);
            if (sheet == null) return;
            int rowNum = sheet.getLastRowNum();

            for (int i = 1; i <= rowNum; i++) {
                Row row = sheet.getRow(i);

                Cell keyCell = row.getCell(mKeyCol);
                if (keyCell == null) continue;
                if (keyCell.getCellType() != CellType.STRING) continue;

                if (mFillContents.containsKey(keyCell.getStringCellValue())) {
                    ArrayList<String> lineContent = mFillContents.get(keyCell.getStringCellValue());
                    for (int j=0; j < lineContent.size(); j++) {
                        int fillCol = mFillFromCol+j;
                        Cell fillCell = row.getCell(fillCol);
                        if (fillCell == null) {
                            fillCell = row.createCell(fillCol);
                        }
                        String fillContent = lineContent.get(j);
                        if (fillContent == null) continue;
                        fillCell.setCellValue(fillContent);
                    }
                }
            }
            IOUtils.closeQuietly(fis);
            target.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void loadConfig() {
        String workingDir = System.getProperty("user.dir") + SEPERATOR;
        System.out.println("working dir is " + workingDir);
        File configFile = new File(workingDir  + CONFIG);
        System.out.println("working file is " + configFile);
        try {
            Scanner scanner = new Scanner(configFile);
            int lineNum = 0;
            while (scanner.hasNext()) {
                String config = scanner.nextLine();
                lineNum++;
                if (config.isEmpty() || config.startsWith("#")) continue;
                String[] pair = config.split("=");
                if (pair.length != 2) {
                    System.out.println("invalid config at line#" + lineNum);
                    continue;
                }
                String key = pair[0].trim();
                String value = pair[1].trim();
                if (value.isEmpty()) {
                    System.out.println("invalid config at line#" + lineNum);
                    continue;
                }
                if (key.equals("src")) {
                    mSourceFile = workingDir + value;
                    System.out.println(mSourceFile);
                } else if (key.equals("target")) {
                    mTargetFile = workingDir + value;
                } else if (key.equals("key")) {
                    mKeyCol = value.charAt(0) - 'A';
                } else if (key.equals("refCol")) {
                    mRefCol = value.charAt(0) - 'A';
                } else if (key.equals("refValue")) {
                    mRefValue = value;
                } else if (key.equals("fillStart")) {
                    mFillFromCol = value.charAt(0) - 'A';
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
        new ExcelSync().syncContents();
    }
}
