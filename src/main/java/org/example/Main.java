package org.example;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.*;

public class Main {


    public static void main(String[] args)  {

        String fileLocation = "/Users/ryum/IdeaProjects/excelparser/src/main/resources/Обновление справочников.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(fileLocation);
            Workbook workbook = new XSSFWorkbook(inputStream);

            for (int j = 0; j < workbook.spliterator().estimateSize(); j++) {
                System.out.print("------------------------------------------------------------------------------------------------ \n" );
                Sheet sheet = workbook.getSheetAt(j);
                Iterator<Row> rowIterator = sheet.rowIterator();

                List<String> columns = new ArrayList<>();
                List<String> sqlInsertStatements = new ArrayList<>();

                Row headerRow = rowIterator.next();
                for (Cell headerCell : headerRow) {
                    columns.add(headerCell.getStringCellValue());
                }

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    List<String> values = new ArrayList<>();

                    for (int i = 0; i < columns.size(); i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        switch (cell.getCellType()) {
                            case STRING:
                                if (cell.getStringCellValue().equals("NULL")) {
                                    values.add(null);
                                } else {
                                    values.add("'" + cell.getStringCellValue() + "'");
                                }
                                break;
                            case NUMERIC:

                                if (DateUtil.isCellDateFormatted(cell)) {
                                    values.add("'" + new Timestamp(cell.getDateCellValue().getTime()) + "'");
                                } else {
                                    double numericValue = cell.getNumericCellValue();
                                    if ((int) numericValue == numericValue) {
                                        values.add(String.valueOf((int) numericValue));
                                    } else {
                                        values.add(String.valueOf(numericValue));
                                    }
                                }
                                break;
                            case BOOLEAN:
                                values.add(cell.getBooleanCellValue() ? "TRUE" : "FALSE");
                                break;
                            case BLANK:
                            case _NONE:
                            case ERROR:
                            default:
                                values.add(null);
                                break;
                        }
                    }

                    String valueString = String.join(", ", values);
                    String sqlInsert = "INSERT INTO " + workbook.getSheetName(j) + " (" + String.join(", ", columns) + ") VALUES (" + valueString + ");";
                    sqlInsertStatements.add(sqlInsert);
                }

                workbook.close();

                for (String sql : sqlInsertStatements) {
                    System.out.println(sql);
                }
            }


        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}