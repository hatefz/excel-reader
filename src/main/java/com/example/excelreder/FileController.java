package com.example.excelreder;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@RestController
public class FileController {

    @PostMapping(value = "/file", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<Boolean> sendSmsBulkWithFile(@RequestParam(value = "file") MultipartFile file) throws IOException, InvalidFormatException {

//        readXLSFile(file);
        System.out.println("------------------------------");
        readXLSFileWithBlankCells(file);

//        XSSFWorkbook wb = XSSFWorkbookFactory.createWorkbook(OPCPackage.open(file.getInputStream()));
//        XSSFSheet sheet = wb.getSheetAt(0);
//
//        int totalElements = sheet.getPhysicalNumberOfRows();
//        int lastRow = sheet.getLastRowNum();
//        int total = 0;
//        for (Row row : sheet) {
//            Iterator<Cell> cellIterator = row.cellIterator();
//
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                total++;
//                switch (cell.getCellType()) {
//                    case NUMERIC:
//                        System.out.print(cell.getNumericCellValue() + " ,");
//                        break;
//                    case _NONE:
//                        System.out.print(" NULL,");
//                        break;
//                    case BLANK:
//                        System.out.print(" BLANK,");
//                        break;
//                    case STRING:
//                        System.out.print(cell.getStringCellValue() + " ,");
//                        break;
//
//                }
//            }
//            System.out.println();
//        }
//        System.out.println(totalElements + " " + total + " " + lastRow);
        return new ResponseEntity<>(true, HttpStatus.OK);
    }

    public static void readXLSFile(MultipartFile file) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook(file.getInputStream());
            XSSFSheet sheet = wb.getSheetAt(0);

            XSSFRow row;
            XSSFCell cell;

            Iterator rows = sheet.rowIterator();

            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                Iterator cells = row.cellIterator();

                while (cells.hasNext()) {
                    cell = (XSSFCell) cells.next();
                    System.out.print(cell.toString() + " ");
                }
                System.out.println();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void readXLSFileWithBlankCells(MultipartFile file) {
        try {
            List<String[]> dataLines = new ArrayList<>();

            XSSFWorkbook wb = new XSSFWorkbook(file.getInputStream());

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            Iterator<Row> rows = sheet.rowIterator();
            String[] data;

            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                int numCells = row.getLastCellNum();
                data = new String[numCells];
                for (int i = 0; i < numCells; i++) {
                    cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if (cell.getCellType() == CellType.BLANK) {
                        data[i] = "";
                    } else if (cell.getCellType() == CellType.NUMERIC) {
                        data[i] = String.valueOf(cell.getNumericCellValue());
                    } else if (cell.getCellType() == CellType.STRING) {
                        data[i] = cell.getStringCellValue();
                    }
                }
                dataLines.add(data);
            }

            File csvOutputFile = new File("test-csv.csv");
            try (PrintWriter pw = new PrintWriter(csvOutputFile)) {
                dataLines.stream()
                        .map(this::convertToCSV)
                        .forEach(pw::println);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String convertToCSV(String[] data) {
        return String.join(",", data);
    }
}
