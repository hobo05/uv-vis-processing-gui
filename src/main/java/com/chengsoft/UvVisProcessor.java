package com.chengsoft;

import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;
import com.google.common.collect.Sets;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.PathMatcher;
import java.nio.file.Paths;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

/**
 * Created by Tim on 2/18/2016.
 */
public class UvVisProcessor {

    private static PathMatcher XLS_MATCHER = FileSystems.getDefault().getPathMatcher("glob:**.xls");
    private static PathMatcher XLSX_MATCHER = FileSystems.getDefault().getPathMatcher("glob:**.xlsx");

    private static Pattern ALPHA_PATTERN = Pattern.compile("^\\s+\\w$");
    private static Pattern NUMBER_PATTERN = Pattern.compile("^\\d+(\\.\\d+)?$");

    public static void processAndWriteExcel(String inputExcel, String outputExcel) throws IOException, InvalidFormatException {
        Path path = Paths.get(inputExcel);
        Workbook inputWorkbook = null;
        if (XLS_MATCHER.matches(path)) {
            inputWorkbook = new HSSFWorkbook(new POIFSFileSystem(path.toFile()));
        } else if (XLSX_MATCHER.matches(path)) {
            inputWorkbook = new XSSFWorkbook(path.toFile());
        }

        Workbook outputWorkbook = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream(outputExcel);

        processAndWriteSheet(inputWorkbook, outputWorkbook, ".*Wavelength: (\\d+) nm$", "Pho_Scanning1", "Wavelength");
        processAndWriteSheet(inputWorkbook, outputWorkbook, ".*Em: (\\d+) nm$", "Fl_Scanning1", "Emission");

        outputWorkbook.write(fileOut);
        fileOut.close();
    }

    private static void processAndWriteSheet(
            Workbook inputWorkbook,
            Workbook outputWorkbook,
            String headerRegex,
            String sheetName,
            String headerName) throws IOException, InvalidFormatException {

        Pattern headerPattern = Pattern.compile(headerRegex);

        int curWavelength = 0;
        int maxListLength = 0;
        Set<Integer> waveLengthSet = Sets.newHashSet();
        Multimap<Integer, Double> absorbanceMap = ArrayListMultimap.create();
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        int curRow = 0;
        while (curRow <= sheet.getLastRowNum()) {
            curRow++;
            Row row = sheet.getRow(curRow);
            if (row == null) {
                continue;
            }

            // Don't process if the first cell type is blank
            Cell firstCell = row.getCell(row.getFirstCellNum());
            int firstCellType = firstCell.getCellType();
            if (Cell.CELL_TYPE_BLANK == firstCellType) {
                continue;
            }

            // Retrieve first cell's value
            String firstCellValue = firstCell.getStringCellValue();
            Matcher matcher = headerPattern.matcher(firstCellValue);
            if (matcher.matches()) {
                curWavelength = Integer.valueOf(matcher.group(1));
                // Add to the set
                waveLengthSet.add(curWavelength);
            }

            // Conditions:
            //  - Make sure we have found the first wavelength at least
            //  - Find the row where the values start (e.g. A, B, C, etc...)
            //  - The first absorbance must exist
            String well_2_absorbance = getStringValue(row.getCell(2, Row.CREATE_NULL_AS_BLANK));
            String well_3_absorbance = getStringValue(row.getCell(3, Row.CREATE_NULL_AS_BLANK));
            if (waveLengthSet.contains(curWavelength)
                    && ALPHA_PATTERN.matcher(firstCellValue).matches()
                    && NUMBER_PATTERN.matcher(well_2_absorbance).matches()) {

                // Add well_2_absorbance to the list since we know it exists
                absorbanceMap.put(curWavelength, Double.valueOf(well_2_absorbance));

                // Verify that well_3_absorbance exists before adding it
                if (NUMBER_PATTERN.matcher(well_3_absorbance).matches())
                    absorbanceMap.put(curWavelength, Double.valueOf(well_3_absorbance));

                // Determine maxListLength
                maxListLength = Math.max(maxListLength, absorbanceMap.get(curWavelength).size());
            }
        }


        Sheet outputSheet = outputWorkbook.createSheet(sheet.getSheetName());

        Row headerRow = outputSheet.createRow(0);
        headerRow.createCell(0).setCellValue(headerName);
        IntStream.range(1, maxListLength+1).boxed()
                .forEach(i -> headerRow.createCell(i).setCellValue(i));

        new TreeMap<>(absorbanceMap.asMap()).entrySet().stream()
                .forEach(entry -> {
                    Row newRow = outputSheet.createRow(outputSheet.getLastRowNum()+1);
                    newRow.createCell(0).setCellValue(entry.getKey());
                    entry.getValue().stream()
                            .forEach(s -> {
                                int newCellNum = newRow.getLastCellNum();
                                Cell newCell = newRow.createCell(newCellNum);
                                newCell.setCellValue(s);
                            });
                });
    }

    private static String getStringValue(Cell cell) {
        String value = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
                value = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
        }
        return value;
    }
}
