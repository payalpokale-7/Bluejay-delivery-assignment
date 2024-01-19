import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class EmployeeAnalyzer {

    public static void main(String[] args) {
        String filePath = "path/to/your/file.xlsx";  // Update this with the actual path to your file
        analyzeEmployeeData(filePath);
    }

    public static void analyzeEmployeeData(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Assuming the data is in the first sheet, adjust if needed
            Sheet sheet = workbook.getSheetAt(0);

            // Iterator to traverse through rows
            Iterator<Row> rowIterator = sheet.iterator();

            // Skip header row
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            // Variables to track consecutive days and time between shifts
            int consecutiveDays = 0;
            Cell prevDateCell = null;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Assuming Date is in the first column and Hours in the third column, adjust if needed
                Cell dateCell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell hoursCell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                // Convert date cell to Date format if it's not blank
                if (dateCell.getCellType() == CellType.NUMERIC) {
                    double dateValue = dateCell.getNumericCellValue();
                    dateCell.setCellValue(DateUtil.getJavaDate(dateValue));
                }

                // Analyze data
                if (prevDateCell != null) {
                    // Check for consecutive days
                    if (dateCell.getDateCellValue().getTime() - prevDateCell.getDateCellValue().getTime() == 24 * 60 * 60 * 1000) {
                        consecutiveDays++;
                    } else {
                        consecutiveDays = 0;
                    }

                    // Check for less than 10 hours between shifts but greater than 1 hour
                    double timeBetweenShifts = (dateCell.getDateCellValue().getTime() - prevDateCell.getDateCellValue().getTime()) / (60.0 * 60 * 1000);
                    if (1 < timeBetweenShifts && timeBetweenShifts < 10) {
                        System.out.println(row.getCell(1) + " has less than 10 hours between shifts on " + dateCell.getDateCellValue());
                    }
                }

                // Check for more than 14 hours in a single shift
                if (hoursCell.getCellType() == CellType.NUMERIC && hoursCell.getNumericCellValue() > 14) {
                    System.out.println(row.getCell(1) + " worked more than 14 hours on " + dateCell.getDateCellValue());
                }

                // Reset consecutive days counter if not consecutive
                if (consecutiveDays == 7) {
                    System.out.println(row.getCell(1) + " worked 7 consecutive days starting from " + prevDateCell.getDateCellValue());
                    consecutiveDays = 0;
                }

                prevDateCell = dateCell;
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
