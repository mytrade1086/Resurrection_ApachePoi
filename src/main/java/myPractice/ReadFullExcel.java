package myPractice;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadFullExcel {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream("C:\\Users\\sumit\\IdeaProjects\\Resurrection_ApachePoi\\src\\main\\resources\\withDATA.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        Workbook wb = WorkbookFactory.create(fis);
        Iterator<Row> ri = wb.getSheetAt(0).rowIterator();
        while (ri.hasNext()) {
            Row row = ri.next();
            int lastCell = row.getLastCellNum();
            System.out.println("Last Cell Number is " + lastCell);
            for (int i = 0; i < lastCell; i++) {

                switch (row.getCell(i).getCellType()) {
                    case STRING:
                        System.out.print(row.getCell(i).getRichStringCellValue().getString()+"\t");
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(row.getCell(i))) {
                            System.out.print(row.getCell(i).getDateCellValue()+"\t");
                        } else {
                            System.out.print(row.getCell(i).getNumericCellValue()+"\t");
                        }
                        break;
                    case BOOLEAN:
                        System.out.print(row.getCell(i).getBooleanCellValue()+"\t");
                        break;
                    case FORMULA:
                        System.out.print(row.getCell(i).getCellFormula()+"\t");
                        break;
                    case BLANK:
                        System.out.println();
                        break;
                    default:
                        System.out.println();
                }
            }
            System.out.println();
        }
        wb.close();
    }
}
