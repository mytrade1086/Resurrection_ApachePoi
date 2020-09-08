package myPractice;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class WritingToXls {

    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook();
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream("C:\\Users\\sumit\\IdeaProjects\\Resurrection_ApachePoi\\src\\main\\resources\\myFile.xlsx");
            wb.write(fos);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        } finally {
            //fos.close();
            wb.close();
        }


        FileInputStream fis = new FileInputStream("C:\\Users\\sumit\\IdeaProjects\\Resurrection_ApachePoi\\src\\main\\resources\\myFile.xlsx");
        Workbook wb2 = WorkbookFactory.create(fis);
//        fis.close();

        Row row = wb2.createSheet("sumit").createRow(0);
        Cell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("new Line");

        Cell cell2 = row.createCell(1, CellType.NUMERIC);
        cell2.setCellValue(14);

        Cell cell3 = row.createCell(3, CellType.NUMERIC);
        cell2.setCellValue(144.20);

       // Cell cell4=row.createCell(4).setCellValue(new Date());
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss a");
        cell.setCellValue(simpleDateFormat.format(new Date()));


        Cell cell4 = row.createCell(4);
        cell4.setCellValue(simpleDateFormat.format(new Date()));





        wb2.write(new FileOutputStream("C:\\Users\\sumit\\IdeaProjects\\Resurrection_ApachePoi\\src\\main\\resources\\myFile.xlsx"));
        wb2.close();
    }


}



