package myPractice;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

public class WorkBookFactoryDemo {

    public static void main(String[] args) throws IOException {

        // File should already exist. Else we get File Not Found Exception
        // Both new File("filepath") or new FileInputStream("filepath") works
        //FileInputStream takes more memory as uses buffer

        Workbook wb=WorkbookFactory.create(new File("C:\\Users\\sumit\\IdeaProjects\\Resurrection_ApachePoi\\src\\main\\resources\\blankfile.xlsx"));
        Sheet sh1=wb.getSheet("Sheet1");
        System.out.println(sh1.getLastRowNum());//-1 as no rows else 0 indexed number
        wb.close();
    }
}
