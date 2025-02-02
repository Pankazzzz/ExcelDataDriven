package ExcelDriven.ExcelDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.time.Duration;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class UploadDownload {

    static DataFormatter formatter = new DataFormatter();

    public static void main(String[] args) throws Exception {
        String fileName = "C:\\Users\\dell\\Downloads\\download.xlsx";
        WebDriver driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
        driver.manage().window().maximize();

        driver.findElement(By.id("downloadButton")).click();

        // Wait for file to be fully downloaded
        Thread.sleep(3000);

        // Update Excel file
        String updatedPrice = "200";
        int rowNumber = getRowNumber(fileName, "Papaya");
        int columnNumber = getColumnColumnNumber(fileName, "price");
        boolean value = updateData(fileName, rowNumber, columnNumber, updatedPrice);
        System.out.println("Excel Update Status: " + value);

        driver.findElement(By.xpath("//input[@id='fileinput']")).sendKeys(fileName);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        By toastLocator = By.cssSelector(".Toastify__toast-body div:nth-child(2)");
        wait.until(ExpectedConditions.visibilityOfElementLocated(toastLocator));

        String toasterMessage = driver.findElement(toastLocator).getText();
        System.out.println(toasterMessage);
        Assert.assertEquals("Updated Excel Data Successfully.", toasterMessage);
        wait.until(ExpectedConditions.invisibilityOfElementLocated(toastLocator));

        // Verify updated price
        String fruitName = "Papaya";
        String priceElement = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");
        String price = driver.findElement(By.xpath("//div[text()='" + fruitName + "']/parent::div/parent::div/div[@id='cell-" + priceElement + "-undefined']")).getText();
        System.out.println("Extracted Price: " + price);
        Assert.assertEquals(price, updatedPrice);

        // Ensure file is not locked before deleting
        System.gc(); 
        Thread.sleep(500);

        // Delete the file
        File file = new File(fileName);
        if (file.exists() && file.canWrite()) {
            boolean deleted = file.delete();
            System.out.println("Deleted: " + deleted);
        } else {
            System.out.println("File could not be deleted.");
        }

        driver.close();
    }

    public static int getRowNumber(String filename, String fruitName) throws IOException {
        try (FileInputStream fis = new FileInputStream(filename);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> totalRows = sheet.rowIterator();
            int rowNumbers = sheet.getLastRowNum();
            Row row = sheet.getRow(0);
            Iterator<Cell> cells = row.cellIterator();

            int fruitColumnIndex = -1;
            int k = 0;
            while (cells.hasNext()) {
                Cell c = cells.next();
                if (c.getStringCellValue().equalsIgnoreCase("fruit_name")) {
                    fruitColumnIndex = k;
                }
                k++;
            }

            int rowCount = 0;
            while (totalRows.hasNext()) {
                Row r = totalRows.next();
                if (r.getCell(fruitColumnIndex).getStringCellValue().equalsIgnoreCase(fruitName)) {
                    return rowCount;
                }
                rowCount++;
            }
            return 0;
        }
    }

    public static int getColumnColumnNumber(String filename, String priceColumnName) throws Exception {
        try (FileInputStream fis = new FileInputStream(filename);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            Iterator<Cell> cells = row.cellIterator();

            int priceColumnIndex = -1;
            int k = 0;
            while (cells.hasNext()) {
                Cell c = cells.next();
                if (c.getStringCellValue().equalsIgnoreCase(priceColumnName)) {
                    priceColumnIndex = k;
                }
                k++;
            }
            return priceColumnIndex;
        }
    }

    public static boolean updateData(String filename, int row, int column, String value) throws Exception {
        try (FileInputStream fis = new FileInputStream(filename);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(filename)) {

            XSSFSheet sheet = workbook.getSheet("Sheet1");
            Row rowField = sheet.getRow(row);
            Cell cellField = rowField.getCell(column);
            cellField.setCellValue(value);

            workbook.write(fos);
            return true;
        }
    }
}
