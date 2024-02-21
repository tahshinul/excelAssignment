import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.Public;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.Flow.Publisher;

@SuppressWarnings("unused")
public class assigned {
    public static void main(String[] args) throws EncryptedDocumentException, IOException, InterruptedException {
    	WebDriver driver = new ChromeDriver();
 		driver.manage().window().maximize();
 		driver.get("https://google.com/");
 		
    	 String keywords[] = getKeywords();
    	 List<String> shortestList = new ArrayList<>();
	     List<String> longestList = new ArrayList<>();
    	 for (int i = 0; i < keywords.length; i++) {
			driver.findElement(By.id("APjFqb")).sendKeys(keywords[i]);
			Thread.sleep(1000);
			List<WebElement> elements = driver.findElements(By.xpath("//div[contains(@class,'lnnVSe')] //div[contains(@class,'wM6W7d')] //span"));
	        List<String> elementTexts = new ArrayList<>();
	        for (WebElement element : elements) {
	            elementTexts.add(element.getText());
	        }
	        String[] x = elementTexts.toArray(new String[0]);
	        String[] y = removeBlankStrings(x); 
	        shortestList.add(shortestString(y));
	        longestList.add(longestString(y));
			driver.findElement(By.id("APjFqb")).clear();
			Thread.sleep(1100);
		}
    	 writeToExcel(shortestList,3);
    	 writeToExcel(longestList,4);
    	 
    	 
    	 driver.quit();
    }
    
    public static String[] getKeywords() throws EncryptedDocumentException, IOException {
        FileInputStream file = new FileInputStream("C:\\Users\\tahsh\\Downloads\\Excel.xlsx");
        Workbook workbook = WorkbookFactory.create(file);
        String dayOfWeekString = day();
        Sheet sheet = workbook.getSheet(dayOfWeekString);
        int columnIndex = 2;
        List<String> columnValues = new ArrayList<>();
        for (Row row : sheet) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                columnValues.add(cell.getStringCellValue());
            }
        }

        String[] columnArray = columnValues.toArray(new String[0]);
        workbook.close();
        file.close();

        
        
        return columnArray;
    	
    	
    }
    

    public static String shortestString(String[] strings) {
        if (strings == null || strings.length == 0) {
            return null;
        }
        String shortest = strings[0];
        for (int i = 1; i < strings.length; i++) {
            if (strings[i].length() < shortest.length()) {
                shortest = strings[i];
            }
        }

        return shortest;
    }
    
    public static String[] removeBlankStrings(String[] array) {
        List<String> list = new ArrayList<>();
        for (String str : array) {
            if (str != null && !str.trim().isEmpty()) {
                list.add(str);
            }
        }

        String[] result = new String[list.size()];
        return list.toArray(result);
    }
    
    
    public static String longestString(String[] strings) {
        if (strings == null || strings.length == 0) {
            return null;
        }
        String longest = strings[0];
        for (int i = 1; i < strings.length; i++) {
            if (strings[i].length() > longest.length()) {
                longest = strings[i];
            }
        }
        return longest;
    }
    
    
    public static void writeToExcel(List<String> strings,int x) {
    	
    	
        String dayOfWeek = day();
    	
    	String filePath = "C:\\Users\\tahsh\\Downloads\\Excel.xlsx"; // Path to the Excel file
        String sheetName = dayOfWeek; // Name of the sheet
        int columnIndex = x; // Index of the column
        int startRow = 2; // Index of the row
    
        try (FileInputStream fileIn = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fileIn)) {

               Sheet sheet = workbook.getSheet(sheetName);
               if (sheet == null) {
                   sheet = workbook.createSheet(sheetName);
               }

               for (int i = 0; i < strings.size(); i++) {
                   Row row = sheet.getRow(startRow + i);
                   if (row == null) {
                       row = sheet.createRow(startRow + i);
                   }
                   Cell cell = row.createCell(columnIndex);
                   cell.setCellValue(strings.get(i));
               }

               // Write the workbook content back to the file
               try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                   workbook.write(fileOut);
                   System.out.println("Excel file has been updated successfully.");
               }

           } catch (IOException e) {
               e.printStackTrace();
           }
    
    
    
    
    
    
    
    
    
    }
    
    
    
    // Get the day of the week
    public static String day() {
    	LocalDate currentDate = LocalDate.now();
        DayOfWeek dayOfWeek = currentDate.getDayOfWeek();
        String dayOfWeekString = dayOfWeek.getDisplayName(TextStyle.FULL_STANDALONE, Locale.ENGLISH);
        return dayOfWeekString;
    
    }
    
    
    
}
