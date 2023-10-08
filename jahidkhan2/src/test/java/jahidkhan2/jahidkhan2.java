package jahidkhan2;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;

import java.io.File;

import java.util.Iterator;


public class jahidkhan2 {

	public static void main(String[] args) {
		
					System.setProperty("webdriver.gecko.driver", "E:\\geckodriver.exe");

			        
			        WebDriver driver = new FirefoxDriver();

			        try {
			    		String excelFilePath1 = "F:\\autosearch\\Excel.xlsx";
			            String sheetName1 = "Friday";
			            int columnIndex1 = 0;  
			            
			            String[] columnData1 = readColumnData(excelFilePath1, sheetName1, columnIndex1);
			        	
			        	
			            FileInputStream fis = new FileInputStream(excelFilePath1);
			            XSSFWorkbook wb=new XSSFWorkbook(fis);
			            XSSFSheet sheet = wb.getSheet(sheetName1);
			            
			            driver.get("https://www.google.com/");
			            
			            XSSFRow row = null;
			            XSSFCell cell1 = null;
			            XSSFCell cell2 = null;
			            XSSFCell cell3 = null;
			           

			            
			           int i;
		
			           for (i=1;i<=sheet.getLastRowNum();i++) {
			        	   row = sheet.createRow(i);
			            WebElement searchBox = driver.findElement(By.xpath("//textarea[@id='APjFqb']"));
			            searchBox.sendKeys(columnData1[i]);

			            
			            Thread.sleep(3000);

			           
			            List<WebElement> suggestions = driver.findElements(By.xpath("//ul[@role='listbox']//li"));

			            
			            List<String> suggestionTexts = new ArrayList<>();
			            for (WebElement suggestion : suggestions) {
			                suggestionTexts.add(suggestion.getText());
			                
			            }

			            
			              String shortestSuggestion = suggestionTexts.stream()
			                    .min((s1, s2) -> s1.length() - s2.length())
			                    .orElse("No suggestions found");

			            String longestSuggestion = suggestionTexts.stream()
			                    .max((s1, s2) -> s1.length() - s2.length())
			                    .orElse("No suggestions found");
			           

	                    
			            System.out.println("Shortest Suggestion: " + shortestSuggestion);
			            System.out.println("Longest Suggestion: " + longestSuggestion);
			            System.out.println("\n");
			            
			            
			            cell1 = row.createCell(1);
			            cell2 = row.createCell(2);
			            cell3 = row.createCell(0);
			            
			            
			            cell1.setCellType(CellType.STRING);
			            cell2.setCellType(CellType.STRING);
			            cell3.setCellType(CellType.STRING);
			            
			            cell1.setCellValue(shortestSuggestion);
			            cell2.setCellValue(longestSuggestion);
			            cell3.setCellValue(columnData1[i]);
			            searchBox.clear();
			            
			            
                        Thread.sleep(1000);
			        }
			           FileOutputStream fos  = new FileOutputStream(excelFilePath1);
			           wb.write(fos);
			           fos.close();
			           //wb.close();
			        }
			        catch (Exception e) {
			            e.printStackTrace();
			        } finally {
			            
			            driver.quit();
			        }
			        }
				
				

		
				    public static String[] readColumnData(String excelFilePath, String sheetName, int columnIndex) {
				        List<String> columnDataList = new ArrayList<>();

				        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
				             Workbook workbook = new XSSFWorkbook(fis)) {

				            Sheet sheet = workbook.getSheet(sheetName);

				            if (sheet != null) {
				                Iterator<Row> iterator = sheet.iterator();

				                while (iterator.hasNext()) {
				                    Row currentRow = iterator.next();
				                    Cell cell = currentRow.getCell(columnIndex);

				                    if (cell != null) {
				                        
				                        String cellValue = cell.getStringCellValue();
				                        columnDataList.add(cellValue);
				                    }
				                }
				            }
				      
				        } catch (Exception e) {
				            e.printStackTrace();
				        }

				        
				        return columnDataList.toArray(new String[0]);

					}

	}

	

