package testexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo 
{
    public static void main(String[] args) 
    {
        try
        {
            FileInputStream file = new FileInputStream(new File("/Users/jimmylo/Desktop/testDir/excel/abc.xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            //System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                        	String value = cell.getStringCellValue();
                        	if(value.contains("@")){
                        		System.out.print(cell.getStringCellValue() +",");
                        	}
                            break;
                    }
//                    if(cell.getColumnIndex()== 1){
//                    	cell.setCellValue("test2");
//                    }
                }
                
            }
            
            file.close();
            
//            Map<String, Object[]> data = new TreeMap<String, Object[]>();
//            data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
//            data.put("2", new Object[] {1, "Amittyffvkutfktufutf", "Shukla2222266"});
//            data.put("3", new Object[] {2, "Lokesh", "Gupta33333u"});
//            data.put("4", new Object[] {3, "John", "Adwards"});
//            data.put("5", new Object[] {4, "Brian", "Schultz"});
//            sheet.getRow(0).getCell(1).setCellValue("testi222222");
//            //Iterate over data and write to sheet
//            Set<String> keyset = data.keySet();
//            int rownum = 6;
//            for (String key : keyset)
//            {
//                Row row = sheet.createRow(rownum++);
//                Object [] objArr = data.get(key);
//                int cellnum = 0;
//                for (Object obj : objArr)
//                {
//                   Cell cell = row.createCell(cellnum++);
//                   if(obj instanceof String)
//                        cell.setCellValue((String)obj);
//                    else if(obj instanceof Integer)
//                        cell.setCellValue((Integer)obj);
//                }
//            }
//            
//            try
//            {
//                //Write the workbook in file system
//                FileOutputStream out = new FileOutputStream(new File("/Users/jimmylo/Desktop/testDir/excel/howtodoinjava_demo2.xlsx"));
//                workbook.write(out);
//                out.close();
//                System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
//            } catch (Exception e) 
//            {
//                e.printStackTrace();
//            }
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
}
