package jsonConvertor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import com.google.gson.Gson;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DateUtil;
//import org.json.JSONObject;

import java.io.FileWriter;
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
// import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

//import org.apache.poi.ss.usermodel.CellType;


public class Convertor {
	
	
	    public static void main(String[] args) throws Exception
	    {
	    	Convertor obj1=new Convertor();
	    	obj1.main();
	    }
	public void main()
	    {
	        try
	        {
	        	//making map
	        	
	            FileInputStream file = new FileInputStream(new File("C:\\Users\\kriti\\Downloads\\Invoice.xlsx"));
	 
	            //Create Workbook instance holding reference to .xlsx file
	            XSSFWorkbook workbook = new XSSFWorkbook(file);
	 
	            //Get first/desired sheet from the workbook
	            XSSFSheet sheet = workbook.getSheetAt(0);
	 
	            List<String> keys = new ArrayList<>();
	            //Iterate through each rows one by one
	            Iterator<Row> rowIterator = sheet.iterator();
	            keys = readRow(rowIterator.next().cellIterator());
	            // System.out.println(keys);
	            List<Map<String,String>> Lmap =new ArrayList<>();
	            while (rowIterator.hasNext()) 
	            {
	            	Map <String, String> hm=new HashMap<String, String>();
	                Row row = rowIterator.next();
	                Iterator<Cell> cellIterator = row.cellIterator();
	                List<String> rowValues = readRow(cellIterator);
	                // System.out.println(keys.size()+" :"+rowValues.size());
	                // System.out.println(rowValues);
	                
	                for(int i=0;i<keys.size();i++)
	                {
	                	hm.put(keys.get(i),rowValues.get(i));
	                }
	               
	            Lmap.add(hm);
	            // System.out.println(Lmap);
	            
	            }System.out.println(ListMethod(Lmap));
	            
	            file.close();
	            }
	        catch (Exception e) 
	        {
	            e.printStackTrace();
	        }
	    }
	    public List<String> readRow(Iterator<Cell> cellIterator) {
	    	List<String> list = new ArrayList<>();
	    	// int i =1;
	    	while (cellIterator.hasNext()) 
            {
	    		Cell cell = cellIterator.next();
	    		// System.out.print(i++ + ": ");
                switch (cell.getCellType()) 
                {
                    case NUMERIC:
                    	if(DateUtil.isCellDateFormatted(cell)) {
                    		DataFormatter df = new DataFormatter();
                    		//System.out.println("Adding: " + df.formatCellValue(cell));
                    		list.add(df.formatCellValue(cell));
                    	} else {
                    		//System.out.println("Adding: " + cell.getNumericCellValue());
                    		list.add(Double.toString(cell.getNumericCellValue()));
                    	}
                        break;
                    case STRING:
                    	//System.out.println("+ cell.getStringCellValue());
                    	String s = cell.getStringCellValue(); 
                        list.add(s==null?"":s);
                    	break;
                        
                        default:
                        	System.out.println("No value");
                        	break;
                }
            }
            //System.out.println("");
	    	return list;
        }
	    public String ListMethod(List<Map<String, String>> l)
		{
			// JSONObject json = new JSONObject(l);
			Gson gson = new Gson();
			return gson.toJson(l);
			// return JSONObject.valueToString(json);
			// add/remove/get/contains/size
		}
	    }
	


