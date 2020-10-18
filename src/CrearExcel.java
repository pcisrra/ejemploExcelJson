import java.io.File;
import java.io.*;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import java.util.Iterator;

public class CrearExcel
{
	public static void main(String args[])
    {
		JSONParser pars = new JSONParser();
		Object ob1 = null;
        JSONArray jsvec = null;
        int i=1;
		try
		{
			XSSFWorkbook hojaEx = new XSSFWorkbook();   
			XSSFSheet hoja = hojaEx.createSheet("Hoja1");
          
			XSSFRow fil = hoja.createRow(0);
			Cell colum = fil.createCell(0);
			Cell colum1 = fil.createCell(1);
			Cell colum2 = fil.createCell(2);
			Cell colum3 = fil.createCell(3);
			colum.setCellValue("Numero");
			colum1.setCellValue("Palabra");
			colum2.setCellValue("Tipo");
			colum3.setCellValue("Significado");
			
			ob1 = pars.parse(new FileReader("D:/Projects/Palabras.json"));
			jsvec = (JSONArray) ob1;
			Iterator<JSONObject> rep = jsvec.iterator();
            while (rep.hasNext())
            {
            	JSONObject jos = (JSONObject) rep.next();
            	
            	fil = hoja.createRow(i);
            	colum = fil.createCell(0);
            	colum1 = fil.createCell(1);
            	colum2 = fil.createCell(2);
            	colum3 = fil.createCell(3);
            	colum.setCellValue(String.valueOf(jos.get("numero")));
            	colum1.setCellValue(String.valueOf(jos.get("Palabra")));
            	colum2.setCellValue(String.valueOf(jos.get("Tipo")));
            	colum3.setCellValue(String.valueOf(jos.get("Significado")));
            	i++;
            }
			
			FileOutputStream out = new FileOutputStream(new File("D:/Projects/ejemploTransformado.xlsx"));      
			hojaEx.write(out);
			out.close();
			System.out.println("archivo creado");
		}
		catch(IOException e)
		{
			System.out.println("Error creando el archivo excel");  
		}
		catch(ParseException pas)
		{
			System.out.println(pas.getMessage());
		}
    } 
}