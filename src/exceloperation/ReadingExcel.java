package exceloperation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import com.google.gson.stream.JsonReader;
import com.ibm.cloud.sdk.core.http.Response;
import java.io.IOException;
import java.util.*;
import com.ibm.watson.common.WatsonHttpHeaders;
import com.ibm.watson.language_translator.v3.util.Language;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.ibm.cloud.sdk.core.security.Authenticator;
import com.ibm.cloud.sdk.core.security.IamAuthenticator;
import com.ibm.watson.language_translator.v3.LanguageTranslator;
import com.ibm.watson.language_translator.v3.model.TranslateOptions;
import com.ibm.watson.language_translator.v3.model.TranslationResult;
import com.ibm.watson.natural_language_classifier.v1.NaturalLanguageClassifier;
import com.ibm.watson.natural_language_classifier.v1.model.Classification;
import com.ibm.watson.natural_language_classifier.v1.model.ClassifyOptions;
import javax.json.*;
import java.io.StringReader;

public class ReadingExcel 
{

	public static void main(String[] args) throws IOException
	{
		String excelFilePath=".\\datafiles\\abc.xlsx";								//Store the File Path
		
		FileInputStream inputstream=new FileInputStream(excelFilePath);				//Open the file in READ MODE
		XSSFWorkbook workbook=new XSSFWorkbook(inputstream);						//Fetches the workbook from the given file
		XSSFSheet sheet=workbook.getSheet("Sheet1");								//Fetches the given sheet number from the workbook
		
		int rows=sheet.getLastRowNum();												//Fetches the last row num from the sheet to iterate rows
		int cols=sheet.getRow(1).getLastCellNum();									//Fetches the last cell num from the sheet to iterate cell
		
		List<String> l1=new ArrayList<String>();									//Create List of String to store the data
		
		for(int r=1;r<=rows;r++)													
		{
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell=row.getCell(c);										//Create an object for specified cell
				String cvalue=cell.getStringCellValue();							//Stores the cell data 
				String object=LanguageTranslator(cvalue);							//Function call 
				//ClassifyLanguage(cvalue);
				System.out.println(object);											//Print the return value of the function 
				l1.add(object);														//Add the return value to List
			}
		}
		writeFile(l1);																//Function Call
		inputstream.close();														//To close the excel file
	}
			
		public static String getTranslatedString(String jsonString) 				//In this function we tokenize the translated data and fetches the required data
		{
			javax.json.JsonReader reader = javax.json.Json.createReader(new StringReader(jsonString));
			JsonObject jsonObject = reader.readObject();
			String searchResult = jsonObject.getJsonArray("translations").getJsonObject(0).getString("translation");
			return searchResult;																					//return call to main function
		}
		
		public static String LanguageTranslator(String cvalue)
		{
			Authenticator authenticator = new IamAuthenticator("yIapuW7SnDhA8Lf9EDWwAt5RHBSoemaMspVAYUUbzvuQ");										//Using API key we Authenticate their service
			LanguageTranslator service = new LanguageTranslator("2018-05-01", authenticator);
			service.setServiceUrl("https://api.eu-gb.language-translator.watson.cloud.ibm.com/instances/16e7498b-f09b-4b06-b7eb-d351c92f319b");		//pass the url of API
			Map<String, String> headers = new HashMap<>();																					
		    headers.put(WatsonHttpHeaders.X_WATSON_TEST, "1");
		    service.setDefaultHeaders(headers);																										//Set the Header
		    TranslateOptions translateOptions =
		            new TranslateOptions.Builder()
		                .addText(cvalue)																											//Actual text to be translated
		                .target(Language.ENGLISH)																									//Target Language
		                .build();
			
			Response<TranslationResult> translationResult = service.translate(translateOptions).execute();											//Fetches the translated data from their API and stores it in form of Result
			String s=getTranslatedString(translationResult.getResult().toString());																	//Call to convert a translation result data to string
			//System.out.println(s);
			return s;
		}
		
		
		/*public static void ClassifyLanguage(String cvalue)
		{
			Authenticator authenticator = new IamAuthenticator("yIapuW7SnDhA8Lf9EDWwAt5RHBSoemaMspVAYUUbzvuQ");
			NaturalLanguageClassifier service = new NaturalLanguageClassifier(authenticator);
			service.setServiceUrl("https://api.eu-gb.language-translator.watson.cloud.ibm.com/instances/16e7498b-f09b-4b06-b7eb-d351c92f319b");
			Map<String, String> headers = new HashMap<>();
		    headers.put(WatsonHttpHeaders.X_WATSON_TEST, "1");
		    service.setDefaultHeaders(headers);
			
			ClassifyOptions classifyOptions =new ClassifyOptions.Builder().classifierId("<classifierId>").text(cvalue).build();
		    Classification classification = service.classify(classifyOptions).execute().getResult();

			System.out.println(classification);
		}*/	
		
		public static void writeFile(List<String> l1)
		{
			try
			{
				String excelFilePath=".\\datafiles\\abc.xlsx";					//we open file again to write the translated data
				
				FileInputStream inputstream=new FileInputStream(excelFilePath);
				XSSFWorkbook workbook=new XSSFWorkbook(inputstream);
				XSSFSheet sheet=workbook.getSheet("Sheet1");
				
				Iterator <String> iter=l1.iterator();
				int rownum=1;
				int cellnum=2;											//cell number is initialised to 2 throughout as we have to write the translated data in cell 2 of each row
				
				while(iter.hasNext())
				{						
					String temp=(String) iter.next();					
					Cell cell=sheet.getRow(rownum).getCell(cellnum);	//fetch the cell to write the translated data
					cell.setCellValue(temp);							//set the cell value
					rownum++;
				}
				inputstream.close();									//close the file
				
				FileOutputStream outputstream=new FileOutputStream(excelFilePath); //Open the file in WRITE MODE
				workbook.write(outputstream);										//Write the data into workbook
				outputstream.close();												//Close the file								
			}
			catch(FileNotFoundException e)
			{
				e.printStackTrace();
			}
			catch(IOException e)
			{
				e.printStackTrace();
			}
		}
}
