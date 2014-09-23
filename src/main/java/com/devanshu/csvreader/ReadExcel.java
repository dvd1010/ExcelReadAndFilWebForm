package com.devanshu.csvreader;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;
import java.util.*;


import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.HttpVersion;
import org.apache.http.NameValuePair;
import org.apache.http.client.HttpClient;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.params.CoreProtocolPNames;
import org.apache.http.protocol.HTTP;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ReadExcel {

	public static final String URL = "";
	public static final String FILE_PATH = "";
	public static Map<Integer, String> formMap;
	static{
		formMap = new TreeMap<Integer, String>();
		formMap.put(0, "registrant.givenName");
		formMap.put(1, "registrant.surname");
		formMap.put(2, "registrant.email");
		formMap.put(3, "registrant.jobTitle");
		formMap.put(4, "registrant.organization");
		formMap.put(5, "registrant.phone");
	}
	
	public static void main(String[] args){
		ReadExcel readCsv = new ReadExcel();
		try {
			readCsv.read();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void read() throws Exception{
		HttpClient httpClient = new DefaultHttpClient();
		httpClient.getParams().setParameter(CoreProtocolPNames.PROTOCOL_VERSION,HttpVersion.HTTP_1_1);
		HttpPost post = new HttpPost(URL);
		try{ 

			Map<Integer, String> headerMap = new TreeMap<Integer, String>();
			File file = new File(FILE_PATH);
			FileInputStream fileInputStream = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
			HSSFSheet sheet = workbook.getSheetAt(0);
			if (sheet == null) {
	            throw new Exception("No Worksheet Found");
	        }
			java.util.Iterator<Row> iterator = sheet.iterator();
	        Row headerRow = iterator.next();
	        if (headerRow == null) {
	            throw new Exception("Worksheet does not contain Header in first Row.");
	        }
	        for(Cell cell : headerRow){
	        	headerMap.put(cell.getColumnIndex(), cell.getStringCellValue());
	        }
	        
	        while(iterator.hasNext()){
	        	List <NameValuePair> nameValuePairs = new ArrayList <NameValuePair>();  
	        	Row row = iterator.next();
	        	for(Cell cell : row){
	        		if(formMap.get(cell.getColumnIndex()) != null)
	        			nameValuePairs.add(new BasicNameValuePair(formMap.get(cell.getColumnIndex()), cell.getStringCellValue())); 
	        	}
	        	if(!nameValuePairs.isEmpty())
		        	post.setEntity(new UrlEncodedFormEntity(nameValuePairs, HTTP.UTF_8));
		        	HttpResponse response = httpClient.execute(post);  
		        	HttpEntity enty = response.getEntity();
		            if (enty != null)
		                enty.consumeContent();
	        }

        	
		}catch(IOException e){
			e.printStackTrace();
		}
		finally {
			httpClient.getConnectionManager().shutdown();
	    }
	}
}
