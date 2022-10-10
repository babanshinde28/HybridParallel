package com.kdd.config;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;

import com.aventstack.extentreports.MediaEntityBuilder;
import com.kdd.actions.ActionsClass;
import com.kdd.exceptions.InvalidKeywordException;
import com.kdd.reports.ReportManager;
import com.kdd.utility.ExcelReader;
import com.kdd.utility.ScreenshotUtility;

public class Executor implements GlobalVariables{

	private static final Logger Log = Logger.getLogger(Executor.class.getName());

	public Map<Integer, String> getListOfTestCasesToExecute() {
		Log.info("Getting List of Test Cases to execute.");
		return new ExcelReader().getTestCasesToRun(testDataSheet, runModeColumn, testCaseColumn);
	}

	/*
	 * //public void executeTestCase(String testDataSheet) throws Exception { public
	 * void executeTestCase(String testCaseID) throws Exception { String keyword;
	 * String value; String stepDesc;
	 * 
	 * try {
	 * 
	 * int numberOfRows = ExcelReader.getNumberOfRows(testCaseID);
	 * 
	 * //int numberOfRows = ExcelReader.getNumberOfCols(testDataSheet);
	 * 
	 * for (int iRow = 1; iRow < numberOfRows; iRow++){ keyword =
	 * ExcelManager.getInstance().getExcelReader().getCellData(iRow, keywordColumn,
	 * testCaseID); value =
	 * ExcelManager.getInstance().getExcelReader().getCellData(iRow, dataColumn,
	 * testCaseID); //stepDesc =
	 * ExcelManager.getInstance().getExcelReader().getCellData(iRow,
	 * testStepDescriptionColumn, testCaseID); stepDesc =
	 * ExcelManager.getInstance().getExcelReader().getCellData(iRow,
	 * testCaseDescriptionColumn, testCaseID);
	 * Log.info("Keyword: "+keyword+" Value: "+ value);
	 * 
	 * System.out.println("executeTestCase :1.keyword 2.value 3.stepDesc" +
	 * keyword+"_"+value+"_"+stepDesc);
	 * 
	 * try { executeAction(keyword, value);
	 * 
	 * } catch (Exception e) {
	 * Log.error("Exception Occurred while executing the step:\n", e); String
	 * imageFilePath =
	 * ScreenshotUtility.takeScreenShot(DriverManager.getInstance().getDriver(),
	 * testCaseID);
	 * ReportManager.getTest().fail(stepDesc+"<br/>Keyword: "+keyword+" | Value: "+
	 * value,
	 * MediaEntityBuilder.createScreenCaptureFromPath(imageFilePath).build());
	 * ReportManager.getTest().info(e); throw e; }
	 * 
	 * if(stepDesc !=null && !stepDesc.isEmpty() ) {
	 * ReportManager.getTest().pass(stepDesc+"<br/>Keyword: "+keyword+" | Value: "+
	 * value); } }
	 * 
	 * } catch (Exception e) { int testCaseRow = (Integer)
	 * SessionDataManager.getInstance().getSessionData("testCaseRow");
	 * ExcelManager.getInstance().getExcelReader().setCellData(FAIL, testCaseRow,
	 * resultColumn, testDataPath, testCaseID); throw e; } }
	 */
	
	public void executeTestCase(String testCaseID) throws Exception {
		//public void executeTestCase(String testCaseID) throws Exception {
			String keyword;
			//String value;
			List<Map<String,String>> value;
			String stepDesc;
			
			System.out.println("Executer.java executeTestCase(String testCaseID) :" + testCaseID);
			
			try {
				
				//int numberOfRows = ExcelReader.getNumberOfRows(testCaseID);
				
				int numberOfCols = ExcelReader.getNumberOfCols(testCaseID,testDataSheet);
				
				for (int iCol = keywordColumn; iCol <= numberOfCols-1; iCol++)
				  
				 {
					System.out.println("Executer.java executeTestCase(String testCaseID) keywordColumn:" + keywordColumn);
					
					keyword = ExcelManager.getInstance().getExcelReader().getCellColData(testCaseID,iCol, testDataSheet).toString();
					
					System.out.println("Executer.java keyword :" + keyword+" Column Number "+iCol);
					
					//if(keyword.equalsIgnoreCase("login"))
					//{
					//	String datasheet="LoginData";
						
					//}
					
					value =ExcelManager.getInstance().getExcelReader().getMapTestData("LoginData",testCaseID.toString());
					System.out.println("excelData as Map -value: "+value);
										
					
					//value = ExcelManager.getInstance().getExcelReader().getCellColData(testCaseID, dataColumn, testDataSheet);
					
					//value = ExcelManager.getInstance().getExcelReader().getExcelData("LoginData",testCaseID);
					//System.out.println("Executer.java value  :" + value+" & TestCase ID"+testCaseID);
					
					//stepDesc = ExcelManager.getInstance().getExcelReader().getCellData(iRow, testStepDescriptionColumn, testCaseID);
					//stepDesc = ExcelManager.getInstance().getExcelReader().getCellColData(testCaseID,iCol, testDataSheet);
					
					 stepDesc = ExcelManager.getInstance().getExcelReader().getCellColData(testCaseID,testCaseDescriptionColumn, testDataSheet).toString();
					
										
					 Log.info("Keyword: "+keyword+" Value: "+ value);	
					
					//System.out.println("executeTestCase :1.keyword 2.value 3.stepDesc" + keyword+"_"+value+"_"+stepDesc);
					
					try {
						  //executeAction(keyword, value);	
						    executeAction(keyword,value);						  						  
					    } 
					
					catch (Exception e)
					   {
						Log.error("Exception Occurred while executing the step:\n", e);					
						 String imageFilePath = ScreenshotUtility.takeScreenShot(DriverManager.getInstance().getDriver(), testDataSheet);
						ReportManager.getTest().fail(stepDesc+"<br/>Keyword: "+keyword+" | Value: "+ value, MediaEntityBuilder.createScreenCaptureFromPath(imageFilePath).build());
						 
						//ReportManager.getTest().fail("<br/>Keyword: "+keyword+" | Value: "+ value, MediaEntityBuilder.createScreenCaptureFromPath(imageFilePath).build());
						ReportManager.getTest().info(e);
						throw e;
					   }
					
					  if(stepDesc !=null && !stepDesc.isEmpty() )
					   {
						 ReportManager.getTest().pass(stepDesc+"<br/>Keyword: "+keyword+" | Value: "+ value);
						//ReportManager.getTest().pass(stepDesc+"<br/>Keyword: "+keyword+" | Value: "+ "");
					  }
					
				 }//For Loop End
				
			} 
			
			catch (Exception e) 
			{
				int testCaseRow = (Integer) SessionDataManager.getInstance().getSessionData("testCaseRow");
				ExcelManager.getInstance().getExcelReader().setCellData(FAIL, testCaseRow, resultColumn, testDataPath, testDataSheet);
				throw e;
			}
			
	}

	

	//private void executeAction(String keyword, String value) throws Exception {	
	
	private void executeAction(String keyword, List<Map<String, String>> value) throws Exception {	
		
		ActionsClass keywordActions = new ActionsClass();
		
		Method method[] = keywordActions.getClass().getMethods();
		
		boolean keywordFound = false;
		
				
		for(int i = 0;i < method.length;i++)
		 {
			System.out.println("method.length :"+method.length);
			if(method[i].getName().equals(keyword))
			   {
				System.out.println("method[i].getName() :"+method[i].getName().toString());
				method[i].invoke(keywordActions, value);
				keywordFound = true;
				break;
				
			}
		}
		if(!keywordFound) 
		{
			throw new InvalidKeywordException("Invalid Keyword "+keyword);
		}
	}
	
	
	
}
