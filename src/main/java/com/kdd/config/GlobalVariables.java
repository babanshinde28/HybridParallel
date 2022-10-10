package com.kdd.config;

public interface GlobalVariables {

	String baseDirectory = System.getProperty("user.dir");
	String configPath = baseDirectory+"/config.properties";

	// Application URL
	
	//String baseURL = "https://jpetstore.aspectran.com/catalog/";
	
	// Application Direct Login Page
	//String baseURL="https://jpetstore.aspectran.com/account/signonForm";
    String baseURL="https://demo.guru99.com/test/newtours/";
    
	// Path to test data sheet with test cases to run and test case details
	//String testDataPath = baseDirectory+"/src/test/resources/data/TestDataSheet.xlsx";
	//String testDataPath = baseDirectory+"/src/test/resources/data/TestDataSheetNew.xls";
	String testDataPath = "C://DataSheets//TestDataSheetNewPolicy.xls";
	
	//String testDataPath = baseDirectory+"/src/test/resources/data/TestDataSheetNewPolicy.xls";
	
	//String testDataSheet = "TestCases";
	String testDataSheet = "TestSuite";
	
	//***** TestSuite Sheet
	//int testCaseColumn = 0;
	//int testCaseDescriptionColumn = 1;
	//int runModeColumn = 2;
	//int resultColumn = 3;
	
	int testCaseColumn = 0;
	int runModeColumn = 1;
	int resultColumn = 2;
	int testCaseDescriptionColumn = 4;

	//***** Columns in Test Case sheet
	//int testStepsColumn = 0;
	//int testStepDescriptionColumn = 1;
	//int keywordColumn = 2;
	//int dataColumn = 3;	
	
	//***** Columns in Test Case sheet
		//int testStepsColumn = 0;
		//int testStepDescriptionColumn = 1;
		int keywordColumn = 5;
		int dataColumn = 3;	

	String PASS = "PASS";
	String FAIL = "FAIL";
	String SKIP = "SKIP";

	// Wait times
	int implicitWaitTime = 10;
	long objectWaitTime = 30;

	// Screenshot folder and file details
	String screenshotFolder = baseDirectory+"/screenshots/";
	String fileFormat = ".png";
	
	// Report related details 
	String extentConfigFilePath = baseDirectory+"/extent-config.xml";
	String htmlReportPath = baseDirectory+"/target/html-report/";
	String htmlFileName = "ExtentReport.html";
}
