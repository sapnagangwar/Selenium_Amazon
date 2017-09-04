import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.remote.RemoteWebDriver;

public class ReUsableMethods extends Driver {
	static BufferedWriter bw = null;
	static BufferedWriter bw1 = null;
	static String htmlname;
	static String objType;
	static String objName;
	static String TestData;
	static String rootPath;
	static int report;


	static Date cur_dt = null;
	static String filenamer;
	static String TestReport;
	int rowcnt;
	static String exeStatus = "True";
	static int iflag = 0;
	static int j = 1;

	static String fireFoxBrowser;
	static String chromeBrowser;
	static String microsoftedgeBrowser;

	static String result;

	static int intRowCount = 0;
	static String dataTablePath;
	static int i;
	static String browserName;


	/* Name Of the method: verify
	 * Brief Description: compares two strings for equality
	 * Arguments: two strings
	 * Created by: Automation team
	 * Creation Date: Aug 23 2017
	 * Last Modified: Aug 23 2017
	 */
	public static String verify(String expected, String actual){
		if(expected.equals(actual)){
			return "Pass";
		}
		else{
			return "Fail";
		}
	}


	/* Name Of the method: enterText
	 * Brief Description: Enter the text value to the text box
	 * Arguments: obj --> Text box object, textVal --> value to be entered, objName --> name of the object
	 * Created by: Automation team
	 * Creation Date: Aug 23 2017
	 * Last Modified: Aug 23 2017
	 * */
	public static String enterText(WebElement obj, String textVal, String objName) throws IOException{
		if(obj.isDisplayed()){
			obj.sendKeys(textVal);
			Update_Report("Pass", objName +" entered", "text entered",driver);
			return "Pass";

		}else{
			Update_Report("Fail", objName+ " not entered", objName +"not visible",driver);
			return "Fail";

		}
	}

	/* Name Of the method: clickButton
	 * Brief Description: Click on button
	 * Arguments: obj --> web object,  objName --> name of the object
	 * Created by: Automation team
	 * Creation Date: Aug 23 2017
	 * Last Modified: Aug 23 2017
	 * */

	public static String click(WebElement obj,  String objName) throws IOException{

		if(obj.isDisplayed()){
			obj.click();
			Update_Report("Pass", objName + " clicked", " clicked successfully",Driver.driver);
			return "Pass";			

		}else{
			Update_Report("Fail", objName + " not visible", " not able to click,check your application",Driver.driver);
			return "Fail";
		}

	}

	/* Name of the Method: clearText
	 * Brief Description: Clear the text value to the text box
	 * Arguments: TextBox object,textVal--->value to be entered;objName--->name of the object
	 * created by :Automation team
	 * Creation date: Aug 23,2017
	 * last modified:Aug 23 ,2017
	 */
	public static String clearTextBox(WebElement obj, String objName) throws IOException{
		if(obj.isDisplayed()){
			obj.clear();
			Update_Report("Pass ", objName +" cleared", "cleared successfully",driver);
			return "Pass";		

		}
		else{
			Update_Report("Fail ", objName +" is not traceable, please check your application", " unsuccessful",driver);
			return "Fail";		
		}
	}

	/* Name of the Method: getTextInfo
	 * Brief Description: Get text value of WebElement
	 * Arguments: WebElement
	 * created by :Automation team
	 * Creation date: Aug 23,2017
	 * last modified:Aug 23 ,2017
	 */

	public static String getTextInfo(WebElement obj, String objName) throws IOException{
		if(obj.isDisplayed()){
			obj.getText();
			Update_Report("Pass ", objName + "present : ", obj.getText(),driver);
			return "Pass";
		}
		else{
			Update_Report("Fail", objName + " not present", "no error msg dispalyed",driver);
			return "Fail";
		}
	}

	/* Name of the Method: getLocator
	 * Brief Description: Return a instance of By class based on type of locator
	 * Arguments: string
	 * created by :Automation team
	 * Creation date: Aug 23,2017
	 * last modified:Aug 23 ,2017
	 */

	public static By getLocator(String locatorType, String locatorValue) throws Exception{

		if(locatorType.toLowerCase().equals("id"))
			return By.id(locatorValue);
		else if(locatorType.toLowerCase().equals("name"))
			return By.name(locatorValue);
		else if((locatorType.toLowerCase().equals("classname")) || (locatorType.toLowerCase().equals("class")))
			return By.className(locatorValue);
		else if((locatorType.toLowerCase().equals("tagname")) || (locatorType.toLowerCase().equals("tag")))
			return By.className(locatorValue);
		else if((locatorType.toLowerCase().equals("linktext")) || (locatorType.toLowerCase().equals("link")))
			return By.linkText(locatorValue);
		else if(locatorType.toLowerCase().equals("partiallinktext"))
			return By.partialLinkText(locatorValue);
		else if((locatorType.toLowerCase().equals("cssselector")) || (locatorType.toLowerCase().equals("css")))
			return By.cssSelector(locatorValue);
		else if(locatorType.toLowerCase().equals("xpath"))
			return By.xpath(locatorValue);
		else
			throw new Exception("Locator type '" + locatorType + "' not defined!!");
	}

	/* Method Name: readXlSheet
	 * Method description:Read content from Xl sheet
	 * Arguments:dt_path --> Path of Xl sheet, sheetName --> Name of the sheet user is accessing 
	 * Created by:Automation Team
	 * Creation Date: July 26 2017
	 * Last Modified: July 26 2017
	 */
	public static String[][] readXlSheet(String dt_path, String sheetName) throws IOException{

		/*Step 1: Get the XL Path*/
		File xlFile = new File(dt_path);

		/*Step2: Access the Xl File*/
		FileInputStream xlDoc = new FileInputStream(xlFile);


		/*Step3: Access the work book */
		HSSFWorkbook wb = new HSSFWorkbook(xlDoc);

		/*Step4: Access the Sheet */
		HSSFSheet sheet = wb.getSheet(sheetName);

		int iRowCount = sheet.getLastRowNum()+1;
		int iColCount = sheet.getRow(0).getLastCellNum();

		String[][] xlData = new String[iRowCount][iColCount];

		for(int i = 0; i < iRowCount; i++){
			for(int j = 0; j <iColCount; j++){
				xlData[i][j] = sheet.getRow(i).getCell(j).getStringCellValue();

			}

		}

		return xlData;

	}


	/*
	 * Name of the Method: startReport
	 * Brief description : Creates HTML report template
	 * Arguments: scriptname:test script name to run,ReportsPath:HTML report path to create,browserName:browser the script is running
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */
	public static void startReport(String scriptName, String ReportsPath,String browserName) throws IOException{
		j =0;
		String strResultPath = null;
		String testScriptName =scriptName;

		cur_dt = new Date(); 
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String strTimeStamp = dateFormat.format(cur_dt);

		if (ReportsPath == "") { 

			ReportsPath = "C:\\";
		}

		if (ReportsPath.endsWith("\\")) { 
			ReportsPath = ReportsPath + "\\";
		}

		strResultPath = ReportsPath + "Log" + "/" +testScriptName +"/"; 
		File f = new File(strResultPath);
		f.mkdirs();
		htmlname = strResultPath  + testScriptName + "_" + strTimeStamp 
				+ ".html";

		bw = new BufferedWriter(new FileWriter(htmlname));

		bw.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Browser Name</B></FONT></TD><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>"
				+ browserName + "</B></FONT></TD></TR>");
		bw.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR COLS=7><TD BGCOLOR=#BDBDBD WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>SL No</B></FONT></TD>"
				+ "<TD BGCOLOR=#BDBDBD WIDTH=10%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Step Name</B></FONT></TD>"
				+ "<TD BGCOLOR=#BDBDBD WIDTH=10%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Execution Time</B></FONT></TD> "
				+ "<TD BGCOLOR=#BDBDBD WIDTH=10%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD>"
				+ "<TD BGCOLOR=#BDBDBD WIDTH=47%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Detail Report</B></FONT></TD></TR>");


	}

	/*
	 * Name of the Method: Update_Report
	 * Brief description : Updates HTML report with test results
	 * Arguments: Res_type:holds the response of test script,Action:Action performed,result:contains test results
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */


	public static void Update_Report(String Res_type,String Action, String result,WebDriver dr) throws IOException {
		Date exec_time = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String str_time = dateFormat.format(exec_time);

		if (Res_type.startsWith("Pass")) {
			//String ss1Path= screenshot(dr);

			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=3%><FONT FACE=VERDANA SIZE=2>"
					+ (j++)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10%><FONT FACE=VERDANA SIZE=2>"
					+Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
					+ "Passed"
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
					+ result + "</FONT></TD></TR>");

		} else if (Res_type.startsWith("Fail")) {
			//To generate report in only single file

			String ss1Path= screenshot(dr);
			exeStatus = "Failed";
			report = 1;
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=3%><FONT FACE=VERDANA SIZE=2>"
					+ (j++)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10%><FONT FACE=VERDANA SIZE=2>"
					+Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "<a href= "
					+ ss1Path

					+ "  style=\"color: #FF0000\"> Failed </a>"

						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ result + "</FONT></TD></TR>");


		} 
	}

	/*
	 * Name of the Method: screenshot
	 * Brief description : creates screenshots
	 * Arguments: WebDriver
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */
	public static String screenshot(WebDriver dr) throws IOException{

		Date exec_time = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String str_time = dateFormat.format(exec_time);
		String  ss1Path = "C:\\Users\\QA\\Desktop\\Report\\ScreenShots\\"+ str_time+".png";
		//String  ss1Path = "C:\\Users\\QA\\Desktop\\"+ str_time+".png";
		File scrFile = ((TakesScreenshot)dr).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(scrFile, new File(ss1Path));
		return ss1Path;
	}


	/*
	 * Name of the Method: updatexlsTestSuit
	 * Brief description : Updates Test result in an xls sheet
	 * Arguments: result type(Pass/Fail),row and col where result can be updated
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */
	public static void updatexlsTestSuit(int row, int col, boolean condition) throws IOException, InterruptedException{
		FileInputStream fsIP= new FileInputStream(new File("C:\\Users\\QA\\Desktop\\Amazon\\testsuit.xls"));  
		//Access the workbook                  
		HSSFWorkbook w= new HSSFWorkbook(fsIP);
		//Access the worksheet, so that we can update / modify it. 
		HSSFSheet worksheet = w.getSheetAt(0); 
		// declare a Cell object
		Cell cell = null; 
		cell = worksheet.getRow(row).getCell(col); 


		/* create HSSFHyperlink objects */

		HSSFHyperlink link =new HSSFHyperlink(HSSFHyperlink.LINK_FILE);

		String filePath = htmlname;

		link.setAddress(filePath);


		if(condition==true ){

			// Get current cell value value and overwrite the value
			cell.setCellValue("Y");

			cell.setHyperlink(link);

			Thread.sleep(10000);

			HSSFCellStyle titleStyle =w.createCellStyle();
			titleStyle.setFillForegroundColor(new HSSFColor.LIGHT_GREEN().getIndex());
			titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(titleStyle);

			Thread.sleep(10000);
		}
		else{
			cell.setCellValue("N");
			cell.setHyperlink(link);

			Thread.sleep(1000);
			HSSFCellStyle titleStyle =w.createCellStyle();
			titleStyle.setFillForegroundColor(new HSSFColor.RED().getIndex());
			titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(titleStyle);

		}
		Thread.sleep(2000);

		//Close the InputStream  
		fsIP.close(); 
		//Open FileOutputStream to write updates
		FileOutputStream output_file = new FileOutputStream(new File("C:\\Users\\QA\\Desktop\\Amazon\\testsuit.xls"));  
		//write changes
		w.write(output_file);
		//close the stream
		output_file.close();

	}

	/*
	 * Name of the Method: readDataSheet
	 * Brief description :reading test data
	 * Arguments: file path
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */

	public static HSSFSheet readDataSheet(String dt_path) throws IOException{
		File xlFile = new File(dt_path);
		FileInputStream xlDoc = new FileInputStream(xlFile);
		HSSFWorkbook wb = new HSSFWorkbook(xlDoc);
		HSSFSheet sheet = wb.getSheet("Sheet1");
		return sheet;

	}

	/*
	 * Name of the Method: getWebElement
	 * Brief description :get WebElement from xls sheet
	 * Arguments: file path
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */

	public static WebElement getWebElement(HSSFSheet sheet, int rowIndex) throws Exception{
		String x = sheet.getRow(rowIndex).getCell(2).getStringCellValue();
		String y = sheet.getRow(rowIndex).getCell(3).getStringCellValue();
		return driver.findElement(getLocator(x, y));

	}

	/*
	 * Name of the Method: verify
	 * Brief description : compare two list
	 * Arguments: lists
	 * Created by: Automation team
	 * Creation date : July 17 2017
	 * last modified:  July 17 2017
	 */

	public static String verify( ArrayList<String> expectedlist, ArrayList<String> actualList){

		if(expectedlist.equals(actualList))
			return "Pass";
		else
			return "Fail";
	}



}



