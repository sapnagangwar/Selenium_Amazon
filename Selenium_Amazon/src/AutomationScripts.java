
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;


import org.openqa.selenium.Capabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.*;
import org.openqa.selenium.*;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;


public class AutomationScripts extends ReUsableMethods {

	public static void SearchIphone() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();

		//reading test data
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase1.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String searchData = (String)sheet.getRow(1).getCell(1).getStringCellValue();
		String expectedSearchresult = (String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expected = (String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expectedItemsInCart =(String)sheet.getRow(1).getCell(4).getStringCellValue();

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		//driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expected, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);
		Thread.sleep(2000);

		//locate search box
		WebElement searchBox = getWebElement(sheetObj, 1);
		String actual = driver.getCurrentUrl();

		//verify result
		String s2 = verify(expected, actual);
		Update_Report(s2, "verify url", "verified", driver);

		//enter iPhone6 in search box and click search btn
		String s3 = enterText(searchBox, searchData, "iPhone6");
		
		WebElement searchBtn = getWebElement(sheetObj, 2);
		String s4 = click(searchBtn, "search button clicked");

		Thread.sleep(500);
		
		//select iPhone6 phone from the options and click 

		String oldWindow = driver.getWindowHandle();
		Thread.sleep(500);
		WebElement selectedSearchedProduct = getWebElement(sheetObj, 4);
		String s5 = click(selectedSearchedProduct, "iphone6");

		Set<String> temp = driver.getWindowHandles();
		String actualTitle=null;
		for(String w:temp){
			driver.switchTo().window(w);
			actualTitle= driver.getTitle();
		}
		//product title 
		Thread.sleep(500);
		WebElement productTitle  = getWebElement(sheetObj, 5);
		String s6 = click(productTitle, "selected product title");

		//add to cart

		Thread.sleep(500);
		WebElement addToCart = getWebElement(sheetObj, 6);
		String s7 = click(addToCart, "adding iPhone6 to cart");

		//close the add on plan window
		Set<String> getAllWindows = driver.getWindowHandles();

		String[] getWindow = getAllWindows.toArray(new String[getAllWindows.size()]);
		driver.switchTo().window((getWindow[0]));

		Thread.sleep(2000);

		driver.findElement(By.xpath("//*[@id='a-popover-6']/div/div[1]/button")).click();

		//display on cart button
		Thread.sleep(500);
		WebElement cartDispaly = getWebElement(sheetObj, 7);
		System.out.println(cartDispaly.getText());
		String ActualItemsIncart = cartDispaly.getText();

		//verify items in cart

		String s8 =  verify(expectedItemsInCart, ActualItemsIncart);
		Update_Report(s8, "verify items in cart", "verified", driver);

		boolean condition=true;

		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") && s5.equals("Pass")
				 && s6.equals("Pass") && s7.equals("Pass") && s8.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 

		if(browserName.equals("chrome")){
			updatexlsTestSuit(1, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(1, 4, condition);
		}
		else{
			updatexlsTestSuit(1, 5, condition);
		}
		
		driver.close();
	}

	public static void TC02() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);

		//reading test data
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase2.xls";
		HSSFSheet sheet = readDataSheet(dt_path);
		
		String expectedTitleLink = (String)sheet.getRow(1).getCell(1).getStringCellValue();
		String expectedTodaysLinkTitle=(String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expectedCurrentURL=(String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expected=(String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expectedDepartmentItems =(String)sheet.getRow(1).getCell(4).getStringCellValue();

		String [] str = expectedDepartmentItems.split(",");
		ArrayList<String> expectedlist = new ArrayList<String>();
		for(String temp:str){
			expectedlist.add(temp);
		}

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualCurrentURL = driver.getCurrentUrl();

		String s1 = verify(expectedCurrentURL, actualCurrentURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls
		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);
		Thread.sleep(2000);

		//locate department dropdown and do mouse hover
		
		WebElement department = getWebElement(sheetObj, 10);
		Actions action = new Actions(driver);
		action.moveToElement(department).build().perform();

		//locate departments menu items and store in a list
				ArrayList<String> actualList = new ArrayList<String>();

				WebElement w1 = getWebElement(sheetObj,14);
				Thread.sleep(50);
				WebElement w2 = getWebElement(sheetObj,15);
				Thread.sleep(50);
				WebElement w3 = getWebElement(sheetObj,16);
				Thread.sleep(50);
				WebElement w4 = getWebElement(sheetObj,17);
				Thread.sleep(50);
				WebElement w5 = getWebElement(sheetObj,18);
				Thread.sleep(50);
				WebElement w6 = getWebElement(sheetObj,19);
				Thread.sleep(50);
				WebElement w7 = getWebElement(sheetObj,20);
				Thread.sleep(50);
				WebElement w8 = getWebElement(sheetObj,21);
				Thread.sleep(50);
				WebElement w9 = getWebElement(sheetObj,22);
				Thread.sleep(50);
				WebElement w10 = getWebElement(sheetObj,23);
				Thread.sleep(50);
				WebElement w11= getWebElement(sheetObj,24);
				Thread.sleep(50);
				WebElement w12= getWebElement(sheetObj,25);
				Thread.sleep(50);
				WebElement w13= getWebElement(sheetObj,26);
				Thread.sleep(50);
				WebElement w14= getWebElement(sheetObj,27);
				Thread.sleep(50);
				WebElement w15= getWebElement(sheetObj,28);
				Thread.sleep(50);
				WebElement w16= getWebElement(sheetObj,29);
				Thread.sleep(50);
				WebElement w17= getWebElement(sheetObj,30);
				Thread.sleep(50);
				WebElement w18= getWebElement(sheetObj,31);
				Thread.sleep(50);
				WebElement w19= getWebElement(sheetObj,32);
				Thread.sleep(50);
				WebElement w20= getWebElement(sheetObj,33);
				Thread.sleep(50);
				WebElement w21= getWebElement(sheetObj,34);
				Thread.sleep(50);

				actualList.add(w1.getText());
				actualList.add(w2.getText());
				actualList.add(w3.getText());
				actualList.add(w4.getText());
				actualList.add(w5.getText());
				actualList.add(w6.getText());
				actualList.add(w7.getText());
				actualList.add(w8.getText());
				actualList.add(w9.getText());
				actualList.add(w10.getText());
				actualList.add(w11.getText());
				actualList.add(w12.getText());
				actualList.add(w13.getText());
				actualList.add(w14.getText());
				actualList.add(w15.getText());
				actualList.add(w16.getText());
				actualList.add(w17.getText());
				actualList.add(w18.getText());
				actualList.add(w19.getText());
				actualList.add(w20.getText());
				actualList.add(w21.getText());

				//verify dropdown list
				String s2 = verify( expectedlist, actualList);
				Update_Report(s2, "verify departments dropdown", "compared dropdown list with the expected dropdown", driver);

		//locate amazon link
		Thread.sleep(500);
		WebElement amazonLink = getWebElement(sheetObj, 12);
		Actions actionLink = new Actions(driver);
		actionLink.moveToElement(amazonLink).build().perform();

		//verify amazon.com link
		String s3 = click(amazonLink, "amazon link clicked");

		String actualTitleLink = driver.getTitle();
		System.out.println("a: " + actualTitleLink);
		System.out.println("e: " + expectedTitleLink);

		String s4 = verify(expectedTitleLink, actualTitleLink);
		Update_Report(s2, "verify amazon link", "link", driver);
		driver.navigate().back();
		Thread.sleep(1500);

		//locate "Today's Deals " link

		Thread.sleep(500);

		WebElement todaysLink =getWebElement(sheetObj, 13);
		Actions actionTodaysLink = new Actions(driver);
		actionTodaysLink.moveToElement(todaysLink).build().perform();

		//verify todays deal link
		String s5 = click(todaysLink, "todays link clicked");

		String actualTodaysLinkTitle = driver.getTitle();
		System.out.println(driver.getTitle());

		String s6 = verify(expectedTodaysLinkTitle, actualTodaysLinkTitle);
		Update_Report(s6, "verify today's Deal", "Today's Deal link", driver);
		driver.navigate().back();
		Thread.sleep(1500);

		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") && s5.equals("Pass") && s6.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 

		if(browserName.equals("chrome")){
			updatexlsTestSuit(2, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(2, 4, condition);
		}
		else{
			updatexlsTestSuit(2, 5, condition);
		}

	}

	public static void TC03() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);
		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase3.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();

		//read expected departments dropdown from xls and store items in a list
		String expectedDepartmentItems =(String)sheet.getRow(1).getCell(2).getStringCellValue();

		String [] str = expectedDepartmentItems.split(",");
		ArrayList<String> expectedlist = new ArrayList<String>();
		for(String temp:str){
			expectedlist.add(temp);
		}


		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();
		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls
		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//locate department dropdown and do mouse hover
		WebElement departments =getWebElement(sheetObj, 10);

		Thread.sleep(500);

		Actions action = new Actions(driver);
		action.moveToElement(departments).build().perform();
		Thread.sleep(35000);

		//locate departments menu items and store in a list
		ArrayList<String> actualList = new ArrayList<String>();

		WebElement w1 = getWebElement(sheetObj,14);
		Thread.sleep(50);
		WebElement w2 = getWebElement(sheetObj,15);
		Thread.sleep(50);
		WebElement w3 = getWebElement(sheetObj,16);
		Thread.sleep(50);
		WebElement w4 = getWebElement(sheetObj,17);
		Thread.sleep(50);
		WebElement w5 = getWebElement(sheetObj,18);
		Thread.sleep(50);
		WebElement w6 = getWebElement(sheetObj,19);
		Thread.sleep(50);
		WebElement w7 = getWebElement(sheetObj,20);
		Thread.sleep(50);
		WebElement w8 = getWebElement(sheetObj,21);
		Thread.sleep(50);
		WebElement w9 = getWebElement(sheetObj,22);
		Thread.sleep(50);
		WebElement w10 = getWebElement(sheetObj,23);
		Thread.sleep(50);
		WebElement w11= getWebElement(sheetObj,24);
		Thread.sleep(50);
		WebElement w12= getWebElement(sheetObj,25);
		Thread.sleep(50);
		WebElement w13= getWebElement(sheetObj,26);
		Thread.sleep(50);
		WebElement w14= getWebElement(sheetObj,27);
		Thread.sleep(50);
		WebElement w15= getWebElement(sheetObj,28);
		Thread.sleep(50);
		WebElement w16= getWebElement(sheetObj,29);
		Thread.sleep(50);
		WebElement w17= getWebElement(sheetObj,30);
		Thread.sleep(50);
		WebElement w18= getWebElement(sheetObj,31);
		Thread.sleep(50);
		WebElement w19= getWebElement(sheetObj,32);
		Thread.sleep(50);
		WebElement w20= getWebElement(sheetObj,33);
		Thread.sleep(50);
		WebElement w21= getWebElement(sheetObj,34);
		Thread.sleep(50);

		actualList.add(w1.getText());
		actualList.add(w2.getText());
		actualList.add(w3.getText());
		actualList.add(w4.getText());
		actualList.add(w5.getText());
		actualList.add(w6.getText());
		actualList.add(w7.getText());
		actualList.add(w8.getText());
		actualList.add(w9.getText());
		actualList.add(w10.getText());
		actualList.add(w11.getText());
		actualList.add(w12.getText());
		actualList.add(w13.getText());
		actualList.add(w14.getText());
		actualList.add(w15.getText());
		actualList.add(w16.getText());
		actualList.add(w17.getText());
		actualList.add(w18.getText());
		actualList.add(w19.getText());
		actualList.add(w20.getText());
		actualList.add(w21.getText());

		//verify dropdown list
		String s2 = verify( expectedlist, actualList);
		Update_Report(s2, "verify departments dropdown", "compared dropdown list with the expected dropdown", driver);

		System.out.println(Arrays.toString(expectedlist.toArray()));
		System.out.println(Arrays.toString(actualList.toArray()));

		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") ){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 

		if(browserName.equals("chrome")){
			updatexlsTestSuit(3, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(3, 4, condition);
		}
		else{
			updatexlsTestSuit(3, 5, condition);
		}


	}
	public static void TC04() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);
		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase4.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();

		//read  Sign In expected dropdown from xls and store items in a list

		String expectedDepartmentItems =(String)sheet.getRow(1).getCell(2).getStringCellValue();

		String [] str = expectedDepartmentItems.split(",");
		ArrayList<String> expectedlist = new ArrayList<String>();
		for(String temp:str){
			expectedlist.add(temp);
		}

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls
		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//mouse hover on SignIn to access dropdown
		WebElement signIn = getWebElement(sheetObj, 35);
		Actions action = new Actions(driver);
		action.moveToElement(signIn).build().perform();
		Thread.sleep(5000);

		//display SignIn Btn
		WebElement signInBtn = getWebElement(sheetObj, 36);
		String s2 = click(signInBtn, "Sign In");
		driver.navigate().back();
		Thread.sleep(1500);


		//again locate :stale
		WebElement signIn1 = getWebElement(sheetObj, 35);
		Actions action1 = new Actions(driver);
		action1.moveToElement(signIn1).build().perform();
		Thread.sleep(500);

		//Display New Customer Start here link
		WebElement newCustomerLink = getWebElement(sheetObj, 37);
		Thread.sleep(1000);
		String s3 = click(newCustomerLink, "New Customer , Start Here");
		Thread.sleep(2000);
		driver.navigate().back();
		Thread.sleep(4000);

		WebElement w = getWebElement(sheetObj, 35);
		//WebElement w = driver.findElement(By.xpath("//*[@id='nav-link-accountList']"));
		Actions action2 = new Actions(driver);
		action2.moveToElement(w).build().perform();

		Thread.sleep(5000);

		//menu items and store in a list
		WebElement account = getWebElement(sheetObj, 39);

		WebElement order = getWebElement(sheetObj, 40);

		WebElement list = getWebElement(sheetObj, 41);

		WebElement recommendations = getWebElement(sheetObj, 42);

		WebElement subsAndSave = getWebElement(sheetObj, 43);

		WebElement memAndSubs = getWebElement(sheetObj, 44);

		WebElement serviceReq = getWebElement(sheetObj, 45);

		WebElement primeMem = getWebElement(sheetObj, 46);

		WebElement garage = getWebElement(sheetObj, 47);

		WebElement register = getWebElement(sheetObj, 48);

		WebElement creditCard = getWebElement(sheetObj, 49);

		WebElement contentAnddevices = getWebElement(sheetObj, 50);

		WebElement musicLib= getWebElement(sheetObj, 51);

		WebElement photos = getWebElement(sheetObj, 52);

		WebElement drive = getWebElement(sheetObj, 53);

		WebElement video = getWebElement(sheetObj, 54);

		WebElement kindle = getWebElement(sheetObj, 55);

		WebElement watchlist= getWebElement(sheetObj, 56);

		WebElement videoLib = getWebElement(sheetObj, 57);

		WebElement androidAppsAndDevices = getWebElement(sheetObj, 58);

		ArrayList<String> actualList = new ArrayList<String>();

		actualList.add(account.getText());
		actualList.add(order.getText());
		actualList.add(list.getText());
		actualList.add(recommendations.getText());
		actualList.add(subsAndSave.getText());
		actualList.add(memAndSubs.getText());
		actualList.add(serviceReq.getText());
		actualList.add(primeMem.getText());
		actualList.add(garage.getText());
		actualList.add(register.getText());
		actualList.add(creditCard.getText());
		actualList.add(contentAnddevices.getText());
		actualList.add(musicLib.getText());
		actualList.add(photos.getText());
		actualList.add(drive.getText());
		actualList.add(video.getText());
		actualList.add(kindle.getText());
		actualList.add(watchlist.getText());

		//verify dropdown list
		String s4 = verify( expectedlist, actualList);
		Update_Report(s4, "verify SignIn dropdown", "compared dropdown list with the expected dropdown", driver);

		System.out.println(Arrays.toString(expectedlist.toArray()));
		System.out.println(Arrays.toString(actualList.toArray()));

		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") ){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 

		if(browserName.equals("chrome")){
			updatexlsTestSuit(4, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(4, 4, condition);
		}
		else{
			updatexlsTestSuit(4, 5, condition);
		}

	}

	public static void TC05() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);
		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase5.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();

		//read  "all" menu  dropdown from search bar  from xls and store items in a list

		String expectedAllSearchBoxMenuItems =(String)sheet.getRow(1).getCell(2).getStringCellValue();

		String [] array = expectedAllSearchBoxMenuItems.split(",");
		ArrayList<String> expectedlist = new ArrayList<String>();
		for(String temp:array){
			expectedlist.add(temp);
		}

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//locate all and doubleclick all search - list drop down items

		WebElement allSearch = getWebElement(sheetObj, 59);
		Actions builder = new Actions(driver);
		builder.doubleClick(allSearch).perform();

		Thread.sleep(2000);

		Select s = new Select(allSearch);

		List<WebElement> items = s.getOptions();

		ArrayList<String> actualList = new ArrayList<String>();

		int index =0;

		for(WebElement e : items){
			actualList.add(e.getText());
		}

		//locate Clothing, Shoes & Jewelry , click
		s.selectByVisibleText("Clothing, Shoes & Jewelry");

		Thread.sleep(1500);

		//use click or clickAndHold on builder
		builder.click(allSearch).perform();

		Select subMenu = new Select(allSearch);

		subMenu.selectByIndex(15);
		Thread.sleep(2000);
		Update_Report("Sub menu", "sub menu: Women selection", "verify", driver);

		subMenu.selectByIndex(16);
		Thread.sleep(2000);
		Update_Report("Sub menu", "sub menu: Men selection", "verify", driver);

		subMenu.selectByIndex(17);
		Thread.sleep(2000);
		Update_Report("Sub menu", "sub menu: Girls selection", "verify", driver);

		subMenu.selectByIndex(18);
		Thread.sleep(2000);
		Update_Report("Sub menu", "sub menu: Boys selection", "verify", driver);

		subMenu.selectByIndex(19);
		Thread.sleep(2000);
		Update_Report("Sub menu", "sub menu: Baby selection", "verify", driver);

		Thread.sleep(1500);

		//verify dropdown list
		String s2 = verify( expectedlist, actualList);
		Update_Report(s2, "verify All search dropdown", "compared dropdown list with the expected dropdown", driver);

		System.out.println(Arrays.toString(expectedlist.toArray()));
		System.out.println(Arrays.toString(actualList.toArray()));

		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 

		if(browserName.equals("chrome")){
			updatexlsTestSuit(5, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(5, 4, condition);
		}
		else{
			updatexlsTestSuit(5, 5, condition);
		}

	}

	public static void TC06() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);
		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase6.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();
		String expectedSearchData =(String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expectedTitle = (String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expectedItemsInCart = (String)sheet.getRow(1).getCell(4).getStringCellValue();
		String expFinalCartItem = (String)sheet.getRow(1).getCell(5).getStringCellValue();

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//enter serch data in search panel and validate

		WebElement serachBox = getWebElement(sheetObj, 1);

		enterText(serachBox, expectedSearchData, "enter search data in search panel");

		String phoneModel = serachBox.getText();

		WebElement searchBtn =getWebElement(sheetObj, 2);
		click(searchBtn, "clicked search button");

		String s2 = verify(expectedSearchData, phoneModel);
		Update_Report(s2, " verify search data ", "compared the entered search data with actual seach data", driver);

		//select phone model
		WebElement phone64Gb= getWebElement(sheetObj, 61);
		click(phone64Gb, "clicked the selected product");
		System.out.println(driver.getTitle());
		System.out.println(driver.getCurrentUrl());

		String actualTitle = driver.getTitle();
		String s3 = verify(expectedTitle, actualTitle);
		Update_Report(s3, " verify product page ", "compared the entered product  with actual product page title", driver);

		//add to cart

		WebElement addToCart =getWebElement(sheetObj, 6);
		String s4 = click(addToCart, "adding iPhone6 to cart");


		//close the add on plan window

		Set<String> getAllWindows = driver.getWindowHandles();
		String[] getWindow = getAllWindows.toArray(new String[getAllWindows.size()]);
		driver.switchTo().window(getWindow[0]);

		Thread.sleep(2000);

		driver.findElement(By.xpath("//*[@id='siNoCoverage-announce']")).click();

		//WebElement noThanksBtn = getWebElement(sheetObj, 62);
		//String s5 = click(noThanksBtn, "click for no insurance");
		Thread.sleep(1000);

		//display on cart button

		WebElement cartDispaly = getWebElement(sheetObj, 7);
		System.out.println(cartDispaly.getText());
		String ActualItemsIncart = cartDispaly.getText();

		//verify items in cart

		String s6 =  verify(expectedItemsInCart, ActualItemsIncart);
		Update_Report(s6, "verify items in cart", "verified", driver);

		//delete item from cart
		WebElement cart = getWebElement(sheetObj, 63);
		String s7 = click(cart, "cart button clicked");

		WebElement delete64Gb = getWebElement(sheetObj, 64);
		Thread.sleep(2000);
		String s8 = click(delete64Gb, "delete link clicked");
		Thread.sleep(2000);

		//display on cart button

		WebElement cartDispalyAfterDel = getWebElement(sheetObj, 7);
		System.out.println(cartDispalyAfterDel.getText());
		System.out.println(expFinalCartItem);

		WebElement cartfinalCount = getWebElement(sheetObj, 66);
		String ActualItemsIncartAfterDel = cartfinalCount.getText();


		//verify items in cart

		String s9 =  verify(expFinalCartItem, ActualItemsIncartAfterDel);
		Update_Report(s9, "verify items in cart", "verified", driver);

		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") && s6.equals("Pass")&& s7.equals("Pass") && s8.equals("Pass") && s9.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		if(browserName.equals("chrome")){
			updatexlsTestSuit(6, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(6, 4, condition);
		}
		else{
			updatexlsTestSuit(6, 5, condition);
		}

	}
	public static void TC07() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);
		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase7.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();
		String expectedPageTitle =(String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expectedHeadingMsg = (String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expectedSubMenu1 = (String)sheet.getRow(1).getCell(4).getStringCellValue();
		String expectedSubMenu2 = (String)sheet.getRow(1).getCell(5).getStringCellValue();
		String expectedSubMenu3 = (String)sheet.getRow(1).getCell(6).getStringCellValue();
		String expectedSubMenu4 = (String)sheet.getRow(1).getCell(7).getStringCellValue();
		String expectedSubMenu5 = (String)sheet.getRow(1).getCell(8).getStringCellValue();
		String expectedSubMenu6 = (String)sheet.getRow(1).getCell(9).getStringCellValue();

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//click on Help
		WebElement helpLink = getWebElement(sheetObj, 67);

		String s2 = click(helpLink, "clicked Help");

		Thread.sleep(1000);

		//verify help page
		String actualPageTitle = driver.getTitle();

		String s3 = verify(expectedPageTitle, actualPageTitle);
		Update_Report(s3, "verified help page title", "compared expected and actual page titles", driver);

		//display and verify heading msg

		WebElement headingMsg = getWebElement(sheetObj, 68);
		String actualHeadingMsg = headingMsg.getText();

		String s4 = verify(expectedHeadingMsg, actualHeadingMsg);
		Update_Report(s4, "verified help heading message", "compared expected and actual help messages", driver);

		//verify 6 sub menu

		WebElement orders = getWebElement(sheetObj, 69);
		String actualSubMenu1 = orders.getText();
		String s5 = verify(expectedSubMenu1, actualSubMenu1);
		Update_Report(s5, "verified Sub menu: your orders", "compared expected and actual ", driver);

		WebElement returnAndRefund = getWebElement(sheetObj,70 );
		String actualSubMenu2 = returnAndRefund.getText();
		String s6 = verify(expectedSubMenu2, actualSubMenu2);
		Update_Report(s6, "verified Sub menu: return and refunds", "compared expected and actual ", driver);

		WebElement deviceSupport = getWebElement(sheetObj,71 );
		String actualSubMenu3 = deviceSupport.getText();
		String s7 = verify(expectedSubMenu3, actualSubMenu3);
		Update_Report(s7, "verified Sub menu: device Support", "compared expected and actual ", driver);

		WebElement managePrime = getWebElement(sheetObj,72);
		String actualSubMenu4 = managePrime.getText();
		String s8 = verify(expectedSubMenu4, actualSubMenu4);
		Update_Report(s8, "verified Sub menu: Manage Prime", "compared expected and actual ", driver);

		WebElement paymentAndGiftCards = getWebElement(sheetObj,73 );
		String actualSubMenu5 = paymentAndGiftCards.getText();
		String s9 = verify(expectedSubMenu5, actualSubMenu5);
		Update_Report(s9, "verified Sub menu: payment and gift cards", "compared expected and actual ", driver);

		WebElement accountSettings = getWebElement(sheetObj,74);
		String actualSubMenu6 = accountSettings.getText();
		String s10= verify(expectedSubMenu6, actualSubMenu6);
		Update_Report(s10, "verified Sub menu: account settings", "compared expected and actual ", driver);

		//locate and verify "Find more solution" search box and serach icon

		WebElement findMoreSolSearchBox = getWebElement(sheetObj,75 );


		driver.findElement(By.xpath("//*[@id='helpsearch']")).sendKeys("Find more solutions....");
		Thread.sleep(1000);
		WebElement searchIcon = getWebElement(sheetObj,76 );
		searchIcon.click();
		String s11 = driver.findElement(By.xpath("//*[@id='helpsearch']")).getAttribute("placeholder");
		Thread.sleep(1000);
		System.out.println(s11);

		//condition for test Pass/Fail
		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") 
				&& s5.equals("Pass")&& s6.equals("Pass") && s7.equals("Pass") 
				&& s8.equals("Pass") && s9.equals("Pass") && s10.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 
		if(browserName.equals("chrome")){
			updatexlsTestSuit(7, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(7, 4, condition);
		}
		else{
			updatexlsTestSuit(7, 5, condition);
		}



	}

	public static void TC08() throws Exception {

		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);
		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase8.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();
		String textToBeSearched =(String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expQtySelectedInDropdown =(String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expfinalCartCount =(String)sheet.getRow(1).getCell(4).getStringCellValue();

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		driver.manage().window().maximize();
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//enter searched text in search 
		WebElement searchBox = getWebElement(sheetObj, 1);
		enterText(searchBox, textToBeSearched, "search text entered");

		WebElement searchBtn = getWebElement(sheetObj, 2);
		String s2 = click(searchBtn, "search clicked");

		//verify searched text
		System.out.println(driver.getTitle());

		//click on selected book link and verify
		WebElement searchedBookLink  = getWebElement(sheetObj,78);
		String s3 = click(searchedBookLink, "selected book link clicked");
		//System.out.println(driver.getTitle());

		//locate and set quantity 4

		Thread.sleep(500);

		WebElement qty  = getWebElement(sheetObj,79);
		String s4 = click(qty, "qty drop down clicked");

		WebElement qty5  = getWebElement(sheetObj,80);
		String s5 = click(qty5, "selected 5 book for cart");
		System.out.println(qty5.getText());
		String actualQtySelected= qty5.getText();

		String s6 = verify(expQtySelectedInDropdown,actualQtySelected);
		Update_Report(s6, "verified drop down display", "compared expected with actual", driver);


		//driver.findElement(By.cssSelector("span.a-dropdown-label")).click();

		//driver.findElement(By.id("quantity_4")).click();

		driver.findElement(By.xpath("//*[@id='add-to-cart-button']")).click();;

		/*
		//Add to Cart and verify
		WebElement addToCart  = getWebElement(sheetObj,81);
		Thread.sleep(500);
		String s7 = click(addToCart, "added to cart");
		 */
		Thread.sleep(1500);

		WebElement finalCartCount  = getWebElement(sheetObj,66);
		String actualfinalCartCount = finalCartCount.getText();

		String s8 = verify(expfinalCartCount,actualfinalCartCount);
		Update_Report(s8, "verified cart display", "compared expected with actual", driver);

		//condition for test Pass/Fail
		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") && s5.equals("Pass") 
				&& s6.equals("Pass") && s8.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 
		if(browserName.equals("chrome")){
			updatexlsTestSuit(8, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(8, 4, condition);
		}
		else{
			updatexlsTestSuit(8, 5, condition);
		}


	}

	public static void TC09() throws Exception {
		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();

		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase9.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();
		String textToBeSearched =(String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expQtySelectedInDropdown =(String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expfinalCartCount =(String)sheet.getRow(1).getCell(4).getStringCellValue();

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//enter searched text in search 
		WebElement searchBox = getWebElement(sheetObj, 1);
		enterText(searchBox, textToBeSearched, "search text entered");

		WebElement searchBtn = getWebElement(sheetObj, 2);
		String s2 = click(searchBtn, "search clicked");

		//verify searched text
		System.out.println(driver.getTitle());

		//click on selected book link and verify
		WebElement searchedBookLink  = getWebElement(sheetObj,78);
		String s3 = click(searchedBookLink, "selected book link clicked");
		//System.out.println(driver.getTitle());

		//locate and set quantity 4

		Thread.sleep(500);

		WebElement qty  = getWebElement(sheetObj,79);
		String s4 = click(qty, "qty drop down clicked");

		WebElement qty5  = getWebElement(sheetObj,80);
		String s5 = click(qty5, "selected 5 book for cart");
		System.out.println(qty5.getText());
		String actualQtySelected= qty5.getText();

		String s6 = verify(expQtySelectedInDropdown,actualQtySelected);
		Update_Report(s6, "verified drop down display", "compared expected with actual", driver);


		driver.findElement(By.cssSelector("span.a-dropdown-label")).click();
		Thread.sleep(1500);

		driver.findElement(By.id("quantity_4")).click();
		Thread.sleep(3500);

		//Add to Cart
		driver.findElement(By.xpath("//*[@id='add-to-cart-button']")).click();

		Thread.sleep(500);

		//Click to cart
		WebElement cartBtn = getWebElement(sheetObj, 7);
		String s7 = click(cartBtn, "cart button clicked");

		//update Qty 
		driver.findElement(By.id("a-autoid-2-announce")).click();
		driver.findElement(By.id("dropdown1_3")).click();

		//save for later link
		Thread.sleep(2000);
		// WebElement w = driver.findElement(By.name("submit.save-for-later.C3UQM48TLC8H1Q"));
		WebElement w = driver.findElement(By.xpath("//*[@id='activeCartViewForm']/div[2]/div/div[4]/div/div[1]/div/div/div[2]/div/span[2]/span/input"));

		Thread.sleep(2000);
		w.click();

		//condition for test Pass/Fail
		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass") && s5.equals("Pass") 
				&& s6.equals("Pass") && s7.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 
		if(browserName.equals("chrome")){
			updatexlsTestSuit(9, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(9, 4, condition);
		}
		else{
			updatexlsTestSuit(9, 5, condition);
		}


	}

	public static void TC10() throws Exception {
		//get browser name
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browserName = cap.getBrowserName().toLowerCase();
		//System.out.println(browserName);

		//read data sheet
		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\TCxls\\TestCase10.xls";
		HSSFSheet sheet = readDataSheet(dt_path);

		String expectedURL =(String)sheet.getRow(1).getCell(1).getStringCellValue();
		String textToBeSearched =(String)sheet.getRow(1).getCell(2).getStringCellValue();
		String expSearchDisplay1 =(String)sheet.getRow(1).getCell(3).getStringCellValue();
		String expSearchDisplay2 =(String)sheet.getRow(1).getCell(4).getStringCellValue();
		String expSearchDisplay3 =(String)sheet.getRow(1).getCell(5).getStringCellValue();

		/*Launch URL*/
		driver.get("https://www.amazon.com/");
		Thread.sleep(500);

		String actualURL = driver.getCurrentUrl();

		//verify application launch
		String s1 = verify(expectedURL, actualURL);
		Update_Report(s1, "verify url", "verified", driver);

		//read ff objectName, ff object type and ff object property from Object Rep.xls

		String objRepo_path = "C:\\Users\\QA\\Desktop\\Amazon\\ObjectRep.xls";
		HSSFSheet sheetObj = readDataSheet(objRepo_path);

		//driver.findElement(By.xpath("//*[@id='twotabsearchtextbox']")).sendKeys("Iphone");

		//enter searched text in search 
		WebElement searchBox = getWebElement(sheetObj, 1);
		enterText(searchBox, textToBeSearched, "search text entered");

		//verify searched text
		System.out.println(driver.getTitle());

		try {
			Thread.sleep(1000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		//List<WebElement> allOptions = driver.findElements(By.xpath("//*[text()='iphon']"));
		List<WebElement> allOptions = driver.findElements(By.xpath("//*[contains(text(),'iphon')]"));

		for (int i = 0; i < allOptions.size(); i++) {
			String option = ((WebElement) allOptions.get(i)).getText();
			Thread.sleep(200);
			System.out.println(option);
		}

		Thread.sleep(1200);
		String actualSearchDisplay1 = allOptions.get(0).getText();
		String actualSearchDisplay2 = allOptions.get(1).getText();
		String actualSearchDisplay3 = allOptions.get(2).getText();

		String s2 = verify(expSearchDisplay1, actualSearchDisplay1);
		Update_Report(s2, "drp down display1", "compared exp and actual", driver);

		String s3 = verify(expSearchDisplay2, actualSearchDisplay2);
		Update_Report(s3, "drp down display2", "compared exp and actual", driver);

		String s4 = verify(expSearchDisplay3, actualSearchDisplay3);
		Update_Report(s4, "drp down display3", "compared exp and actual", driver);

		//condition for test Pass/Fail

		boolean condition=true;
		if(s1.equals("Pass") && s2.equals("Pass") && s3.equals("Pass") && s4.equals("Pass")){
			condition = true;
		}
		else{
			condition = false;
		}

		//update testsuit.xls 

		if(browserName.equals("chrome")){
			updatexlsTestSuit(10, 3, condition);
		}
		else if(browserName.equals("firefox")){
			updatexlsTestSuit(10, 4, condition);
		}
		else{
			updatexlsTestSuit(10, 5, condition);
		}
	}
}

