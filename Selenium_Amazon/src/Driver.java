import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Driver {
	static WebDriver driver;

	public static void main(String[] args) throws Exception {

		String dt_path = "C:\\Users\\QA\\Desktop\\Amazon\\testsuit.xls";
		String[][] recData = ReUsableMethods.readXlSheet(dt_path, "Sheet1");

		//for chrome driver

		for(int i = 1; i < recData.length; i++){

			String execute = recData[i][1];
			System.out.println(execute);

			if(execute.equalsIgnoreCase("Y")){

				try{
					System.setProperty("webdriver.chrome.driver", "C:\\Users\\QA\\Downloads\\chromedriver_win32\\chromedriver.exe");
					driver = new ChromeDriver();

					String testCase = recData[i][2];

					System.out.println(testCase);

					ReUsableMethods.startReport(testCase, "C:\\Users\\QA\\Desktop\\Report\\", "chrome");
					/*Java Reflection*/
					Method tc = AutomationScripts.class.getMethod(testCase);
					tc.invoke(tc);



					ReUsableMethods.bw.close();

				} catch (InvocationTargetException e) {

					e.getCause().printStackTrace();
				}catch(Exception e){
					System.out.println(e);
				}


			}

		}

		//for firefox driver
		for(int i = 1; i < recData.length; i++){

			String execute = recData[i][1];
			System.out.println(execute);

			if(execute.equalsIgnoreCase("Y")){

				try{

					System.setProperty("webdriver.gecko.driver", "C:/Users/QA/Downloads/geckodriver-v0.18.0-win64/geckodriver.exe");
					driver = new FirefoxDriver(); 

					String testCase = recData[i][2];

					System.out.println(testCase);

					ReUsableMethods.startReport(testCase, "C:\\Users\\QA\\Desktop\\Report\\", "firefox");
					/*Java Reflection*/
					Method tc1 = AutomationScripts.class.getMethod(testCase);
					tc1.invoke(tc1);

					ReUsableMethods.bw.close();

				} catch (InvocationTargetException e) {

					e.getCause().printStackTrace();
				}catch(Exception e){
					System.out.println(e);
				}


			}

		}

		//for microsoftedge driver

		for(int i = 1; i < recData.length; i++){

			String execute = recData[i][1];
			System.out.println(execute);

			if(execute.equalsIgnoreCase("Y")){

				try{

					System.setProperty("webdriver.edge.driver", "C:\\Users\\QA\\Downloads\\MicrosoftWebDriver.exe");
					driver = new EdgeDriver();

					String testCase = recData[i][2];

					System.out.println(testCase);

					ReUsableMethods.startReport(testCase, "C:\\Users\\QA\\Desktop\\Report\\", "chrome");
					/*Java Reflection*/
					Method tc2 = AutomationScripts.class.getMethod(testCase);
					tc2.invoke(tc2);

					ReUsableMethods.bw.close();

				} catch (InvocationTargetException e) {

					e.getCause().printStackTrace();
				}catch(Exception e){
					System.out.println(e);
				}


			}

		}
	}

}



