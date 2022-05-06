package BaseLayerDDT;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class BaseClass {
	public static WebDriver driver;
	
	public static void Initialization() {
		System.setProperty("webdriver.chrome.driver","F:\\pratiksha new notes\\chromedriver_win32 (1)\\chromedriver.exe");
	    driver = new ChromeDriver();
	    driver.manage().window().maximize();
	    driver.manage().timeouts().pageLoadTimeout(30,TimeUnit.SECONDS);  
	    driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);  
		driver.get("https://opensource-demo.orangehrmlive.com");
	}

}
