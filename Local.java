import java.time.Duration;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Local {
	
	public void testScript() throws Exception 
	{
		
		String testDataPath = "./TestData/DataTest.xlsx";
		
		ExcelUtil excel = new ExcelUtil();
		HashMap<Integer, ArrayList<String>> data = excel.getData(testDataPath, "Test Data");
		HashMap<Integer, ArrayList<String>> contactData = excel.getData(testDataPath, "Contact List");
		HashMap<Integer, ArrayList<String>> addressData = excel.getData(testDataPath, "Address List");
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Youssef.Elhossiny\\Downloads\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));//Program waits 5 seconds if it fails to synchronize with the browser. After the 5 seconds it will fail.
		driver.get("http://192.168.2.21:8080/SuretyMaster/login.jsf#!"); //URL
		driver.findElement(By.id("loginForm:emailAddress")).sendKeys("jane.doe@xenex.cloud"); //Username
		driver.findElement(By.name("loginForm:password")).sendKeys("Joe707"); //Password
		
		//If there is a space between the class locator it means that there is two different classes. 
		driver.findElement(By.className("btn-default")).click(); //Click sign in button
		
		int dataRows = data.size();
		
		for(int r=1; r<dataRows; r++) //Goes through the first row of obligee profile and then goes to the next one once the loop is complete
		{ 			
			Thread.sleep(2000);
			//driver.findElement(By.id("menuForm:mainMenu5006_span")).click(); // Click on Obligee on the top bar
			//driver.findElement(By.xpath("//div/span[@id='menuForm:menu5037:anchor']")).click(); //Click on Obligee Profile in the drop down menu
			//Thread.sleep(1000);
			String profileSave = "Fail";
			try {			
				driver.findElement(By.cssSelector("a[class='bulletin']:nth-child(7)")).click(); //Click on Obligee List
				driver.findElement(By.cssSelector("a[class='addNewLink']")).click(); //Click on add new button
				
				WebElement obligeeType = driver.findElement(By.xpath("//select[@id='profileForm:type']")); // Obligee type
				obligeeType.click();
				Select sel = new Select(obligeeType);
				sel.selectByVisibleText(excel.getValue(data, r, "Obligee Type")); //Gets the data from excel and selects the matching visible text on the page
				
				driver.findElement(By.id("profileForm:save")).click(); //click save (Due to bug in page can remove once fixed)
				driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
				Thread.sleep(1000);
				
				driver.findElement(By.xpath("//input[@id='profileForm:nme']")).sendKeys(excel.getValue(data, r, "Obligee")); //Obligee 
				WebElement Status = driver.findElement(By.xpath("//select[@id='profileForm:status']")); //Status of Obligee
				Status.click();
				Select status = new Select(Status);
				status.selectByVisibleText(excel.getValue(data, r, "Status")); //Gets the data from excel and selects the matching visible text on the page
				
				driver.findElement(By.id("profileForm:phone")).sendKeys(excel.getValue(data, r, "Phone")); //Enter Phone number
				
				if(excel.getValue(data, r, "Website") != null && !excel.getValue(data, r, "Website").trim().equals("")) { //If data under the Website coloumn in excel is not null or doesn't equal to a blank enter the statement
					driver.findElement(By.id("profileForm:websiteURL")).sendKeys(excel.getValue(data, r, "Website")); //Enter Website
				}
				
				if(excel.getValue(data, r, "Fax") != null && !excel.getValue(data, r, "Fax").trim().equals("")) { //If data under the Fax coloumn in excel is not null or doesn't equal to a blank enter the statement
					driver.findElement(By.id("profileForm:fax")).sendKeys(excel.getValue(data, r, "Fax")); //Enter Fax
				}
				
				WebElement Language = driver.findElement(By.xpath("//select[@id='profileForm:lang']")); //Select desired language
				Language.click();
				Select lang = new Select(Language);
				lang.selectByVisibleText(excel.getValue(data, r, "Language")); //Gets the data from excel and selects the matching visible text on the page
				
				if (excel.getValue(data, r, "Obligee Type").contains("State Government") || excel.getValue(data, r, "Obligee Type").contains("Federal Government") || excel.getValue(data, r, "Obligee Type").contains("Local / Country / City / Township")) { //If the obligee type is State Government or Federal Government or Local / Country / City / Township enter the loop. 
				WebElement Country = driver.findElement(By.xpath("//select[@id='profileForm:cntry']")); //Select desired Country
				Country.click();
				Select country = new Select(Country);
				country.selectByVisibleText(excel.getValue(data, r, "Country")); //Gets the data from excel and selects the matching visible text on the page
				}
				
				if(excel.getValue(data, r, "Obligee Type").contains("State Government") || excel.getValue(data, r, "Obligee Type").contains("Local / Country / City / Township")) { //If the obligee type is State Government or Local / Country / City / Township enter the loop. 
				WebElement State = driver.findElement(By.xpath("//select[@id='profileForm:state']")); //Select Desired state
				State.click();
				Select state = new Select(State);
				state.selectByVisibleText(excel.getValue(data, r, "State")); //Gets the data from excel and selects the matching visible text on the page
				}		
				
				driver.findElement(By.id("profileForm:save")).click(); //click Save button
				
				if(driver.findElement(By.cssSelector("dl[id='globalMsg'] span[class='rich-messages-label']"))!=null) //If save message pops up print Pass in excel, else it will print Fail 
				{
					profileSave = "Pass";
				}
			}	catch(Exception e) {
				
			}	finally {
				excel.writeValue(data, profileSave, r, "Test Data", "Success?", testDataPath); //Print either pass or fail in the "Success?" Coloumn in excel
			}
			
			
//			driver.findElement(By.xpath("(//a[@class='bulletin'])[7]")).click();
//			driver.findElement(By.xpath("(//a[@href='#'][normalize-space()='Amu security'])[1]")).click();
//			driver.findElement(By.xpath("//td[@id='profileForm:contact_lbl']")).click(); //contact list
//			driver.findElement(By.xpath("//a[@id='profileForm:cntctLst:0:contactName']")).click(); //Add new button
			
			int contactRows = contactData.size();
			
			for (int c=1; c<contactRows; c++) { //Goes through contact data page on excel, goes through first row and then to the next once loop is done
				if (excel.getValue(contactData, c, "Obligee").contains(excel.getValue(data, r, "Obligee"))){ //If the same obligee name is present in the contact data as the obligee profile created enter the if statement
					String contactSave = "Fail";
					try {
						Thread.sleep(1000);
						driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
						driver.findElement(By.id("profileForm:contact_lbl")).click(); //After saving the obligee profile click on the Contact List Label
						Thread.sleep(2000);
						driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
						driver.findElement(By.xpath("//a[normalize-space()='Add New']")).click(); // Add new button
						
						driver.findElement(By.id("contactForm:frstNme")).sendKeys(excel.getValue(contactData, c, "First Name")); //Enter first name
						
						driver.findElement(By.xpath("//input[@id='contactForm:lstNme']")).sendKeys(excel.getValue(contactData, c, "Last Name")); //Enter Last Name
					
						if (excel.getValue(contactData, c, "Position") != null) {
							WebElement Position = driver.findElement(By.xpath("//select[@id='contactForm:position']")); // Select Position
							Position.click();
							Select position = new Select(Position);
							position.selectByVisibleText(excel.getValue(contactData, c, "Position"));
						}
	
						WebElement Status1 = driver.findElement(By.xpath("//select[@id='contactForm:status']")); //Enter Status
						Status1.click();
						Select status1 = new Select(Status1);
						status1.selectByVisibleText(excel.getValue(contactData, c, "Status"));
						
						if (excel.getValue(contactData, c, "Phone") != null && !excel.getValue(contactData, c, "Phone").trim().equals(""))
						{
							driver.findElement(By.xpath("//input[@id='contactForm:phone']")).sendKeys(excel.getValue(contactData, c, "Phone")); //Enter Phone
						}
						
						if(excel.getValue(contactData, c, "Ext") != null)
						{
							driver.findElement(By.xpath("//input[@id='contactForm:ext']")).sendKeys(excel.getValue(contactData, c, "Ext")); //Enter Extension
						}
						
						if(excel.getValue(contactData, c, "Cell") != null)
						{
							driver.findElement(By.xpath("//input[@id='contactForm:cell']")).sendKeys(excel.getValue(contactData, c, "Cell")); //Enter Cell
						}
						
						if(excel.getValue(contactData, c, "Fax") != null)
						{
							driver.findElement(By.xpath("//input[@id='contactForm:fax']")).sendKeys(excel.getValue(contactData, c, "Fax")); //Enter Fax
						}
						
						driver.findElement(By.xpath("//input[@id='contactForm:email']")).sendKeys(excel.getValue(contactData, c, "Email")); // Enter Email
						
						if((excel.getValue(contactData, c, "Permission to Contact?")).contains("Yes")) //if permission to contact is yes click on the checkbox
						{
							driver.findElement(By.xpath("//input[@id='contactForm:permCntct']")).click(); 
						}
						
						Thread.sleep(2000);
						driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
						driver.findElement(By.xpath("(//input[@id='contactForm:save'])[1]")).click(); //Save button
						if(driver.findElement(By.cssSelector("dl[id='globalMsg'] span[class='rich-messages-label']"))!=null) //If save message pops up print Pass in the excel page if not message will be fail
						{
							contactSave = "Pass";
						}
					}catch(Exception e) {
						
					} finally {
							excel.writeValue(contactData, contactSave, c, "Contact List", "Success?", testDataPath); // Print message "Pass" or "Fail" under "Success?" page
							
							driver.findElement(By.xpath("//a[@id='contactForm:name']")).click(); // Click on the obligee name link to go back
					}
				} 
			}
			
			int addressRows = addressData.size();
			
			for (int a=1; a<addressRows; a++) {
				if (excel.getValue(addressData, a, "Obligee").contains(excel.getValue(data, r, "Obligee"))) { //If obligee name in excels addreessData sheet matches the oblige profile name enter the statement.
					String addressSave = "Fail";
					try {
						driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
						Thread.sleep(1000);
						driver.findElement(By.xpath("//td[@id='profileForm:address_lbl']")).click(); //CLick on the address label to add new adress
						Thread.sleep(1000);
						driver.findElement(By.xpath("//a[@class='addNewLink']")).click(); //Add new button
						
						driver.findElement(By.xpath("//input[@id='addressForm:addr1']")).sendKeys(excel.getValue(addressData, a, "Address 1")); //Enter Address 1
						
						driver.findElement(By.xpath("//input[@id='addressForm:addr2']")).sendKeys(excel.getValue(addressData, a, "Address 2")); //Enter Address 2
						
						driver.findElement(By.xpath("//input[@id='addressForm:zip']")).sendKeys(excel.getValue(addressData, a, "Zip/Postal Code")); //Enter Zip/Postal Code
						
						//driver.findElement(By.xpath("//a[@id='addressForm:listId:0:j_id103']")).click();
//						WebElement addressCountry = driver.findElement(By.xpath("//select[@id='addressForm:cntry']")); //enter desired Country
//						addressCountry.click();
//						Select addresscountry = new Select(addressCountry);
//						addresscountry.selectByVisibleText(excel.getValue(addressData, a, "Country"));

						
						if (excel.getValue(addressData, a, "Primary Address").contains("Yes")) { //If excel data contains "Yes" under Primary Address click the checkbox
							driver.findElement(By.xpath("//input[@id='addressForm:addrPrim']")).click(); //Primary Address
						}
						
						if(excel.getValue(addressData, a, "Mailing Address").contains("Yes")) { //If excel data contains "Yes" under Mailing Address click the checkbox
							driver.findElement(By.xpath("//input[@id='addressForm:addrMail']")).click(); //Mailing Address
						}
						
						if(excel.getValue(addressData, a, "Billing Address").contains("Yes")) { //If excel data contains "Yes" under Billing Address click the checkbox
							driver.findElement(By.xpath("//input[@id='addressForm:addrBill']")).click(); //Billing Address
						}
						
						if(excel.getValue(addressData, a, "Shipping Address").contains("Yes")) { //If excel data contains "Yes" under Shipping Address click the checkbox
							driver.findElement(By.xpath("//input[@id='addressForm:addrShip']")).click(); //Shipping Address
						}
						
						driver.findElement(By.xpath("//input[@id='addressForm:save']")).click(); //CLick save
						if(driver.findElement(By.cssSelector("dl[id='globalMsg'] span[class='rich-messages-label']")) != null) { //If save message pops up save message will change to "Pass" rather than "Fail"
							addressSave = "Pass";
						}
					}
					catch(Exception e) {
						
					} 
					finally {
						excel.writeValue(addressData, addressSave, a, "Address List", "Success?", testDataPath); // Print save message into excel page under "Succeess?" coloumn
					}
					//Click on Obligee Name to send you back to profile
					Thread.sleep(1000);
					driver.findElement(By.xpath("//input[@id='addressForm:back']")).click(); //Click on back button
				} 
			}
			
			driver.findElement(By.xpath("//div[@id='menuForm:homeButton']")).click(); //Home Button
		}
			
	}
		

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Local local = new Local();
		local.testScript();
		
	}

}
