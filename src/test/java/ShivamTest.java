import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import org.openqa.selenium.support.ui.ExpectedConditions;
import java.io.*;
import java.util.List;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Set;


public class ShivamTest {
    WebDriver driver;
    int i = 1;
    String gamename="";
    String numberoftournament ="";
    String link ="";
    int respCode=0;
    String urls = "//*[@class='games-item']//a[@href='xyz']";




    @BeforeClass
    public void LaunchWebsite() {
        System.setProperty("webdriver.chrome.driver", "/Applications/chromedriver");
        driver = new ChromeDriver();
        driver.navigate().to("https://game.tv");

    }

    @Test
    public void method_01() throws IOException, InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, 20);
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebElement games = driver.findElement(By.linkText("Garena Free Fire Tournaments"));
        driver.manage().window().maximize();
        List<WebElement> items = driver.findElements(By.xpath("//*[@class='games-item']//a"));
        ShivamTest.column_name();
        for (WebElement myElement : items)
        {
            js.executeScript("arguments[0].scrollIntoView();", games);
             link = myElement.getAttribute("href");
            String hreflink = link.substring(19);
            By by = By.xpath(urls.replace("xyz", hreflink));
            gamename = myElement.getText();
            WebElement page = driver.findElement(by);
            wait.until(ExpectedConditions.elementToBeClickable(page));
            String selectLinkOpeninNewTab = Keys.chord(Keys.COMMAND,Keys.RETURN);
            driver.findElement(By.linkText(gamename)).sendKeys(selectLinkOpeninNewTab);
            Set<String> allWindowHandles = driver.getWindowHandles();
            String parentWindowHandle = driver.getWindowHandle();
            for(String child : allWindowHandles)
            {
                if(!parentWindowHandle.equalsIgnoreCase(child)) {
                    driver.switchTo().window(child);
                    WebElement tournament = driver.findElement(By.xpath("//*[@class='count-tournaments']"));
                    numberoftournament = tournament.getText();
                    HttpURLConnection huc = null;
                    huc = (HttpURLConnection) (new URL(link).openConnection());
                    huc.setRequestMethod("HEAD");
                    huc.connect();
                    respCode = huc.getResponseCode();
                    driver.close();
                }

            }
            driver.switchTo().window(parentWindowHandle);
            writedata(i, gamename, link, respCode, numberoftournament);
            wait.until(ExpectedConditions.elementToBeClickable(by));
            i++;
        }
    }
    @AfterClass
    public void closebrowser()
    {
        driver.close();
    }


    public  void writedata(int number, String Game_Name, String PAGE_URL, int PAGE_Status, String Tournament_count) throws IOException {
        File scr = new File("/Users/shivam.sohane/Desktop/writefile/shivamtest.xlsx");
        FileInputStream fis = new FileInputStream(scr);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet1 = wb.getSheetAt(0);
        for (int row = i; row <= i; row++) {
            for (int col = 0; col <= 5; col++) {
                if (col == 0)
                    sheet1.createRow(row).createCell(col).setCellValue(number);
                if (col == 1)
                    sheet1.getRow(row).createCell(col).setCellValue(Game_Name);
                if (col == 2)
                    sheet1.getRow(row).createCell(col).setCellValue(PAGE_URL);
                if (col == 3)
                    sheet1.getRow(row).createCell(col).setCellValue(PAGE_Status);
                if (col == 4)
                    sheet1.getRow(row).createCell(col).setCellValue(Tournament_count);
                FileOutputStream fout = new FileOutputStream(scr);
                wb.write(fout);
            }
        }
        wb.close();
    }

    public static void column_name() throws IOException {
        File scr = new File("/Users/shivam.sohane/Desktop/writefile/shivamtest.xlsx");
        FileInputStream fis = new FileInputStream(scr);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet1 = wb.getSheetAt(0);
        sheet1.createRow(0).createCell(0).setCellValue("#");
        sheet1.getRow(0).createCell(1).setCellValue("Game name");
        sheet1.getRow(0).createCell(2).setCellValue("PAGE URL");
        sheet1.getRow(0).createCell(3).setCellValue("Page Status");
        sheet1.getRow(0).createCell(4).setCellValue("Tournament count");
        FileOutputStream fout = new FileOutputStream(scr);
        wb.write(fout);
        wb.close();
    }
}

















