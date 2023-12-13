using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        // Chrome WebDriver를 초기화합니다.
        IWebDriver driver = new ChromeDriver();

        try
        {
            // 웹 페이지로 이동합니다.
            driver.Navigate().GoToUrl("https://www.bizinfo.go.kr/web/lay1/bbs/S1T175C112/AK/61/list.do");

            // JavaScript 함수를 호출합니다.
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            string script = "return fileBlank('/webapp/upload/forms/2023/files' + '/' + '100390892_2023_1.hwp', '100390892_개인별연간근태기록표결근휴가.hwp');";
            string result = jsExecutor.ExecuteScript(script).ToString();

            // 결과를 출력합니다.
            Console.WriteLine("Result of fileBlank function: " + result);
        }
        finally
        {
            // 브라우저를 종료합니다.
            driver.Quit();
        }
    
    }
}