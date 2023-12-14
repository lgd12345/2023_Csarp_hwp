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
            //IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            //string script = "return fileBlank('/webapp/upload/forms/2023/files' + '/' + '100390892_2023_1.hwp', '100390892_개인별연간근태기록표결근휴가.hwp');";
            //jsExecutor.ExecuteScript(script);
            
            //함수를 변경해서 출력해보기

            // <a> 요소 찾기
            IWebElement link = driver.FindElement(By.XPath("//a[contains(@class, 'basic-btn01') and contains(@onclick, 'fileBlank')]"));

            // JavaScript 주입 - 해당 <a> 요소를 클릭할 때 특별한 동작 주입
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            string script = @"
            var docuGateUrlResult;
            var originalClick = arguments[0].click;
            arguments[0].click = function() {
                originalClick.call(arguments[0]);
                docuGateUrlResult = docuGateUrl;
            };
        ";

            string result = (string)jsExecutor.ExecuteScript(script, link);

            // 결과 출력
            Console.WriteLine("Returned value: " + result);

            result = (string)jsExecutor.ExecuteScript("return docuGateUrlResult;");
            Console.WriteLine("Returned value: " + result);

            // <a> 요소 클릭
            link.Click();




            //string result = jsExecutor.ExecuteScript(script).ToString();
            // 결과를 출력합니다.
            //Console.WriteLine("Result of fileBlank function: " + result);
        }
        catch (Exception ex)
        {
            // 예외가 발생했을 때 에러 코드를 출력합니다.
            Console.WriteLine("Error occurred. Error code: " + ex.Message);
        }
        finally
        {
            // 브라우저를 종료합니다.
            //driver.Quit();
        }
    
    }
}