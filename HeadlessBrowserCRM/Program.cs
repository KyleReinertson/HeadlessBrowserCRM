using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;


namespace HeadlessBrowserCRM
{
    class Program
    {

        private static Guid moduserGuid = new Guid("089824D8-0059-47D3-B6DD-D7B7DA317A26");
        private static string FROM = "admin@CRM108762.onmicrosoft.com";
        private static string FROMNAME = "CRM Online User";
        private static string TO = "kyle.reinertson@mural365.com";
        private static string SMTP_USERNAME = "admin@CRM108762.onmicrosoft.com";
        private static string SMTP_PASSWORD = "kyle.reiner@2364";
        private static string CONFIGSET = "ConfigSet";
        private static string HOST = "smtp.office365.com";
        private static int PORT = 587;
        private static string username = "admin@CRM108762.onmicrosoft.com";
        private static string password = "kyle.reiner@2364";
        private static string url = "https://crm108762.api.crm.dynamics.com/XRMServices/2011/Organization.svc";
        private static string path = @"C:\Users\kyle.reinertson\source\repos\TryThis.csv";
        private static string html = "C:\\Users\\kyle.reinertson\\source\\repos\\CRMConsoleActivities\\CRMConsoleActivities\\HTMLPage1.html";
        private static string muralPath = "https://crm108762.crm.dynamics.com/main.aspx?forceClassic=1";
        private static string O365loginButtonId = "idSIButton9";
        private static string O365usernameId = "i0116";
        private static string O365passwordId = "i0118";



        #region ConnecttoCRM

        public static IOrganizationService ConnecttoCRM()
        {
            IOrganizationService organizationService = null;


            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = username;
                clientCredentials.UserName.Password = password;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri(url), null, clientCredentials, null);

                if (organizationService != null)
                {
                    Guid gOrgId = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).OrganizationId;
                    if (gOrgId != Guid.Empty)
                    {

                        Console.WriteLine("***Connection Established Successfully...***");
                        Console.ReadLine();

                    }

                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occurred - " + ex.Message);
                Console.ReadLine();
            }
            return organizationService;

        }

        #endregion ConnecttoCRM


        static void Main(string[] args)
        {

            FirefoxOptions options = new FirefoxOptions();
            options.AddArguments("-headless");

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService();
            service.FirefoxBinaryPath = @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe";
            FirefoxDriver driver = new FirefoxDriver(service);

            driver.Navigate().GoToUrl("https://crm108762.crm.dynamics.com/main.aspx?forceClassic=1");
            IWebElement query = driver.FindElement(By.Id(O365usernameId));
            query.SendKeys(username);
            query = driver.FindElement(By.Id(O365loginButtonId));
            query.Click();

            IWebElement query2 = driver.FindElement(By.Id(O365passwordId));
            query2.SendKeys(password);
            query2 = driver.FindElement(By.Id(O365loginButtonId));
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator = By.Id(O365loginButtonId);
            wait.Until(drv => By.Id(O365loginButtonId));
            query2.Click();

            IWebElement query3 = driver.FindElement(By.Id(O365loginButtonId));
            query3 = driver.FindElement(By.Id(O365loginButtonId));
            WebDriverWait wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator2 = By.Id(O365loginButtonId);
            wait.Until(drv => By.Id(O365loginButtonId));
            query3.Click();

            IWebElement query4 = driver.FindElement(By.Id("TabSFA"));
            query4 = driver.FindElement(By.Id("TabSFA"));
            WebDriverWait wait3 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator3 = By.Id("TabSFA");
            wait.Until(drv => By.Id("TabSFA"));
            query4.Click();

            IWebElement query5 = driver.FindElement(By.Id("Settings"));
            query5 = driver.FindElement(By.Id("Settings"));
            WebDriverWait wait4 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator4 = By.Id("Settings");
            wait.Until(drv => By.Id("Settings"));
            query5.Click();

            IWebElement query6 = driver.FindElement(By.Id("nav_security"));
            query6 = driver.FindElement(By.Id("nav_security"));
            WebDriverWait wait5 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator5 = By.Id("nav_security");
            wait.Until(drv => By.Id("nav_security"));
            query6.Click();

        }
    }
}
