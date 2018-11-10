using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using System.IO;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using System.Linq;
using System.Net.Mail;
using System.Collections.Generic;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Firefox;
using System.Threading;
using OpenQA.Selenium.Support.UI;

namespace HeadlessBrowserCRM
{
    class Program
    {

        private static Guid moduserGuid = new Guid("089824D8-0059-47D3-B6DD-D7B7DA317A26");
        static string FROM = "admin@CRM108762.onmicrosoft.com";
        static string FROMNAME = "CRM Online User";
        static string TO = "kyle.reinertson@mural365.com";
        static string SMTP_USERNAME = "admin@CRM108762.onmicrosoft.com";
        static string SMTP_PASSWORD = "kyle.reiner@2364";
        static string CONFIGSET = "ConfigSet";
        static string HOST = "smtp.office365.com";
        static int PORT = 587;
        static string username = "admin@CRM108762.onmicrosoft.com";
        static string password = "kyle.reiner@2364";
        static string url = "https://crm108762.api.crm.dynamics.com/XRMServices/2011/Organization.svc";
        static string path = @"C:\Users\kyle.reinertson\source\repos\TryThis.csv";
        static string html = "C:\\Users\\kyle.reinertson\\source\\repos\\CRMConsoleActivities\\CRMConsoleActivities\\HTMLPage1.html";
        static string muralPath = "https://crm108762.crm.dynamics.com/main.aspx?forceClassic=1";


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

            // IWebDriver driver = new ChromeDriver();

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService();
            service.FirefoxBinaryPath = @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe";
            FirefoxDriver driver = new FirefoxDriver(service);

            driver.Navigate().GoToUrl("https://crm108762.crm.dynamics.com/main.aspx?forceClassic=1");
            IWebElement query = driver.FindElement(By.Id("i0116"));
            query.SendKeys("admin@CRM108762.onmicrosoft.com");
            query = driver.FindElement(By.Id("idSIButton9"));
            query.Click();
            // now you just have to do the same for the password page
            //  driver.Quit();



           // IWebDriver driver1 = new ChromeDriver();
           // driver1.Navigate().GoToUrl("https://crm108762.crm.dynamics.com/main.aspx?forceClassic=1");
            IWebElement query2 = driver.FindElement(By.Id("i0118"));
            query2.SendKeys("kyle.reiner@2364");
            query2 = driver.FindElement(By.Id("idSIButton9"));
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            var locator = By.Id("idSIButton9");
            wait.Until(drv => By.Id("idSIButton9"));
            query2.Click();


            IWebElement query3 = driver.FindElement(By.Id("idSIButton9"));
            query3 = driver.FindElement(By.Id("idSIButton9"));
            query3.Click();




            /* {
                 var chromeOptions = new ChromeOptions();
                 {

                     chromeOptions.BinaryLocation = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";

                 };

                 // chromeOptions.AddArguments(new List<string>() { "headless", "disable-gpu" });

                 var driver = new ChromeDriver(chromeOptions);
                 //driver.Navigate().GoToUrl("https://crm108762.crm.dynamics.com/main.aspx?forceClassic=1");

                 driver.Navigate().GoToUrl("https://portal.office.com");

                 var userNameField = driver.FindElementById("i0116");
                 WebDriverWait wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                 wait2.Until(drv => By.Id("i0116"));

                 var userPasswordField = driver.FindElementById("i0118");
                 WebDriverWait wait3 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                 wait3.Until(drv => By.Id("i0118"));

                 var loginButton1 = driver.FindElementById("idSIButton9"); 
                 WebDriverWait wait4 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                 wait4.Until(drv => By.Id("idSIButton9"));

                 userNameField.SendKeys(username);
                 userPasswordField.SendKeys(password);
                 loginButton1.Click();
                 Thread.Sleep(5000);

             }*/

        }
    }
}
