using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;

using Microsoft.Crm.Sdk.Samples;

namespace schedulerautomatedtesting
{
    class Program
    {

        private static Guid moduserGuid = new Guid("");
        private static string FROM = "";
        private static string FROMNAME = "CRM Online User";
        private static string TO = "";
        private static string SMTP_USERNAME = "";
        private static string SMTP_PASSWORD = "";
        private static string CONFIGSET = "ConfigSet";
        private static string HOST = "smtp.office365.com";
        private static int PORT = 587;
        // private static string username = "";
        // private static string password = "";
        // private static string url = "";
        private static string username = "";
        private static string password = "";
        private static string url = "";
        //private static string url = "";
        private static string path = @"C:\Users\kyle.reinertson\source\repos\TryThis.csv";
        private static string html = "C:\\Users\\kyle.reinertson\\source\\repos\\CRMConsoleActivities\\CRMConsoleActivities\\HTMLPage1.html";
        private static string muralPath = "";
        private static string O365loginButtonId = "idSIButton9";
        private static string O365usernameId = "i0116";
        private static string O365passwordId = "i0118";
        public static DateTime y = DateTime.Now;
        //public static string appointmentId = ("### TEST ###" + " " + y.ToString("yyyy-MM-dd-HH-mm-ss"));

        //string connectionString = ConnectionStrings();

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
                organizationService = new OrganizationServiceProxy(new Uri(url), null, clientCredentials, null);

                if (organizationService != null)
                {
                    Guid gOrgId = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).OrganizationId;
                    if (gOrgId != Guid.Empty)
                    {

                        Console.WriteLine("***Connection Established Successfully...***");

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
            options.AddArguments("--headless");

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService();
            service.FirefoxBinaryPath = @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe";
            FirefoxDriver driver = new FirefoxDriver(options);

            // driver.Navigate().GoToUrl("");

            driver.Navigate().GoToUrl("");
            IWebElement query = driver.FindElement(By.Id("ContentPlaceHolder_desiredServiceRadio_0"));
            query.SendKeys(username);
            query = driver.FindElement(By.Id("ContentPlaceHolder_desiredServiceRadio_0"));
            query.Click();

            IWebElement query1 = driver.FindElement(By.Id("ContentPlaceHolder_customerZIPCodeTextBox"));
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator1 = By.Id("ContentPlaceHolder_customerZIPCodeTextBox");
            wait1.Until(drv => By.Id("ContentPlaceHolder_customerZIPCodeTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_customerZIPCodeTextBox")).SendKeys("85712");
            query1.Click();

            IWebElement query2 = driver.FindElement(By.Id("ContentPlaceHolder_getTimeSlotsBtn"));
            WebDriverWait wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator2 = By.Id("ContentPlaceHolder_getTimeSlotsBtn");
            wait2.Until(drv => By.Id("ContentPlaceHolder_getTimeSlotsBtn"));
            query2.Click();

            IWebElement query3 = driver.FindElement(By.Id("ContentPlaceHolder_timeZonesDropDown"));
            WebDriverWait wait3 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator3 = By.Id("ContentPlaceHolder_timeZonesDropDown");
            wait3.Until(drv => By.Id("ContentPlaceHolder_timeZonesDropDown"));
            query3.Click();

            driver.FindElement(By.XPath(".//*[@id='ContentPlaceHolder_timeZonesDropDown']/option[5]")).Click();

            IList<IWebElement> all = driver.FindElements(By.ClassName("calendar_default_event"));
            IWebElement query4 = all[all.Count - 1];
            // driver.FindElementByClassName("calendar_default_event").Click();
            WebDriverWait wait4 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator4 = By.Id("calendar_default_event");
            wait4.Until(drv => By.Id("calendar_default_event"));
            query4.Click();

            IWebElement query5 = driver.FindElement(By.Id("ContentPlaceHolder_businessNameTextBox"));
            WebDriverWait wait5 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator5 = By.Id("ContentPlaceHolder_businessNameTextBox");
            wait5.Until(drv => By.Id("ContentPlaceHolder_businessNameTextBox"));
            //DateTime y = DateTime.Now;
            driver.FindElement(By.Id("ContentPlaceHolder_businessNameTextBox")).SendKeys("### TEST ###" + " " + y.ToString("yyyy-MM-dd-HH-mm-ss"));
            query5.Click();

            IWebElement query6 = driver.FindElement(By.Id("ContentPlaceHolder_firstNameTextBox"));
            WebDriverWait wait6 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator6 = By.Id("ContentPlaceHolder_firstNameTextBox");
            wait6.Until(drv => By.Id("ContentPlaceHolder_firstNameTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_firstNameTextBox")).SendKeys("### TEST ###" + " " + y.ToString("yyyy-MM-dd-HH-mm-ss"));
            query6.Click();

            IWebElement query7 = driver.FindElement(By.Id("ContentPlaceHolder_lastNameTextBox"));
            WebDriverWait wait7 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator7 = By.Id("ContentPlaceHolder_lastNameTextBox");
            wait7.Until(drv => By.Id("ContentPlaceHolder_lastNameTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_lastNameTextBox")).SendKeys("### TEST ###" + " " + y.ToString("yyyy-MM-dd-HH-mm-ss"));
            query7.Click();

            IWebElement query8 = driver.FindElement(By.Id("ContentPlaceHolder_emailTextBox"));
            WebDriverWait wait8 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator8 = By.Id("ContentPlaceHolder_emailTextBox");
            wait8.Until(drv => By.Id("ContentPlaceHolder_emailTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_emailTextBox")).SendKeys("kyle.reinertson@mural365.com");
            query8.Click();

            IWebElement query9 = driver.FindElement(By.Id("ContentPlaceHolder_contactPhoneNumberTextBox"));
            WebDriverWait wait9 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator9 = By.Id("ContentPlaceHolder_contactPhoneNumberTextBox");
            wait9.Until(drv => By.Id("ContentPlaceHolder_contactPhoneNumberTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_contactPhoneNumberTextBox")).SendKeys(y.ToString("yyyMMddHHmmss"));
            query9.Click();

            IWebElement query10 = driver.FindElement(By.Id("ContentPlaceHolder_btnTextBox"));
            WebDriverWait wait10 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator10 = By.Id("ContentPlaceHolder_btnTextBox");
            wait10.Until(drv => By.Id("ContentPlaceHolder_btnTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_btnTextBox")).SendKeys(y.ToString("yyyMMddHHmmss"));
            query10.Click();

            IWebElement query11 = driver.FindElement(By.Id("ContentPlaceHolder_agentNotesTextBox"));
            WebDriverWait wait11 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator11 = By.Id("ContentPlaceHolder_agentNotesTextBox");
            wait11.Until(drv => By.Id("ContentPlaceHolder_agentNotesTextBox"));
            driver.FindElement(By.Id("ContentPlaceHolder_agentNotesTextBox")).SendKeys("### TEST ###" + " " + y.ToString("yyyy-MM-dd-HH-mm-ss"));
            query11.Click();

            IWebElement query12 = driver.FindElement(By.Id("ContentPlaceHolder_btnSubmit"));
            WebDriverWait wait12 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            By locator12 = By.Id("ContentPlaceHolder_btnSubmit");
            wait12.Until(drv => By.Id("ContentPlaceHolder_btnSubmit"));
            query12.Click();



            IOrganizationService serviceCRM = ConnecttoCRM();

            ConditionExpression condition = new ConditionExpression
            {
                AttributeName = "subject",
                Operator = ConditionOperator.Equal,
            };
            condition.Values.Add(("### TEST ###" + " " + y.ToString("yyyy-MM-dd-HH-mm-ss")).ToString());

            ColumnSet columns = new ColumnSet(true);

            QueryExpression expression = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "serviceappointment"
            };
            expression.Criteria.AddCondition(condition);
            EntityCollection result1 = serviceCRM.RetrieveMultiple(expression);

            foreach (var r in result1.Entities)
            {

                Console.WriteLine(r.Attributes["subject"]);
                SetStateRequest setState = new SetStateRequest();
                setState.EntityMoniker = r.ToEntityReference();
                setState.State = new OptionSetValue();
                setState.State.Value = 1;
                setState.Status = new OptionSetValue();
                setState.Status.Value = 8;
                SetStateResponse setStateResponse = (SetStateResponse)serviceCRM.Execute(setState);
                Console.WriteLine("ServiceActivityClosed");
                // Console.ReadLine();


            };

        }

    }

}
