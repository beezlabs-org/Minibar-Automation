using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Linq;
using System.Threading;
using System.Collections.Generic;
using System.IO;


namespace Beezlabs.RPA.Bots





{

    public static class ExtendedMethod
    {
        public static void Rename(this FileInfo fileInfo, string newName)
        {
            try
            {
                fileInfo.MoveTo(fileInfo.Directory.FullName + "\\" + newName);
            }
            catch
            {
                throw new Exception();
            }
        }
    }




    public class CsvdownloadFromMinibar : RPABotTemplate
    {
        private BotExecutionModel botExecutionModel;
        string SiteUrl = null;
        string path = null;
        string StartDate = null;
        string EndDate = null;
        string downloadFilepath = null;
        string BotPath = null;
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
          //  System.Diagnostics.Debugger.Launch();
          //  System.Diagnostics.Debugger.Break();
            LogMessage(this.GetType().FullName, $"Bot Execution Started successfully");
            this.botExecutionModel = botExecutionModel;
            try
            {
                InputsReceiver();
                WeekelySalesReportDownload();
                RenameCsv();
                Success("Bot executed sucessfully");
            }
            catch (Exception ex)
            {
                Failure($"Failed Due to {ex.Message} {ex.StackTrace}");
            }

        }

        private void RenameCsv()
        {
            try
            {
                Thread.Sleep(2000);
                string strPath = path;
                string[] files = Directory.GetFiles(path);
                string filenamess = null;
              
                string NewfileName = "Bacardi_weekly_sales" + "_" + StartDate + "_" + EndDate + ".csv";
                foreach (string filea in files)
                    filenamess = Path.GetFileName(filea);
                strPath = strPath + "\\" + filenamess;
                
                FileInfo file = new FileInfo(strPath);

                file.Rename(NewfileName);

                 BotPath = Path.Combine(path,NewfileName);

                string btpath = path + "\\" + NewfileName;
               
                AddVariable("Filepath", btpath);

            }
            catch (Exception ex)
            {
                Failure($"Failed Due to {ex.Message} {ex.StackTrace}");
            }
        }

        private void InputsReceiver()
        {

            try
            {
                SiteUrl = botExecutionModel.proposedBotInputs["siteURL"].value.ToString();
                //  SiteUrl = "https://app.periscopedata.com/shared/d18a9d06-3b0e-49eb-b598-c6a16129b91d?";
                LogMessage(this.GetType().FullName, $"Fetched the site URL");
                LogMessage(this.GetType().FullName, $"Getting the working directory path");
                path = GetWorkingDirectory();
                downloadFilepath = path;
               
                LogMessage(this.GetType().FullName, $"Fetched the working directory file");
                StartDate = botExecutionModel.proposedBotInputs["StartDateofminibar"].value.ToString();

                EndDate = botExecutionModel.proposedBotInputs["endDateofminibar"].value.ToString();

                LogMessage(this.GetType().FullName, $"Initilizing the chrome driver started");
            }

            catch (Exception ex)
            {
                throw new Exception($"Exception occured {ex.Message},{ex.StackTrace}");
            }



        }

        public void WeekelySalesReportDownload()
        {
            try
            {

                Dictionary<String, Object> chromePrefs = new Dictionary<String, Object>();

                chromePrefs.Add("download.default_directory", downloadFilepath);


                ChromeOptions chromeOptions = new ChromeOptions();
                LogMessage(this.GetType().FullName, $"Chrome options initialised");


                chromeOptions.AddAdditionalOption("pref", chromePrefs);

                chromeOptions.AddUserProfilePreference("download.default_directory", path);
                chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
                chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                var watch = new System.Diagnostics.Stopwatch();
                watch.Start();
                using (IWebDriver driver = new ChromeDriver(chromeOptions))
                {
                    driver.Manage().Window.Maximize();
                    driver.Navigate().GoToUrl(SiteUrl);
                    LogMessage(this.GetType().FullName, $"Url passed to chrome");
                    Thread.Sleep(1000);
                    driver.FindElement(By.ClassName("dashboard-settings-frame")).Click();

                    Thread.Sleep(1000);

                    WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 120));

                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div[1]/div[3]/div[4]/div[2]/div/div[1]/div[3]/div[5]/div/div[2]/div[1]/div/div[1]"))).Click();


                    WebElement SDate = (WebElement)driver.FindElement(By.XPath("/html/body/div[1]/div[3]/div[4]/div[2]/div/div[1]/div[3]/div[5]/div/div[2]/div[2]/div[1]/input"));

                    Thread.Sleep(500);

                    SDate.SendKeys(StartDate);

                    WebElement EDate = (WebElement)driver.FindElement(By.XPath("/html/body/div[1]/div[3]/div[4]/div[2]/div/div[1]/div[3]/div[5]/div/div[2]/div[2]/div[2]/input"));
                    Thread.Sleep(500);
                    EDate.SendKeys(EndDate);
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div[3]/div[4]/div[2]/div/div[1]/div[5]/div[4]/div"))).Click();

                    Thread.Sleep(10000);

                    Thread.Sleep(5000);
                    
                    IWebElement ele = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div[1]/div[3]/div[4]/div[3]/div[3]/div[1]/div[27]/div[2]")));

                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ele);
                    //download the itemized sale's in that portal



                    var elementOfitemizedsales = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div[1]/div[3]/div[4]/div[3]/div[3]/div[1]/div[27]/div[2]/div[2]")));

                    Actions action = new Actions(driver);

                    action.MoveToElement(elementOfitemizedsales).Perform();


                    IWebElement menuelementofsalesReport = driver.FindElement(By.XPath("//*[@id='body']/div[1]/div[3]/div[4]/div[3]/div[3]/div[1]/div[27]/div[2]/div[4]/div[2]"));
                    Thread.Sleep(500);
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuelementofsalesReport);




                    //   List<IWebElement> productsInfoList = driver.FindElements(By.XPath("//div[@class='xtuZKo']")).ToList();

                    Thread.Sleep(1500);


                    IWebElement reportDownloadButton = driver.FindElement(By.XPath("//*[@id='body']/div[1]/div[3]/div[4]/div[3]/div[3]/div[1]/div[27]/div[2]/div[4]/div[2]/div/div[3]"));
                    Thread.Sleep(500);
                    reportDownloadButton.Click();
                    Thread.Sleep(1000);



                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Exception occured {ex.Message},{ex.StackTrace}");
            }

        }
    }



}

