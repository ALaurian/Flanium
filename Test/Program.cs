using System.Collections;
using System.Data;
using Flanium;
using OpenQA.Selenium;
using WinRT;



//Initializes the chrome driver and navigates to the website
var browser = Initializers.InitializeChrome(PageLoadStrategy.Eager);

browser.Navigate().GoToUrl("https://rpachallenge.com/");

//Download excel file
WebEvents.ActionJs.Click(browser,"//*[text()=' Download Excel ']");
Thread.Sleep(1000);
var excelFileName = Helpers.HandleDownloads(browser)[0];
var getExcelFile = @"C:\Users\" + Environment.UserName + @"\Downloads\" + excelFileName;

//Opens the excel file and converts it to a DataTable
var excelScope = new Blocks.ExcelEngine(getExcelFile);
var exceldataTable = excelScope.ToDataTable(1,1);
excelScope.Close();

//Clicks on the start button
var startButton = WebEvents.ActionJs.Click(browser,"//button[normalize-space()='Start']");

//Creates the rpaChallenge Workflow
var rpaChallengeWorkflow = new Blocks.Engine();

    rpaChallengeWorkflow.AddActions(new Func<object, object>[]
    {
        (dispatcher) => exceldataTable.Rows.Cast<DataRow>().GetEnumerator(),
        (next) =>
        {
            if (rpaChallengeWorkflow.GetOutput("dispatcher").As<IEnumerator>().MoveNext() == false)
            {
                return rpaChallengeWorkflow.Stop();
            }

            return true;
        }, 
        (getCurrentRow) => rpaChallengeWorkflow.GetOutput("dispatcher").As<IEnumerator>().Current,
        
        (firstName) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["First Name"],
        (lastName) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["Last Name "],
        (companyName) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["Company Name"],
        (roleInCompany) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["Role in Company"],
        (address) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["Address"],
        (email) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["Email"],
        (phoneNumber) => rpaChallengeWorkflow.GetOutput("getCurrentRow").As<DataRow>()["Phone Number"],
        (setFirstName) => WebEvents.ActionJs.SetValue(browser, "//*[text()='First Name']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("firstName").ToString()),
        (setLastName) => WebEvents.ActionJs.SetValue(browser, "//*[text()='Last Name']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("lastName").ToString()),
        (setCompanyName) => WebEvents.ActionJs.SetValue(browser, "//*[text()='Company Name']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("companyName").ToString()),
        (setRoleInCompany) => WebEvents.ActionJs.SetValue(browser, "//*[text()='Role in Company']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("roleInCompany").ToString()),
        (setAddress) => WebEvents.ActionJs.SetValue(browser, "//*[text()='Address']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("address").ToString()),
        (setEmail) => WebEvents.ActionJs.SetValue(browser, "//*[text()='Email']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("email").ToString()),
        (setPhoneNumber) => WebEvents.ActionJs.SetValue(browser, "//*[text()='Phone Number']//following-sibling::*[1]", rpaChallengeWorkflow.GetOutput("phoneNumber").ToString()),
        (clickSubmit) => WebEvents.ActionJs.Click(browser, "//input[@value='Submit'][1]"),
        
        (jumpToNext) => rpaChallengeWorkflow.JumpTo("next")
    }).Execute();

