using System.Data;
using Flanium;
using WebSurfer;
using WinRT;

//Initializes the chrome driver and navigates to the website
var browser = new WebBrowser(BrowserType.Chrome);
browser.NavigateTo("https://rpachallenge.com/");

//Download excel file
browser.JavaScript.Action.Click("//*[text()=' Download Excel ']");
Thread.Sleep(1000);
var excelFileName = browser.GetDownloadedFiles()[0];
var getExcelFile = @"C:\Users\" + Environment.UserName + @"\Downloads\" + excelFileName;

//Opens the excel file and converts it to a DataTable
var excelScope = new Blocks.ExcelEngine(getExcelFile);
var exceldataTable = excelScope.ToDataTable(1,1);
excelScope.Close();

//Clicks on the start button
var startButton = browser.JavaScript.Action.Click("//button[normalize-space()='Start']");

//Creates the rpaChallenge Workflow
var rpaChallengeWorkflow = new Blocks.Engine().SetDispatcher(exceldataTable.Rows.Cast<DataRow>().ToArray());

rpaChallengeWorkflow.AddActions(new Func<object, object>[]
{
    (dispatcher) => rpaChallengeWorkflow.GetCurrent(),

    (firstName) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["First Name"],
    (lastName) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["Last Name "],
    (companyName) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["Company Name"],
    (roleInCompany) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["Role in Company"],
    (address) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["Address"],
    (email) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["Email"],
    (phoneNumber) => rpaChallengeWorkflow["dispatcher"].As<DataRow>()["Phone Number"],
    (setFirstName) => browser.JavaScript.Action.SetValue( "//*[text()='First Name']//following-sibling::*[1]", rpaChallengeWorkflow["firstName"].ToString()),
    (setLastName) => browser.JavaScript.Action.SetValue( "//*[text()='Last Name']//following-sibling::*[1]", rpaChallengeWorkflow["lastName"].ToString()),
    (setCompanyName) => browser.JavaScript.Action.SetValue( "//*[text()='Company Name']//following-sibling::*[1]", rpaChallengeWorkflow["companyName"].ToString()),
    (setRoleInCompany) => browser.JavaScript.Action.SetValue( "//*[text()='Role in Company']//following-sibling::*[1]", rpaChallengeWorkflow["roleInCompany"].ToString()),
    (setAddress) => browser.JavaScript.Action.SetValue( "//*[text()='Address']//following-sibling::*[1]", rpaChallengeWorkflow["address"].ToString()),
    (setEmail) => browser.JavaScript.Action.SetValue( "//*[text()='Email']//following-sibling::*[1]", rpaChallengeWorkflow["email"].ToString()),
    (setPhoneNumber) => browser.JavaScript.Action.SetValue( "//*[text()='Phone Number']//following-sibling::*[1]", rpaChallengeWorkflow["phoneNumber"].ToString()),
    (clickSubmit) => browser.JavaScript.Action.Click( "//input[@value='Submit'][1]"),

    (goNext) => rpaChallengeWorkflow.Next("dispatcher")
}).Execute();
