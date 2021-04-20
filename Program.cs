// GoogleSheet Class: an instance of this class 
// has one public method (GenerateObjectDataAndUpdateGoogleSheet())which would retrive data
// from the dsi server, encapsilate the data and 
// update the sales google spreadsheet
//***********************************************************************************************
// To run this code: de-comment the main method.
// To run the dll: Create an instance & call the method GenerateObjectDataAndUpdateGoogleSheet().
//***********************************************************************************************
// Code Explanation
// *Variables name in plural are arrays
// *Methods name starts with a capital letter
// *Global variables are in all Caps
// //Comment: Code commented out
// // Comment: Developper comments
// Author: doth55

// Imports
using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4;
using System.Data.SqlClient;
using Google.Apis.Services;
using System.Threading;
using System.Data;
using System.IO;
using System;

namespace SalesGoogleSheet
{
    public class GoogleSheet
    {
        // Variable Declaration
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        string[] scopes = { SheetsService.Scope.Spreadsheets };
        string application_name;
        SheetsService service;
        readonly string sheet;
        string google_sheet_ID;

        // This list of Entry is used to store rows of data obtained from the server
        List<Entry> entries;

        // List of employees sorted by location
        List<Employee> ATH_employees;
        List<Employee> AUG_employees;

        //
        List<String> months;
        List<String> qtrs;

        //
        int current_day;
        int current_month;
        int current_Qtr;
        int current_year;
        int prior_year;
        
        // Constructor
        public GoogleSheet()
        {
            service = AuthorizeGoogleApp();
            application_name = "";
            google_sheet_ID = "";
            sheet = "";

            qtrs = new List<string>();
            entries = new List<Entry>();

            months = new List<string>();
            ATH_employees = new List<Employee>();
            AUG_employees = new List<Employee>();
            FillMonthsQtrs();

            // Employees should be able to see their previous months 
            // numbers until the 5th day of the new month.
            current_day = DateTime.Today.Day;
            if (current_day > 5)
            {
                current_month = DateTime.Today.Month;
            }
            else
            {
                current_month = DateTime.Today.Month - 1;
            }
            current_Qtr = (current_month + 2) / 3;
            current_year = DateTime.Today.Year;
            prior_year = current_year - 1;

        } // GoogleSheet

        public static void Main()
        {
            // Retrieve Data
            //Console.WriteLine("This is C#");
            //*********************************************
            GoogleSheet a = new GoogleSheet();
            a.GenerateObjectDataAndUpdateGoogleSheet();
            
            //*********************************************+
            // The next line keeps the console open for the developper to be 
            // able to read outputs.

            //Console.Read();
        }

        // Given by .NET Quickstart
        private SheetsService AuthorizeGoogleApp()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                // Console.WriteLine("Credential file saved to: " + credPath);
            } // using

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = application_name,
            });
            return service;
        } // AuthorizeGoogleApp

        // This method retrieve data from the server
        // & stores them in the list entries
        //private void GenerateData()
        private void GenerateData()
        {
            using (SqlConnection conn = new SqlConnection(""))
            {
                // 1.  Create a command object identifying the stored procedure
                conn.Open();
                SqlCommand cmd = new SqlCommand("", conn);

                // 2. Set the command object so it knows to execute a stored procedure
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 250;

                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    // Iterate through results & adding them to entries
                    while (rdr.Read())
                    {
                        entries.Add(new SalesGoogleSheet.Entry(Convert.ToInt32(rdr["SalesYear"]), Convert.ToInt32(rdr["SalesMonth"]),
                            Convert.ToString(rdr["Salesperson"]), Convert.ToDouble(rdr["TotalMonthlySales"]), Convert.ToInt32(rdr["Month_Units"])));
                    } // while
                } // using
            } // using
            GetEmployeeList();
        } // GenerateData

        // This method encapsulates data received from GenerateData
        // into Object Data. It also determines the ranges where data would 
        // go on the sheet. At last, it updates the spreadsheet.
        //private void GenerateObjectDataAndUpdateGoogleSheet()
        public void GenerateObjectDataAndUpdateGoogleSheet()
        {
            GenerateData();
            // Declare variables
            string range;
            IList<Object> obj = new List<Object>();
            List<IList<Object>> objData = new List<IList<Object>>();

            // Add months to ObjData starting 
            // with the current month and adds the rest 
            // in the second loop
            //Console.WriteLine("Month List: " + monthList.Count);
            //Console.Read();
            for(int i = (current_month * 4) - 4; i < months.Count; i++)
            {
                //Console.WriteLine(i);
                obj.Add(months[i]);
            } // for
            for (int i = 0; i < (current_month * 4) - 4; i++)
            { 
                obj.Add(months[i]);
            } // for

            // Update month headers
            range = "C1:AX1";
            objData.Add(obj);
            UpdateGoogleSheet(objData, google_sheet_ID, range, service);
            objData = new List<IList<Object>>();

            // Update ATH Month Section
            foreach (Employee i in ATH_employees)
            {
                obj = new List<Object>();
                obj.Add(i.GetName());
                for (int j = current_month; j <= 12; j++)
                {
                    obj.Add(i.GetPriorYearMonthlySales(j));
                    obj.Add(i.GetPriorYearMonthlyUnitsSold(j));
                    obj.Add(i.GetCurrYearMonthlySales(j));
                    obj.Add(i.GetCurrYearMonthlyUnitsSold(j));
                } // for
                for (int j = 1; j < current_month; j++)
                {
                    obj.Add(i.GetPriorYearMonthlySales(j));
                    obj.Add(i.GetPriorYearMonthlyUnitsSold(j));
                    obj.Add(i.GetCurrYearMonthlySales(j));
                    obj.Add(i.GetCurrYearMonthlyUnitsSold(j));
                } // for
                objData.Add(obj);
            } // foreach

            // Update Numbers
            // Update ATH_employees
            range = "B3:AX9";
            UpdateGoogleSheet(objData, google_sheet_ID, range, service);
            objData = new List<IList<Object>>();


            // Update AUG Month Section
            foreach (Employee i in AUG_employees)
            {
                obj = new List<Object>();
                obj.Add(i.GetName());
                for (int j = current_month; j <= 12; j++)
                {
                    obj.Add(i.GetPriorYearMonthlySales(j));
                    obj.Add(i.GetPriorYearMonthlyUnitsSold(j));
                    obj.Add(i.GetCurrYearMonthlySales(j));
                    obj.Add(i.GetCurrYearMonthlyUnitsSold(j));
                } // for

                for (int j = 1; j < current_month; j++)
                {
                    obj.Add(i.GetPriorYearMonthlySales(j));
                    obj.Add(i.GetPriorYearMonthlyUnitsSold(j));
                    obj.Add(i.GetCurrYearMonthlySales(j));
                    obj.Add(i.GetCurrYearMonthlyUnitsSold(j));
                } // for
                objData.Add(obj);
            } // foreach

            // Update Google Sheet
            // Update AUG_employees
            range = "B11:AX16";
            UpdateGoogleSheet(objData, google_sheet_ID, range, service);
            objData = new List<IList<Object>>();
            obj = new List<Object>();

            // Update Qtr Headers
            for (int i = (current_Qtr * 4) - 4; i < qtrs.Count; i++)
            {
                obj.Add(qtrs[i]);
            }
            for (int i = 0; i < (current_Qtr * 4) - 4; i++)
            {
                obj.Add(qtrs[i]);
            }

            // Update Google Sheet
            // Update quaterly header
            range = "C17:R17";
            objData.Add(obj);
            UpdateGoogleSheet(objData, google_sheet_ID, range, service);
            objData = new List<IList<Object>>();

            // Update ATH Qtr Section
            foreach (Employee i in ATH_employees)
            {
                obj = new List<Object>();
                obj.Add(i.GetName());
                for (int j = current_Qtr; j <= 4; j++)
                {
                    obj.Add(i.GetPriorYearQuarterlySales(j));
                    obj.Add(i.GetPriorYearQuarterlyUnitsSold(j));
                    obj.Add(i.GetCurrYearQuaterlySales(j));
                    obj.Add(i.GetCurrYearQuaterlyUnitsSold(j));
                } // for
                for (int j = 1; j < current_Qtr; j++)
                {
                    obj.Add(i.GetPriorYearQuarterlySales(j));
                    obj.Add(i.GetPriorYearQuarterlyUnitsSold(j));
                    obj.Add(i.GetCurrYearQuaterlySales(j));
                    obj.Add(i.GetCurrYearQuaterlyUnitsSold(j));
                } // for
                objData.Add(obj);
            } // foreach

            // Update ATH_employees quaters
            range = "B19:R25";
            UpdateGoogleSheet(objData, google_sheet_ID, range, service);
            objData = new List<IList<Object>>();

            // Update AUG Qtr Section
            foreach (Employee i in AUG_employees)
            {
                obj = new List<Object>();
                obj.Add(i.GetName());
                for (int j = current_Qtr; j <= 4; j++)
                {
                    obj.Add(i.GetPriorYearQuarterlySales(j));
                    obj.Add(i.GetPriorYearQuarterlyUnitsSold(j));
                    obj.Add(i.GetCurrYearQuaterlySales(j));
                    obj.Add(i.GetCurrYearQuaterlyUnitsSold(j));
                } // for

                for (int j = 1; j < current_Qtr; j++)
                {
                    obj.Add(i.GetPriorYearQuarterlySales(j));
                    obj.Add(i.GetPriorYearQuarterlyUnitsSold(j));
                    obj.Add(i.GetCurrYearQuaterlySales(j));
                    obj.Add(i.GetCurrYearQuaterlyUnitsSold(j));
                } // for
                objData.Add(obj);
            } // foreach

            // Update Google Sheet
            // Update AUG_employees quaters
            range = "B27:R32";
            UpdateGoogleSheet(objData, google_sheet_ID, range, service);
        } // GenerateObjectData

        // Headers 
        private void FillMonthsQtrs()
        {
            //monthList = new List<string>();
            months.Add("Jan Last Sales"); months.Add("Jan Last Units"); months.Add("Jan Curr Sales"); months.Add("Jan Curr Units");
            months.Add("Feb Last Sales"); months.Add("Feb Last Units"); months.Add("Feb Curr Sales"); months.Add("Feb Curr Units");
            months.Add("Mar Last Sales"); months.Add("Mar Last Units"); months.Add("Mar Curr Sales"); months.Add("Mar Curr Units");
            months.Add("Apr Last Sales"); months.Add("Apr Last Units"); months.Add("Apr Curr Sales"); months.Add("Apr Curr Units");
            months.Add("May Last Sales"); months.Add("May Last Units"); months.Add("May Curr Sales"); months.Add("May Curr Units");
            months.Add("Jun Last Sales"); months.Add("Jun Last Units"); months.Add("Jun Curr Sales"); months.Add("Jun Curr Units");
            months.Add("Jul Last Sales"); months.Add("Jul Last Units"); months.Add("Jul Curr Sales"); months.Add("Jul Curr Units");
            months.Add("Aug Last Sales"); months.Add("Aug Last Units"); months.Add("Aug Curr Sales"); months.Add("Aug Curr Units");
            months.Add("Sept Last Sales"); months.Add("Sept Last Units"); months.Add("Sept Curr Sales"); months.Add("Sept Curr Units");
            months.Add("Oct Last Sales"); months.Add("Oct Last Units"); months.Add("Oct Curr Sales"); months.Add("Oct Curr Units");
            months.Add("Nov Last Sales"); months.Add("Nov Last Units"); months.Add("Nov Curr Sales"); months.Add("Nov Curr Units");
            months.Add("Dec Last Sales"); months.Add("Dec Last Units"); months.Add("Dec Curr Sales"); months.Add("Dec Curr Units");

            //qtrList = new List<string>();
            qtrs.Add("Jan-Mar Last Sales"); qtrs.Add("Jan-Mar Last Units");
            qtrs.Add("Jan-Mar Curr Sales"); qtrs.Add("Jan-Mar Curr Units");
            qtrs.Add("Apr-Jun Last Sales"); qtrs.Add("Apr-Jun Last Units");
            qtrs.Add("Apr-Jun Curr Sales"); qtrs.Add("Apr-Jun Curr Units");
            qtrs.Add("Jul-Sept Last Sales"); qtrs.Add("Jul-Sept Last Units");
            qtrs.Add("Jul-Sept Curr Sales"); qtrs.Add("Jul-Sept Curr Units");
            qtrs.Add("Oct-Dec Last Sales"); qtrs.Add("Oct-Dec Last Units");
            qtrs.Add("Oct-Dec Curr Sales"); qtrs.Add("Oct-Dec Curr Units");
        } // fillMonthsQtrs

        // This method uses an helper method to read employee's name from the google spreadsheet("SalesNumbers")
        // After doing so, the helper method calls the GetEmployeesData() which retrieves employee information
        // using entries (entries) obtained from the server
        private void GetEmployeeList()
        {
            ReadEntries(3, ATH_employees);
            ReadEntries(11, AUG_employees);
        }

        // Read employee's name from spreadsheet
        // Used to determine current employees
        // This method is useful in preventing new years situations where
        // salesmen have not had the chance make any sale
        private void ReadEntries(Int32 index, List<Employee> employees)
        {
            var range = $"{sheet}!B{index}:C";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(google_sheet_ID, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    try
                    {
                        //Console.WriteLine("{0} | {1}", row[0], row[1]);
                        employees.Add(new Employee((String)row[0]));
                    }
                    catch (Exception e)
                    {
                        break;
                    }
                }
            }
            GetEmployeesData(employees);
        }

        // Update each employee with information regarding their prior &
        // current sales
        private void GetEmployeesData(List<Employee> employees)
        {
            foreach (Employee employee in employees)
            {
                foreach (Entry entry in entries)
                {
                    if (employee.GetName().Equals(entry.GetName()) && prior_year.Equals(entry.GetYear()))
                    {
                        employee.SetPriorYearMonthlySales(entry.GetMonth(), entry.GetMonthlySales());
                        employee.SetPriorYearMonthlyUnitsSold(entry.GetMonth(), entry.GetMonthlyUnitsSold());
                        
                    } else if (employee.GetName().Equals(entry.GetName()) && current_year.Equals(entry.GetYear()))
                    {
                        employee.SetCurrYearMonthlySales(entry.GetMonth(), entry.GetMonthlySales());
                        employee.SetCurrYearMonthlyUnitsSold(entry.GetMonth(), entry.GetMonthlyUnitsSold()); 
                    }
                }
            }
        }

        // This method updates the google sheet
        private void UpdateGoogleSheet(IList<IList<Object>> values, string spread_sheet_Id, string new_range, SheetsService service)
        {
            // Console.WriteLine("Updating Google Sheets....");
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(new ValueRange()
            { Values = values }, google_sheet_ID, new_range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
        } // UpdateGoogleSheet
    } // Program
} // TestGoogleSheets