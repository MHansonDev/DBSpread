using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

public class DBaseSpread
{
    // Global DB Variables
    private SqlConnection dbConn;
    public string userID;
    private string userPass;
    private string server;
    private string trusted;
    private string dbase;
    private int timeout;

    // Global Excel Variables
    private Excel.Application xlApp;
    private Excel.Workbook xlWorkbook;
    private Excel.Worksheet xlWorksheet;

    public DBaseSpread()
    {
        // Set Database Credentials
        userID = "User-PC\\User;";
        userPass = ";";
        server = "localhost;";
        trusted = "yes;";
        dbase = "AdventureWorksLT2008;";
        timeout = 30;

        // Create Database Connection
        dbConn = new SqlConnection("user id=" + userID +
                                   "password=" + userPass +
                                   "server=" + server +
                                   "Trusted_Connection=" + trusted +
                                   "database=" + dbase +
                                   "connection timeout=" + timeout.ToString());
        
        // Connect To Database
        try
        {
            dbConn.Open();
        }
        catch (Exception e)
        {
            Console.WriteLine(e.ToString());
        }

        // Create Excel Spreadsheet
        xlApp = new Excel.Application();
        xlWorkbook = xlApp.Workbooks.Add(Missing.Value);
        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];

        // Set Spreadsheet Headings
        string[] headings = {"Name", "Number", "Color", "Cost"};
        for (int i = 1; i <= headings.Length; i++)
        {
            xlWorksheet.Cells[1, i] = headings[i - 1];
            Excel.Range headRange = xlWorksheet.Cells[1, i];
            headRange.Font.Bold = true;
        }

        // Pull/Write Product Data
        try
        {
            SqlDataReader productReader = null;
            SqlCommand getProducts = new SqlCommand("SELECT *" +
                                                    "FROM SalesLT.Product", dbConn);
            productReader = getProducts.ExecuteReader();
            int rowCount = 2;
            while (productReader.Read())
            {
                // Add Product Name
                xlWorksheet.Cells[rowCount, 1] = productReader["Name"].ToString();
                // Add Product Number
                xlWorksheet.Cells[rowCount, 2] = productReader["ProductNumber"].ToString();
                // Add Product Color
                xlWorksheet.Cells[rowCount, 3] = productReader["Color"].ToString();
                // Add Product Cost
                xlWorksheet.Cells[rowCount, 4] = productReader["StandardCost"].ToString();
                rowCount++;
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e.ToString());
        }

        // Save Spreadsheet
        xlWorksheet.SaveAs("sample", Excel.XlFileFormat.xlWorkbookDefault);
        xlApp.Workbooks.Close();
        xlApp.Quit();

    }
    
    public static void Main()
    {
        DBaseSpread dbSpr = new DBaseSpread();
    }
}