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

        // Create Excel Spreadsheet;
        xlApp = new Excel.Application();
        xlWorkbook = xlApp.Workbooks.Add(Missing.Value);
        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];

        // Get Product Data
        try
        {
            SqlDataReader productReader = null;
            SqlCommand getProducts = new SqlCommand("SELECT *" +
                                                    "FROM SalesLT.Product", dbConn);
            productReader = getProducts.ExecuteReader();
            int rowCount = 1;
            while (productReader.Read())
            {
                xlWorksheet.Cells[rowCount, 1] = productReader["Name"].ToString();
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