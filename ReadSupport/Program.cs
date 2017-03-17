using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ReadSupport {

  class Program {

    static void Main( string[] args ) {


      // Prerequsites should be present: 
      // arg[0] should be commaseparated list of initials
      // inputfile and outputfile should be present
      string OKstring = prereqursiteOK( args );   
      if ( !( "" == OKstring )) {
        System.Console.WriteLine( OKstring );
        System.Console.ReadLine();
        return;
      }

      string[] stringSeparators = new string[] {","};

      string[] emplyeeInitials = args[0].ToString().Split( stringSeparators, StringSplitOptions.RemoveEmptyEntries );
      Array.ForEach<string>( emplyeeInitials, x => emplyeeInitials[Array.IndexOf<string>( emplyeeInitials, x )] = x.Trim() );

      string path = Directory.GetCurrentDirectory();
      string InputFullFileName = path + "\\Time_registrations.xls";
      string OutputFullFileName = path + "\\Time_registrations_Output.xls";
      string OtherString = "Others";
                                   
      string[] supportersAndOther = emplyeeInitials;
      Array.Resize( ref supportersAndOther, emplyeeInitials.Length + 1 );
      supportersAndOther[supportersAndOther.Length-1] = OtherString;

      List<Tuple<string, string, double>> AllTimeUsage;
      AllTimeUsage = new List<Tuple<string, string, double>>();
      string tempstring;
      string YearMonth = "";
      string EmployeeInitials = "";
      string YearMonthsInUse = "";
      Double EmployeeTimeUsage;

      System.Console.WriteLine( "" );
      System.Console.WriteLine( "Employees: " + args[0].ToString() );
      System.Console.WriteLine( "InputFullFileName: " + InputFullFileName );
      System.Console.WriteLine( "OutputFullFileName: " + OutputFullFileName );

      //Cleanup old temp output
      if (System.IO.File.Exists(OutputFullFileName+"Test"))
        System.IO.File.Delete(OutputFullFileName+"Test");

      Excel.Application xlApp = new Excel.Application();
      Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(InputFullFileName);
      Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
      Excel.Range xlRange = xlWorksheet.UsedRange;
      int rowCount = xlRange.Rows.Count;
      int colCount = xlRange.Columns.Count;
      int lineNo = 0;

      for ( int i = 1; i <= rowCount; i++ ) {

        if (ExcelDataIsOk(xlRange,i)) {

           // Dates
           tempstring = ( (DateTime) xlRange.Cells[i, 1].Value ).ToString( "yyyyMMddHHmmss" );
           YearMonth = tempstring.Substring( 0, 4 ) + "-" + tempstring.Substring( 4, 2 );

           // Initials
           tempstring = (string) ( xlRange.Cells[i, 6].Value );
           EmployeeInitials = emplyeeInitials.Contains(tempstring) ? tempstring : OtherString;
                             
           // Timeusage
           EmployeeTimeUsage = ( (double) xlRange.Cells[i, 7].Value );

           // Update set of toubles
           AllTimeUsage.Add( Tuple.Create( YearMonth, EmployeeInitials, EmployeeTimeUsage ) );

           // Get all dates
           if (!YearMonthsInUse.Contains(YearMonth))
             YearMonthsInUse = YearMonthsInUse + YearMonth + ";";                 
         }
      }

      xlWorkbook.Close( 0 );
      
      Excel.Workbook xlWorkbookOut = xlApp.Workbooks.Open( OutputFullFileName );
      Excel._Worksheet xlWorksheetOut = xlWorkbookOut.Sheets[1];
      
      // Header for the output:  blank and the supporters and 'Others'
      lineNo = 1; 
      tempstring="_;";
      foreach ( String thissupporter in supportersAndOther )
        tempstring += thissupporter+";";
      
      CreateRow( xlWorksheetOut, lineNo, tempstring );
      lineNo++;
      
      // Number of lines deterined by number of unique "year-month"-pairs
      // Each line is: "year-month";hours_for_Initials;hours_for_Initials;...;   
      foreach ( string date in ((string[])YearMonthsInUse.Split( ';' ))) {

        if ( date.Length > 0 ) {

          tempstring = date+";";
          foreach ( String thissupporter in supportersAndOther )
            tempstring += GetHours( date, thissupporter, AllTimeUsage ).ToString().Replace( "_", "." ) + ";";

          //System.Console.WriteLine( tempstring );
          CreateRow( xlWorksheetOut, lineNo, tempstring ); 
        }

        lineNo++;
      }

     for ( int rowclean = lineNo; rowclean <= 100; rowclean++ )
       ( (Excel.Range) xlWorksheetOut.Rows[rowclean, Missing.Value] ).Delete( Excel.XlDeleteShiftDirection.xlShiftUp );


      xlWorkbookOut.SaveAs( OutputFullFileName + "Test" );
      xlWorkbookOut.Close();
      xlApp.Quit(); 

      System.IO.File.Delete(OutputFullFileName);
      System.IO.File.Copy(OutputFullFileName+"Test",OutputFullFileName);
      System.IO.File.Delete(OutputFullFileName+"Test");


      System.Console.WriteLine( "Done" );
      System.Console.ReadKey();  
    }

    /// <summary>
    /// This finction checks: 
    /// 
    ///  One input argument  - should be comma separated list of employee initials
    ///  
    ///  Inputfile should be present - Time_registrations_Output.xls
    ///  Input format in input excelfile:
    ///  Sheet: Time_registrations
    ///  Columns must be:    1. Date	 2. Project	 3.Currency	 4.Task  	5.Comment  	6.Employee	7.Hours
    ///  Data in column 1, 6 and 7 is used.
    ///  
    ///  OutputFullFileName = Time_registrations_Output.xls" 
    /// </summary>
    /// <param name="args">arguments for main</param>
    /// <returns>Blank = OK, string with an error description if not</returns>
    private static string prereqursiteOK( string[] args ) {

      if (args.Length != 1)
        return "Exactly one argument should be given. \nThe argument shuld be a comma-separated initials list\n- matching timelog initials.";

      if ( args[0].ToString().Length<2 )
        return "Initials list length is less that 2 - the argument shuld be a comma-separated initials list\n- matching timelog initials.";

      string path = Directory.GetCurrentDirectory();

      if (!(File.Exists(path + "\\Time_registrations.xls")))
        return "Please copy inputfile named 'Time_registrations.xls' here:\n"+path;

      if (!(File.Exists(path + "\\Time_registrations_Output.xls")))
        return "Please copy inputfile named 'Time_registrations_Output.xls' here:\n"+path;

      return "";
    }
    
    private static bool ExcelDataIsOk( Excel.Range xlRange, int i ) {

      bool result = false;

      if ( ( xlRange.Cells[i, 1].Value != null ) &&
            ( xlRange.Cells[i, 6].Value != null ) &&
            ( xlRange.Cells[i, 7].Value != null ) ) {

        if ( ( xlRange.Cells[i, 1].Value is DateTime ) &&
            ( xlRange.Cells[i, 6].Value is String ) &&
            ( xlRange.Cells[i, 7].Value is Double ) ) {

          result = true;
        }
      }

      return result;

    }
    
    private static double GetHours( string date, string name, List<Tuple<string, string, double>> AllTimeUsage ) {

      double result =0;
      foreach (Tuple<string, string, double> elem in AllTimeUsage )
        if ((elem.Item1==date) & (elem.Item2 == name))
           result += elem.Item3;
      
      return result;

    }

    private static void CreateRow( Excel._Worksheet ws, int offset, string line ) {

      int cols = line.Length - line.Replace( ";", "" ).Length;
      string[] linecontent = line.Split( ';' );

      for ( int col = 1; col <= cols; col++ ) {
        Excel.Range cell = ws.Cells[offset, col]; // The cell to write to
        cell.Value = linecontent[col-1];
      }      
    }

  }
}
