using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Linq;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

// This program generates a list of dates excluding weekends in an Excel file, starting from a specified date.
// The user just needs to set a starting date, desired amount of weeks, and file location. 
// Then, run the DateAutomater.sln, press start, and use the generated dates however desired. 
// (If you run the program multiple times without changing file name, it will overwrite the previous file but you have to close it first)

namespace DateAutomation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // CHANGE THESE VARIABLES TO CUSTOMIZE THE START DATE, NUMBER OF WEEKS, FILE LOCATION, AND FILE NAME.

            int desiredWeeks = 100; // Change this to the number of weeks you want to generate dates for

            int startingDay = 16; // Change this to the day of the month you want to start from (enter a monday to follow a typical work week)
            int startingMonth = 6; // Change this to the number month you want to start from
            int startingYear = 2025; // Change this to the year you want to start from

            string desiredFileLocation = "C:\\Users\\aidan\\Documents\\MicrosoftOffice\\Excel\\AutomatedDates.xlsx"; //Change this to the desired path and file name of schedule. 
                                                                                                                     // (double slashes may be required in the path)

            // END OF CUSTOMIZABLE VARIABLES


            List<int> ThirtyOneDayMonths = new List<int>();
            List<int> ThirtyDayMonths = new List<int>();

            int day = startingDay;
            int month = startingMonth;
            int year = startingYear;
            int remainingDays = 0;
            bool changedDay = false;
            int leapYearCount = 4;

            int futureMonth = 12;
            int futureDay = 15;
            int futureYear = 2024;
            int remainingFutureDays = 0;
            bool changedFutureDay = false;
            int futureLeapYearCount = 4;


            Spreadsheet document = new Spreadsheet();
            Worksheet sheet = document.Workbook.Worksheets.Add("Schedule");

            //

            ThirtyOneDayMonths.Add(1);
            ThirtyOneDayMonths.Add(3);
            ThirtyOneDayMonths.Add(5);
            ThirtyOneDayMonths.Add(7);
            ThirtyOneDayMonths.Add(8);
            ThirtyOneDayMonths.Add(10);

            ThirtyDayMonths.Add(4);
            ThirtyDayMonths.Add(6);
            ThirtyDayMonths.Add(9);
            ThirtyDayMonths.Add(11);


            sheet.Cell("A1").Value = "Date";

            //

            for (int i = 2; i < desiredWeeks + 2; i++)
            {
                sheet.Cell("A" + i.ToString()).Value = month + "/" + day + "/" + year + " - " + futureMonth + "/" + futureDay + "/" + futureYear;


                /////////////////////////////////////////////Initial Date Logic //////////////////////////////////

                changedDay = false;

                if (ThirtyOneDayMonths.Contains(month) && changedDay == false)
                {

                    if (day + 7 > 31)
                    {
                        remainingDays = 31 - day;
                        day = 7 - remainingDays;
                        month++;
                        changedDay = true;
                    }
                    else
                    {
                        day = day + 7;
                        changedDay = true;
                    }
                }

                if (month == 12 && changedDay == false)
                {
                    if (day + 7 > 31)
                    {
                        remainingDays = 31 - day;
                        day = 7 - remainingDays;
                        month = 1;
                        year++;
                        leapYearCount++;
                        if (leapYearCount == 5)
                        {
                            leapYearCount = 1;
                        }
                        changedDay = true;
                    }
                    else
                    {
                        day = day + 7;
                        changedDay = true;
                    }
                }


                if (ThirtyDayMonths.Contains(month) && changedDay == false)
                {

                    if (day + 7 > 30)
                    {
                        remainingDays = 30 - day;
                        day = 7 - remainingDays;
                        month++;
                        changedDay = true;
                    }
                    else
                    {
                        day = day + 7;
                        changedDay = true;
                    }

                }

                if (month == 2 && leapYearCount != 4 && changedDay == false)
                {

                    if (day + 7 > 28)
                    {
                        remainingDays = 28 - day;
                        day = 7 - remainingDays;
                        month++;
                        changedDay = true;
                    }
                    else
                    {
                        day = day + 7;
                        changedDay = true;
                    }

                }

                if (month == 2 && leapYearCount == 4 && changedDay == false)
                {
                    if (day + 7 > 29)
                    {
                        remainingDays = 29 - day;
                        day = 7 - remainingDays;
                        month++;
                        changedDay = true;
                    }
                    else
                    {
                        day = day + 7;
                        changedDay = true;
                    }
                }

                //////////////////////////////////// Future Date Logic /////////////////////////
                changedFutureDay = false;

                if (ThirtyOneDayMonths.Contains(futureMonth) && changedFutureDay == false)
                {

                    if (futureDay + 7 > 31)
                    {
                        remainingFutureDays = 31 - futureDay;
                        futureDay = 7 - remainingFutureDays;
                        futureMonth++;
                        changedFutureDay = true;
                    }
                    else
                    {
                        futureDay = futureDay + 7;
                        changedFutureDay = true;
                    }
                }

                if (futureMonth == 12 && changedFutureDay == false)
                {
                    if (futureDay + 7 > 31)
                    {
                        remainingFutureDays = 31 - futureDay;
                        futureDay = 7 - remainingFutureDays;
                        futureMonth = 1;
                        futureYear++;
                        futureLeapYearCount++;
                        if (futureLeapYearCount == 5)
                        {
                            futureLeapYearCount = 1;
                        }
                        changedFutureDay = true;
                    }
                    else
                    {
                        futureDay = futureDay + 7;
                        changedFutureDay = true;
                    }
                }


                if (ThirtyDayMonths.Contains(futureMonth) && changedFutureDay == false)
                {

                    if (futureDay + 7 > 30)
                    {
                        remainingFutureDays = 30 - futureDay;
                        futureDay = 7 - remainingFutureDays;
                        futureMonth++;
                        changedFutureDay = true;
                    }
                    else
                    {
                        futureDay = futureDay + 7;
                        changedFutureDay = true;
                    }

                }

                if (futureMonth == 2 && futureLeapYearCount != 4 && changedFutureDay == false)
                {

                    if (futureDay + 7 > 28)
                    {
                        remainingFutureDays = 28 - futureDay;
                        futureDay = 7 - remainingFutureDays;
                        futureMonth++;
                        changedFutureDay = true;
                    }
                    else
                    {
                        futureDay = futureDay + 7;
                        changedFutureDay = true;
                    }

                }

                if (futureMonth == 2 && futureLeapYearCount == 4 && changedFutureDay == false)
                {
                    if (futureDay + 7 > 29)
                    {
                        remainingFutureDays = 29 - futureDay;
                        futureDay = 7 - remainingFutureDays;
                        futureMonth++;
                        changedFutureDay = true;
                    }
                    else
                    {
                        futureDay = futureDay + 7;
                        changedFutureDay = true;
                    }
                }
            }

            //////Saving Document to Excel File///////
            // Replace file location ("C:\Users\aidan\Documents\MicrosoftOffice\Excel\AutomatedDates.xlsx") with your own.

            if (File.Exists(desiredFileLocation)) // checks if previous version exists
            {
                //FileStream oldDoc = new FileStream(desiredFileLocation, FileMode.Open, FileAccess.ReadWrite);
                //oldDoc.Close(); // closes the file if it is open
                File.Delete(desiredFileLocation); // deletes old version
            }
            document.SaveAs(desiredFileLocation); // saves document in Excel Folder
            Process.Start(desiredFileLocation);
        }
    }
}
