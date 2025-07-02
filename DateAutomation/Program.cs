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

// This program generates a list of dates into an Excel file, starting from a specified date, and incrementing by one week.

namespace DateAutomation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // CUSTOMIZABLE VARIABLES

            Console.WriteLine("Enter the number of weeks you want to generate dates for: ");
            int desiredWeeks = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the starting day (1-31): ");
            int startingDay = int.Parse(Console.ReadLine());
            Console.WriteLine("Enter the starting month (1-12): "); 
            int startingMonth = int.Parse(Console.ReadLine());
            Console.WriteLine("Enter the starting year: ");
            int startingYear = int.Parse(Console.ReadLine());

            string exePath = Assembly.GetExecutingAssembly().Location;
            string exeDirectory = Path.GetDirectoryName(exePath);
            string desiredFileLocation = Path.Combine(exeDirectory, "AutomatedDates.xlsx");
            
            // (double slashes may be required in the path)
            //"C:\\Users\\aidan\\Documents\\MicrosoftOffice\\Excel\\AutomatedDates.xlsx"; //Change this to the desired path and file name of schedule. 
            // (double slashes may be required in the path)

            // END OF CUSTOMIZABLE VARIABLES


            List<int> ThirtyOneDayMonths = new List<int>();
            List<int> ThirtyDayMonths = new List<int>();

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
            
            int day = startingDay;
            int month = startingMonth;
            int year = startingYear;
            int remainingDays = 0;
            bool changedDay = false;
            int leapYearCount = startingYear % 4; // Initialize leap year count based on the starting year
            if (leapYearCount == 0) // If the starting year is a leap year, set count to 4
            {
                leapYearCount = 4;
            }

            int futureMonth = startingMonth;
            int futureDay = startingDay + 6;
            int futureYear = startingYear;
            int remainingFutureDays = 0;
            bool changedFutureDay = false;
            int futureLeapYearCount = leapYearCount;

            // Future Date Variables Initialization
            if (futureMonth == 12)
            {
                if (futureDay > 31)
                {
                    remainingFutureDays = futureDay - 31;
                    futureDay = remainingFutureDays;
                    futureMonth = 1;
                    futureYear++;
                    futureLeapYearCount++;
                    if (futureLeapYearCount == 5)
                    {
                        futureLeapYearCount = 1;
                    }
                }
            }
            else if (futureMonth == 2)
            {
                if (futureLeapYearCount != 4 && futureDay > 28)
                {
                    remainingFutureDays = futureDay - 28;
                    futureDay = remainingFutureDays;
                    futureMonth++;
                }
                else if (futureLeapYearCount == 4 && futureDay > 29)
                {
                    remainingFutureDays = futureDay - 29;
                    futureDay = remainingFutureDays;
                    futureMonth++;
                }
            }
            else if (ThirtyDayMonths.Contains(month) && futureDay > 30)
            {
                remainingFutureDays = futureDay - 30;
                futureDay = remainingFutureDays;
                futureMonth++;
            }
            else if (ThirtyOneDayMonths.Contains(month) && futureDay > 31)
            {
                remainingFutureDays = futureDay - 31;
                futureDay = remainingFutureDays;
                futureMonth++;
            }
            // End of Future Date Variables Initialization
                

            Spreadsheet document = new Spreadsheet();
            Worksheet sheet = document.Workbook.Worksheets.Add("Schedule");

            //




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
                        leapYearCount++ ;
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
            if (File.Exists(desiredFileLocation)) // checks if previous version exists
            {
                File.Delete(desiredFileLocation); // deletes old version
            }
            document.SaveAs(desiredFileLocation); // saves document @ desired location
            Process.Start(desiredFileLocation);
        }
    }
}
