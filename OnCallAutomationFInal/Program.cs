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

namespace OnCallAutomationFInal
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<int>  ThirtyOneDayMonths = new List<int>();
            List<int> ThirtyDayMonths = new List<int>();

            int day = 9;
            int month = 12;
            int year = 2024;
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
            Worksheet sheet = document.Workbook.Worksheets.Add("OnCallSched");

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
            sheet.Cell("B1").Value = "Name";

            //

            for (int i = 2; i < 2000; i++)
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
                        if(leapYearCount == 5)
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

            //

            if (File.Exists(@"C:\Users\aidan\Documents\MicrosoftOffice\Excel\MockOnCallSheet.xlsx"))
            {
                File.Delete(@"C:\Users\aidan\Documents\MicrosoftOffice\Excel\MockOnCallSheet.xlsx");
            }
            document.SaveAs(@"C:\Users\aidan\Documents\MicrosoftOffice\Excel\MockOnCallSheet.xlsx");
            document.Close();
            Process.Start(@"C:\Users\aidan\Documents\MicrosoftOffice\Excel\MockOnCallSheet.xlsx");
        }
    }
}
