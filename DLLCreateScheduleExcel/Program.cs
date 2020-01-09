using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using DLLCreateScheduleExcel.Data;
using DLLCreateScheduleExcel.Entities;
using ClosedXML.Excel;
using DLLCreateScheduleExcel.Services;

namespace DLLCreateScheduleExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Creating excel file...");
            List<ColumnsGrid> columnsGrid = new List<ColumnsGrid>();

            var interval = new IntervalDates(DateTime.Parse("31/01/2020"), DateTime.Parse("01/04/2020"));
            var intervalDates = new TimelineRange();
            List<string> daysList = intervalDates.TimelineDaysList(interval);
            //foreach(var r in res) Console.WriteLine(r);

            List<string> monthsList = intervalDates.TimeLineMonthsList(interval);
            //foreach (var r in resu) Console.WriteLine(r);
            
            var db = new ConnectionDataBase();
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Cronograma");

            //Report Title
            ws.Cell("A1").Value = "Cronograma";
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell("A1").Style.Font.Bold = true;

            ws.Range("A1:G3").Merge();


            //Report Header
            ws.Cell("A4").Value = "Item";
            ws.Cell("B4").Value = "Nome da Atividade";
            ws.Cell("C4").Value = "Responsável";
            ws.Cell("D4").Value = "Data Início";
            ws.Cell("E4").Value = "Data Fim";
            ws.Cell("F4").Value = "Dias Úteis";
            ws.Cell("G4").Value = "Realizado";

            //Report Body of the Grid

            //Report Body of the Timeline
            int count = 8;
            int posMonth = 8;
            int indexMonthPrevious = 0;
            foreach (var d in daysList)
            {
                string[] month = new string[13] { "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" };

                var date = DateTime.Parse(d);
                int indexMonth = (int) date.Month;
                if(indexMonth != indexMonthPrevious)
                {
                    List<string> currentMonth = monthsList.FindAll(months => months == month[indexMonth]);
                    int colSpan = currentMonth.Count();
                    ws.Cell(3, posMonth).Value = month[indexMonth];
                    ws.Range(3, posMonth, 3, (posMonth + colSpan)).Merge();
                    posMonth += colSpan;
                    indexMonthPrevious = indexMonth;
                }

                ws.Cell(4, count).Value = d;           
                count++;
            }


            //Filters and Create table
            var range = ws.Range("A4:G10");
            range.CreateTable();
             
            //Fix the column size with column content 
            ws.Columns("1-"+count).AdjustToContents();
            ws.SheetView.FreezeColumns(7);
            
            //Salve file
            wb.SaveAs(@"C:\Users\murilo.paz.REDESPC\Desktop\Murilo\myProjects\test.xlsx");

            //Release objects
           
            wb.Dispose();          
            


            Console.WriteLine("Finish");
            Console.ReadKey();

            
        }


    }
}
