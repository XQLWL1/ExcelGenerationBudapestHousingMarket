using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace BudapestHousingMarket
{
    public partial class Form1 : Form
    {

        List<Flat> Flats;
        RealEstateEntities contex = new RealEstateEntities();

        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();

        }

        private void CreateExcel()
        {
            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet létrehozása
                xlWB = xlApp.Workbooks.Add();

                // Új munkalap beillesztése az excelbe
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable();


                //felhasználók számára elérhetővé válik
                xlApp.Visible = true;

                //Tiltani és engedélyezni tudjuk az excel alkamazásnak a felhasználó általi vezérlését.
                xlApp.UserControl = true;
            }
            catch (Exception exception)
            {

                string errorMsg = string.Format("Error: {0}\nLine: {1}", exception.Message, exception.Source);
                MessageBox.Show(errorMsg);

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }

        }

        private void CreateTable()
        {
            string[] headers = new string[]
            {
                "Kód",
                "Eladó",
                "Oldal",
                "Kerület",
                "Lift",
                "Szobák száma",
                "Alapterület (m2)",
                "Ár (mFt)",
                "Négyzetméter ár (Ft/m2)"
            };

            for (int i = 0; i < headers.Length; i++)

                xlSheet.Cells[1, i+1] = headers[i];

            object[,] values = new object[Flats.Count, headers.Length];
            int counter = 0;
            foreach (Flat item in Flats)
            {
                values[counter, 0] = item.Code;
                values[counter, 1] = item.Vendor;
                values[counter, 2] = item.Side;
                values[counter, 3] = item.District;
                values[counter, 4] = item.Elevator;
                values[counter, 5] = item.NumberOfRooms;
                values[counter, 6] = item.FloorArea;
                values[counter, 7] = item.Price;

                values[counter, 8] = "=" + GetCell(counter + 2, 8) + "*1000000/" + GetCell(counter + 2, 7);
                counter++;
            }

            xlSheet.get_Range
                (
                    GetCell(2, 1),
                    GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;


            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;

            //fejléc: középre rendezés függőlegesen és vizszintesen is
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            headerRange.EntireColumn.AutoFit();

            //nagyíthajuk:
            headerRange.RowHeight = 60;

            //szinezhetjük
            headerRange.Interior.Color = Color.Salmon;

            //körbe keretezzük a fejlécet
            headerRange.BorderAround2(Excel.XlLineStyle.xlDashDotDot, Excel.XlBorderWeight.xlThick);



        }

        private void LoadData()
        {
            Flats = contex.Flats.ToList();
        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }


    }
}
