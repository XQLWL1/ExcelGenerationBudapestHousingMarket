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

namespace BudapestHousingMarket
{
    public partial class Form1 : Form
    {
        RealEstateEntities contex = new RealEstateEntities();
        List<Flat> Flats;

        //A Microsoft Excel alkalmazás
        Excel.Application xlApp;

        //A létrehozott munkafüzet
        Excel.Workbook xlWB;

        //Munkalap a munkafüzeten belül
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
                Excel.Application xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // Új munkalap
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                //CreateTable();


                //felhasználók számára elérhetővé válik
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception)
            {

                string errorMsg = string.Format("Error: {0}\nLine: {1}", exception.Message, exception.Source);
                MessageBox.Show(errorMsg, "Error");

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
            {
                xlSheet.Cells[1, i] = headers[i-1];

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
                    values[counter, 8] = "";
                    counter++;
                }
            }
        }

        private void LoadData()
        {
            List<Flat> Flats;
            
            /*RealEstateEntities context = new RealEstateEntities();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Flats = contex.Flats.ToList();*/
        }


    }
}
