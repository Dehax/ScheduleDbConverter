using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ScheduleDbConverter
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonChooseFile_Click(object sender, EventArgs e)
        {
            if (openDbFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                textBoxPath.Text = openDbFileDialog.FileName;
                string savePath = convertDatabase(openDbFileDialog.FileName);
                MessageBox.Show(this, "Schedule successfully converted to \"" + savePath + "\"", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                textBoxPath.ResetText();
            }
        }
        /// <summary>
        /// Converts Расписашка (*.db) database file with schedule to Microsoft Office Excel 2010 (*.xlsx) workbook file
        /// </summary>
        /// <param name="path">full path to Расписашка (*.db) database file with schedule</param>
        /// <returns>full path to converted Microsoft Office Excel 2010 (*.xlsx) workbook file</returns>
        private string convertDatabase(string path)
        {
            SQLiteConnection dbConnection = new SQLiteConnection("Data Source=" + path + ";Version=3");
            dbConnection.Open();
            //SQLiteDataAdapter dbAdapter = new SQLiteDataAdapter("SELECT lessons.* FROM lessons", dbConnection);
            const string command = "SELECT names_1.name AS teacher_short, names_1.full_name AS teacher, lessons._id, lessons.day, lessons.number, names.name AS subject, names_2.name AS place, names_3.name AS type, lessons.weeks" +
                                   " FROM lessons, names, names names_1, names names_2, names names_3" +
                                   " WHERE lessons.name_id = names._id AND lessons.teacher_id = names_1._id AND lessons.place_id = names_2._id AND lessons.kind_id = names_3._id";
            SQLiteCommand dbCommand = new SQLiteCommand(command, dbConnection);
            SQLiteDataReader dbReader = dbCommand.ExecuteReader();

            Excel.Application excelApp = new Excel.Application();
            excelApp.SheetsInNewWorkbook = 1;
            //excelApp.Visible = true;
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook excelBook = workbooks.Add();
            Excel._Worksheet worksheet = excelApp.ActiveSheet;
            worksheet.Name = "Schedule";

            //int i = 1;

            //worksheet.Cells[i, "A"] = "teacher_short";
            //worksheet.Cells[i, "B"] = "teacher";
            //worksheet.Cells[i, "C"] = "_id";
            //worksheet.Cells[i, "D"] = "day";
            //worksheet.Cells[i, "E"] = "number";
            //worksheet.Cells[i, "F"] = "subject";
            //worksheet.Cells[i, "G"] = "place";
            //worksheet.Cells[i, "H"] = "type";
            //worksheet.Cells[i, "I"] = "weeks";

            //while (dbReader.Read())
            //{
            //    i++;
            //    worksheet.Cells[i, "A"] = dbReader["teacher_short"];
            //    worksheet.Cells[i, "B"] = dbReader["teacher"];
            //    worksheet.Cells[i, "C"] = dbReader["_id"];
            //    worksheet.Cells[i, "D"] = dbReader["day"];
            //    worksheet.Cells[i, "E"] = dbReader["number"];
            //    worksheet.Cells[i, "F"] = dbReader["subject"];
            //    worksheet.Cells[i, "G"] = dbReader["place"];
            //    worksheet.Cells[i, "H"] = dbReader["type"];
            //    worksheet.Cells[i, "I"] = dbReader["weeks"];
            //}

            string[] weeks = { "Верхняя неделя", "Нижняя неделя" };
            string[] days = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница" };

            for (int week = 1; week <= weeks.Length; week++)
            {
                worksheet.Cells[(week - 1) * days.Length * 5 + 1, "A"] = weeks[week - 1];

                for (int dayNum = 1; dayNum <= days.Length; dayNum++)
                {
                    worksheet.Cells[(week - 1) * days.Length * 5 + (dayNum - 1) * 5 + 1, "B"] = days[dayNum - 1];
                }
            }

            long row;
            long day, number;
            //StringBuilder data = new StringBuilder();
            while (dbReader.Read())
            {
                string week = (string)dbReader["weeks"];

                if (week == "a") // all: 1,2,3,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    row = 1 + 5 * 5 + (day - 1) * 5 + (number - 1);

                    worksheet.Cells[row, "C"] = number;
                    worksheet.Cells[row, "D"] = dbReader["subject"];
                    worksheet.Cells[row, "E"] = dbReader["place"];
                    worksheet.Cells[row, "F"] = dbReader["type"];
                    worksheet.Cells[row, "G"] = dbReader["teacher"];

                    row = 1 + (day - 1) * 5 + (number - 1);

                    worksheet.Cells[row, "C"] = number;
                    worksheet.Cells[row, "D"] = dbReader["subject"];
                    worksheet.Cells[row, "E"] = dbReader["place"];
                    worksheet.Cells[row, "F"] = dbReader["type"];
                    worksheet.Cells[row, "G"] = dbReader["teacher"];
                }
                else if (week == "e") // even: 2,4,6,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    row = 1 + 5 * 5 + (day - 1) * 5 + (number - 1);

                    worksheet.Cells[row, "C"] = number;
                    worksheet.Cells[row, "D"] = dbReader["subject"];
                    worksheet.Cells[row, "E"] = dbReader["place"];
                    worksheet.Cells[row, "F"] = dbReader["type"];
                    worksheet.Cells[row, "G"] = dbReader["teacher"];
                }
                else if (week == "o") // odd: 1,3,5,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    row = 1 + (day - 1) * 5 + (number - 1);

                    worksheet.Cells[row, "C"] = number;
                    worksheet.Cells[row, "D"] = dbReader["subject"];
                    worksheet.Cells[row, "E"] = dbReader["place"];
                    worksheet.Cells[row, "F"] = dbReader["type"];
                    worksheet.Cells[row, "G"] = dbReader["teacher"];
                }
            }

            worksheet.Columns.AutoFit();
            worksheet.Columns.NumberFormat = "@";

            string savePath = path.Substring(0, path.Length - 3) + ".xlsx";
            //worksheet.SaveAs(savePath);
            //worksheet.SaveAs("C:\\Programs\\schedule.xlsx");

            excelBook.Close(true, savePath);
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(excelBook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excelApp);

            return savePath;
        }
    }
}
