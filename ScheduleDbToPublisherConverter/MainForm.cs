using Microsoft.Office.Core;
using System;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Publisher = Microsoft.Office.Interop.Publisher;
using System.ComponentModel;

namespace ScheduleDbToPublisherConverter
{
    public partial class MainForm : Form
    {
        private const float PAGE_A4_WIDTH = 29.7f;
        private const float PAGE_A4_HEIGHT = 21.0f;

        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonChooseFile_Click(object sender, EventArgs e)
        {
            if (openDbFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                textBoxPath.Text = openDbFileDialog.FileName;

                savePubFileDialog.FileName = openDbFileDialog.FileName.Substring(0, openDbFileDialog.FileName.Length - 3);

                if (savePubFileDialog.ShowDialog(this) == DialogResult.OK)
                {
                    exportBackgroundWorker.RunWorkerAsync(new string[] { openDbFileDialog.FileName, savePubFileDialog.FileName });
                }
                else
                {
                    textBoxPath.ResetText();
                }
            }
            else
            {
                textBoxPath.ResetText();
            }
        }
        /// <summary>
        /// Converts Расписашка (*.db) database file with schedule to Microsoft Office Publisher 2010 (*.pub) file
        /// </summary>
        /// <param name="inPath">full path to Расписашка (*.db) database file with schedule</param>
        /// <param name="outPath">full path to Публикация Publisher 2010 (*.pub) publication file with converted schedule</param>
        /// <param name="backWorker">BackgroundWorker for reporting progress</param>
        /// <returns>full path to converted Microsoft Office Publisher 2010 (*.pub) file</returns>
        private string convertDatabase(string inPath, string outPath, BackgroundWorker backWorker)
        {
            string savePath = inPath.Substring(0, inPath.Length - 3) + ".pub";

            SQLiteConnection dbConnection = new SQLiteConnection("Data Source=" + inPath + ";Version=3");
            dbConnection.Open();
            const string command = "SELECT names_1.name AS teacher_short, names_1.full_name AS teacher, lessons._id, lessons.day, lessons.number, names.name AS subject, names_2.name AS place, names_3.name AS type, lessons.weeks, lessons.kind_id AS type_id" +
                                   " FROM lessons, names, names names_1, names names_2, names names_3" +
                                   " WHERE lessons.name_id = names._id AND lessons.teacher_id = names_1._id AND lessons.place_id = names_2._id AND lessons.kind_id = names_3._id";
            SQLiteCommand dbCommand = new SQLiteCommand(command, dbConnection);
            SQLiteDataReader dbReader = dbCommand.ExecuteReader();
            
            Publisher.Application publisherApp = new Publisher.Application();
            publisherApp.ScreenUpdating = false;
            Publisher.Document document = publisherApp.ActiveDocument;
            Publisher.Window publisherWindow = publisherApp.ActiveWindow;
            //publisherWindow.Caption = "Schedule";
            //publisherWindow.WindowState = Publisher.PbWindowState.pbWindowStateMaximize;
            //publisherWindow.Visible = false;

            // Setup design
            document.PageSetup.PageWidth = publisherApp.CentimetersToPoints(PAGE_A4_WIDTH);
            document.PageSetup.PageHeight = publisherApp.CentimetersToPoints(PAGE_A4_HEIGHT);
            

            //Publisher.Page page = document.Pages.Add(1, 0);
            Publisher.Page page = document.Pages[1];

            // Create tables
            Publisher.Shape[] shapes = new Publisher.Shape[10];
            Publisher.Table[] tables = new Publisher.Table[10];

            float[] positionsX = { 0.5f, 10.25f, 20.0f, 0.5f, 10.25f, 0.5f, 10.25f, 20.0f, 0.5f, 10.25f };
            float[] positionsY = { 0.5f, 0.5f, 0.5f, 5.6f, 5.6f, 10.75f, 10.75f, 10.75f, 15.85f, 15.85f };
            string[] weekDays = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница" };

            for (int i = 0; i < 10; i++)
            {
                Publisher.Shape shape = page.Shapes.AddTable(6, 5, publisherApp.CentimetersToPoints(positionsX[i]), publisherApp.CentimetersToPoints(positionsY[i]), publisherApp.CentimetersToPoints(9.25f), publisherApp.CentimetersToPoints(4.6f));
                Publisher.Table table = shape.Table;
                shapes[i] = shape;
                tables[i] = table;
                table.Rows[1].Cells.Merge();

                // Setup rows/cols sizes
                table.Columns[1].Width = publisherApp.CentimetersToPoints(0.425f);
                table.Columns[2].Width = publisherApp.CentimetersToPoints(0.425f);
                table.Columns[3].Width = publisherApp.CentimetersToPoints(3.625f);
                table.Columns[4].Width = publisherApp.CentimetersToPoints(1.375f);
                table.Columns[5].Width = publisherApp.CentimetersToPoints(3.4f);

                table.Rows[1].Height = publisherApp.CentimetersToPoints(0.63f);
                table.Rows[2].Height = publisherApp.CentimetersToPoints(0.62f);
                table.Rows[3].Height = publisherApp.CentimetersToPoints(0.62f);
                table.Rows[4].Height = publisherApp.CentimetersToPoints(0.62f);
                table.Rows[5].Height = publisherApp.CentimetersToPoints(0.62f);
                table.Rows[6].Height = publisherApp.CentimetersToPoints(1.45f);

                table.GrowToFitText = false;

                // Setup borders
                foreach (Publisher.Cell tcell in table.Cells)
                {
                    tcell.BorderLeft.Weight = 1.0f;
                    tcell.BorderTop.Weight = 1.0f;
                    tcell.BorderRight.Weight = 1.0f;
                    tcell.BorderBottom.Weight = 1.0f;
                }

                Publisher.Cell cell = table.Rows[1].Cells[1];
                cell.VerticalTextAlignment = Publisher.PbVerticalTextAlignmentType.pbVerticalTextAlignmentCenter;
                cell.BorderLeft.Weight = 1.0f;
                cell.BorderTop.Weight = 1.0f;
                cell.BorderRight.Weight = 1.0f;
                cell.BorderBottom.Weight = 1.0f;

                Publisher.TextRange cellText = cell.TextRange;
                cellText.Text = weekDays[i % 5];
                cellText.ParagraphFormat.Alignment = Publisher.PbParagraphAlignmentType.pbParagraphAlignmentCenter;
                cellText.Font.Name = "Arial";
                cellText.Font.Bold = MsoTriState.msoTrue;
                cellText.Font.Italic = MsoTriState.msoTrue;
                //cellText.ParagraphFormat.SpaceAfter = 0.0f;
                //cellText.ParagraphFormat.LineSpacing = 1.0f;
                //cellText.ParagraphFormat.SpaceBefore = 0.0f;

                for (int j = 0; j < 5; j++)
                {
                    cell = table.Rows[j + 2].Cells[2];
                    cell.VerticalTextAlignment = Publisher.PbVerticalTextAlignmentType.pbVerticalTextAlignmentCenter;
                    cellText = cell.TextRange;
                    cellText.ParagraphFormat.Alignment = Publisher.PbParagraphAlignmentType.pbParagraphAlignmentCenter;
                    cellText.Font.Name = "Arial";
                    cellText.Font.Bold = MsoTriState.msoTrue;
                    //cellText.ParagraphFormat.SpaceAfter = 0.0f;
                    //cellText.ParagraphFormat.LineSpacing = 1.0f;
                    //cellText.ParagraphFormat.SpaceBefore = 0.0f;
                    cellText.Text = (j + 1).ToString();
                }

                backWorker.ReportProgress((int)(i / 10.0f * 75));
            }

            // Fill data
            string subject, place, teacher_short;
            long type_id;
            string week;
            long day, number;
            int index;
            int counter = 0;

            while (dbReader.Read())
            {
                subject = (string)dbReader["subject"];
                place = (string)dbReader["place"];
                teacher_short = (string)dbReader["teacher_short"];
                type_id = (long)dbReader["type_id"];
                week = (string)dbReader["weeks"];

                if (week == "a") // all: 1,2,3,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    index = (int)day - 1;

                    tables[index].Rows[(int)number + 1].Cells[1].TextRange.Text = (type_id == 1 ? "+" : "");
                    tables[index].Rows[(int)number + 1].Cells[3].TextRange.Text = subject;
                    tables[index].Rows[(int)number + 1].Cells[4].TextRange.Text = place;
                    tables[index].Rows[(int)number + 1].Cells[5].TextRange.Text = teacher_short;

                    index = (int)day + 4;

                    tables[index].Rows[(int)number + 1].Cells[1].TextRange.Text = (type_id == 1 ? "+" : "");
                    tables[index].Rows[(int)number + 1].Cells[3].TextRange.Text = subject;
                    tables[index].Rows[(int)number + 1].Cells[4].TextRange.Text = place;
                    tables[index].Rows[(int)number + 1].Cells[5].TextRange.Text = teacher_short;
                }
                else if (week == "e") // even: 2,4,6,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    index = (int)day + 4;

                    tables[index].Rows[(int)number + 1].Cells[1].TextRange.Text = (type_id == 1 ? "+" : "");
                    tables[index].Rows[(int)number + 1].Cells[3].TextRange.Text = subject;
                    tables[index].Rows[(int)number + 1].Cells[4].TextRange.Text = place;
                    tables[index].Rows[(int)number + 1].Cells[5].TextRange.Text = teacher_short;
                }
                else if (week == "o") // odd: 1,3,5,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    index = (int)day - 1;

                    tables[index].Rows[(int)number + 1].Cells[1].TextRange.Text = (type_id == 1 ? "+" : "");
                    tables[index].Rows[(int)number + 1].Cells[3].TextRange.Text = subject;
                    tables[index].Rows[(int)number + 1].Cells[4].TextRange.Text = place;
                    tables[index].Rows[(int)number + 1].Cells[5].TextRange.Text = teacher_short;
                }

                backWorker.ReportProgress((int)(++counter / (float)dbReader.StepCount * 24 + 75));
            }

            publisherWindow.Caption = "Schedule";
            publisherWindow.Visible = true;
            publisherWindow.WindowState = Publisher.PbWindowState.pbWindowStateMaximize;
            publisherApp.ScreenUpdating = true;

            document.SaveAs(outPath);
            backWorker.ReportProgress(100);
            
            // Release all COM objects
            Marshal.ReleaseComObject(publisherApp);

            return savePath;
        }

        private void exportBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker backWorker = sender as BackgroundWorker;
            string[] arguments = (string[])e.Argument;
            e.Result = convertDatabase(arguments[0], arguments[1], backWorker);
        }

        private void exportBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.exportProgressBar.Value = e.ProgressPercentage;
        }

        private void exportBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(this, e.Error.Message, "Error during export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show(this, "Export was canceled", "Canceled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                this.exportProgressBar.Value = 0;
                MessageBox.Show(this, "Schedule successfully converted to \"" + savePubFileDialog.FileName + "\"", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
