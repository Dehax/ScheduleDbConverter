using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using System.ComponentModel;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace ScheduleNewDbToPdf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static BaseFont arialFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arial.ttf", "cp1251", BaseFont.EMBEDDED);
        private static BaseFont arialbFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arialbd.ttf", "cp1251", BaseFont.EMBEDDED);
        private static BaseFont arialbiFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arialbi.ttf", "cp1251", BaseFont.EMBEDDED);

        private float PAGE_MARGIN_LEFT = CentimetersToPoints(0.5f);
        private float PAGE_MARGIN_RIGHT = CentimetersToPoints(0.5f);
        private float PAGE_MARGIN_TOP = CentimetersToPoints(0.5f);
        private float PAGE_MARGIN_BOTTOM = CentimetersToPoints(0.5f);

        private OpenFileDialog openDbFileDialog;
        private SaveFileDialog savePdfFileDialog;
        private BackgroundWorker exportBackgroundWorker;

        public MainWindow()
        {
            InitializeComponent();

            this.openDbFileDialog = new OpenFileDialog();
            this.openDbFileDialog.Filter = "Schedule SQLite database file|*.db|All files|*.*";
            
            this.savePdfFileDialog = new SaveFileDialog();
            this.savePdfFileDialog.AddExtension = true;
            this.savePdfFileDialog.DefaultExt = "*.pdf";
            this.savePdfFileDialog.Filter = "Schedule PDF file|*.pdf";

            this.exportBackgroundWorker = new BackgroundWorker();
            this.exportBackgroundWorker.WorkerReportsProgress = true;
            this.exportBackgroundWorker.DoWork += new DoWorkEventHandler(this.exportBackgroundWorker_DoWork);
            this.exportBackgroundWorker.ProgressChanged += new ProgressChangedEventHandler(this.exportBackgroundWorker_ProgressChanged);
            this.exportBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.exportBackgroundWorker_RunWorkerCompleted);
            
            //exportBackgroundWorker.RunWorkerAsync(new string[] { @"C:\Users\Dehax\OneDrive\Shedule_PI-13a.db", @"C:\Users\Dehax\OneDrive\Shedule_PI-13a.pdf" });
        }

        private void chooseButton_Click(object sender, RoutedEventArgs e)
        {
            if (openDbFileDialog.ShowDialog(this).Value)
            {
                inputPathTextBox.Text = openDbFileDialog.FileName;

                savePdfFileDialog.FileName = openDbFileDialog.FileName.Substring(0, openDbFileDialog.FileName.Length - 3);

                if (savePdfFileDialog.ShowDialog(this).Value)
                {
                    exportBackgroundWorker.RunWorkerAsync(new string[] { openDbFileDialog.FileName, savePdfFileDialog.FileName });
                }
                else
                {
                    inputPathTextBox.Clear();
                }
            }
            else
            {
                inputPathTextBox.Clear();
            }
        }

        private void exportBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker backWorker = sender as BackgroundWorker;
            string[] arguments = (string[])e.Argument;
            convertDatabase(arguments[0], arguments[1], backWorker);
        }

        private void exportBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.exportProgressBar.Value = e.ProgressPercentage;
        }

        private void exportBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(this, e.Error.Message, "Error during export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show(this, "Export was canceled", "Canceled", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                this.exportProgressBar.Value = 0;
                MessageBox.Show(this, "Schedule successfully converted to \"" + savePdfFileDialog.FileName + "\"", "Success!", MessageBoxButton.OK, MessageBoxImage.Information);
                System.Diagnostics.Process.Start(savePdfFileDialog.FileName);
                //System.Diagnostics.Process.Start(@"C:\Users\Dehax\OneDrive\Shedule_PI-13a.pdf");
            }
            
            //this.Close();
        }

        /// <summary>
        /// Converts Расписашка (*.db) database file with schedule to Microsoft Office Publisher 2010 (*.pub) file
        /// </summary>
        /// <param name="inPath">full path to Расписашка (*.db) database file with schedule</param>
        /// <param name="outPath">full path to Публикация Publisher 2010 (*.pub) publication file with converted schedule</param>
        /// <param name="backWorker">BackgroundWorker for reporting progress</param>
        /// <returns>full path to converted Microsoft Office Publisher 2010 (*.pub) file</returns>
        private void convertDatabase(string inPath, string outPath, BackgroundWorker backWorker)
        {
            //string savePath = inPath.Substring(0, inPath.Length - 3) + ".pdf";
            // TODO: Database format no longer supported.
            SQLiteConnection dbConnection = new SQLiteConnection("Data Source=" + inPath + ";Version=3");
            dbConnection.Open();
            const string command = "SELECT names_1.name AS teacher_short, names_1.full_name AS teacher, lessons._id, lessons.day, lessons.number, names.name AS subject, names_2.name AS place, names_3.name AS type, lessons.weeks, lessons.kind_id AS type_id" +
                                   " FROM lessons, names, names names_1, names names_2, names names_3" +
                                   " WHERE lessons.name_id = names._id AND lessons.teacher_id = names_1._id AND lessons.place_id = names_2._id AND lessons.kind_id = names_3._id";
            SQLiteCommand dbCommand = new SQLiteCommand(command, dbConnection);
            SQLiteDataReader dbReader = dbCommand.ExecuteReader();

            Document pdfDoc = new Document(PageSize.A4.Rotate(), PAGE_MARGIN_LEFT, PAGE_MARGIN_RIGHT, PAGE_MARGIN_TOP, PAGE_MARGIN_BOTTOM);
            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, new FileStream(outPath, FileMode.Create));
            

            pdfDoc.Open();

            //PdfContentByte canvas = pdfWriter.DirectContent;

            // Create tables
            PdfPTable daysTable = new PdfPTable(3);
            daysTable.WidthPercentage = 100.0f;
            PdfPTable[] tables = new PdfPTable[10];
            string[] weekDays = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница" };

            for (int i = 0; i < 10; i++)
            {
                PdfPTable table = new PdfPTable(5);
                
                // Setup rows/cols sizes
                //table.SetTotalWidth(new float[] { 0.42f, 0.43f, 3.62f, 1.38f, 3.38f });
                table.SetTotalWidth(new float[] { CentimetersToPoints(0.42f), CentimetersToPoints(0.43f), CentimetersToPoints(3.62f), CentimetersToPoints(1.38f), CentimetersToPoints(3.38f) });

                // Header with weekday name
                PdfPCell cell = new PdfPCell(new Phrase(weekDays[i % 5], new Font(arialbiFont, 10.0f)));
                cell.Colspan = 5;
                cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                cell.FixedHeight = CentimetersToPoints(0.63f);
                table.AddCell(cell);

                for (int row = 0; row < 5; row++)
                {
                    // Is lecture
                    cell = new PdfPCell();
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.FixedHeight = CentimetersToPoints(0.63f);
                    cell.Padding = CentimetersToPoints(0.1016f);
                    //cell.BorderWidth = 0.5f;
                    table.AddCell(cell);
                    // # number
                    cell = new PdfPCell(new Phrase((row + 1).ToString(), new Font(arialbFont, 10.0f)));
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.FixedHeight = CentimetersToPoints(0.63f);
                    cell.Padding = CentimetersToPoints(0.1016f);
                    //cell.BorderWidth = 0.5f;
                    table.AddCell(cell);
                    // subject title
                    cell = new PdfPCell();
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.FixedHeight = CentimetersToPoints(0.63f);
                    cell.Padding = CentimetersToPoints(0.1016f);
                    //cell.BorderWidth = 0.5f;
                    table.AddCell(cell);
                    // place
                    cell = new PdfPCell();
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.FixedHeight = CentimetersToPoints(0.63f);
                    cell.Padding = CentimetersToPoints(0.1016f);
                    //cell.BorderWidth = 0.5f;
                    table.AddCell(cell);
                    // teacher
                    cell = new PdfPCell();
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.FixedHeight = CentimetersToPoints(0.63f);
                    cell.Padding = CentimetersToPoints(0.1016f);
                    //cell.BorderWidth = 0.5f;
                    table.AddCell(cell);
                }
                //table.CompleteRow();

                tables[i] = table;

                backWorker.ReportProgress((int)(i / 10.0f * 75));
            }

            daysTable.AddCell(tables[0]);
            daysTable.AddCell(tables[1]);
            daysTable.AddCell(tables[2]);
            daysTable.AddCell(tables[3]);
            daysTable.AddCell(tables[4]);
            daysTable.AddCell(new PdfPCell());
            daysTable.AddCell(tables[5]);
            daysTable.AddCell(tables[6]);
            daysTable.AddCell(tables[7]);
            daysTable.AddCell(tables[8]);
            daysTable.AddCell(tables[9]);
            daysTable.AddCell(new PdfPCell());

            for (int rowIndex = 0; rowIndex < daysTable.Rows.Count; rowIndex++)
            {
                PdfPRow row = (PdfPRow)daysTable.Rows[rowIndex];
                //foreach (PdfPCell cell in row.GetCells())
                //{
                //    cell.Border = PdfPCell.NO_BORDER;
                //    cell.PaddingLeft = 0.0f;
                //    cell.PaddingTop = 0.0f;
                //    cell.PaddingRight = CentimetersToPoints(0.5f);
                //    cell.PaddingBottom = CentimetersToPoints(0.5f);
                //}
                PdfPCell[] cells = row.GetCells();
                for (int cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                {
                    cells[cellIndex].Border = PdfPCell.NO_BORDER;
                    cells[cellIndex].PaddingLeft = 0.0f;
                    cells[cellIndex].PaddingTop = 0.0f;
                    
                    if (cellIndex != 2)
                    {
                        cells[cellIndex].PaddingRight = CentimetersToPoints(0.5f);
                    }
                    else
                    {
                        cells[cellIndex].PaddingRight = 0.0f;
                    }

                    if (rowIndex != 1)
                    {
                        cells[cellIndex].PaddingBottom = CentimetersToPoints(0.5f);
                    }
                    else
                    {
                        cells[cellIndex].PaddingBottom = CentimetersToPoints(1.5f);
                    }
                }
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

                    tables[index].GetRow((int)number).GetCells()[0].Phrase = new Phrase(type_id == 1 ? "+" : "", new Font(arialbFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[2].Phrase = new Phrase(subject, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[3].Phrase = new Phrase(place, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[4].Phrase = new Phrase(teacher_short, new Font(arialFont, 10.0f));

                    index = (int)day + 4;

                    tables[index].GetRow((int)number).GetCells()[0].Phrase = new Phrase(type_id == 1 ? "+" : "", new Font(arialbFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[2].Phrase = new Phrase(subject, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[3].Phrase = new Phrase(place, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[4].Phrase = new Phrase(teacher_short, new Font(arialFont, 10.0f));
                }
                else if (week == "e") // even: 2,4,6,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    index = (int)day + 4;

                    tables[index].GetRow((int)number).GetCells()[0].Phrase = new Phrase(type_id == 1 ? "+" : "", new Font(arialbFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[2].Phrase = new Phrase(subject, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[3].Phrase = new Phrase(place, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[4].Phrase = new Phrase(teacher_short, new Font(arialFont, 10.0f));
                }
                else if (week == "o") // odd: 1,3,5,...
                {
                    day = (long)dbReader["day"];
                    number = (long)dbReader["number"];

                    index = (int)day - 1;

                    tables[index].GetRow((int)number).GetCells()[0].Phrase = new Phrase(type_id == 1 ? "+" : "", new Font(arialbFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[2].Phrase = new Phrase(subject, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[3].Phrase = new Phrase(place, new Font(arialFont, 10.0f));
                    tables[index].GetRow((int)number).GetCells()[4].Phrase = new Phrase(teacher_short, new Font(arialFont, 10.0f));
                }

                backWorker.ReportProgress((int)(++counter / (float)dbReader.StepCount * 24 + 75));
            }

            //for (int tableIndex = 0; tableIndex < tables.Length; tableIndex++)
            //{
            //    tables[tableIndex].WriteSelectedRows(0, -1, CentimetersToPoints(0.5f) + tableIndex * CentimetersToPoints(9.75f), CentimetersToPoints(20.5f), canvas);
            //}

            pdfDoc.Add(daysTable);
            pdfDoc.Close();
            backWorker.ReportProgress(100);
        }

        private static float CentimetersToPoints(float cm)
        {
            return cm * 720 / 25.4f;
        }
    }
}
