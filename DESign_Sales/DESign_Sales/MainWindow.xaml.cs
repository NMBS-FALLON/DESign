using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DESign_BASE_WPF;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace DESign_Sales
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExtractBlueBeamMarkups extractBlueBeamMarkups = new ExtractBlueBeamMarkups();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Job job = extractBlueBeamMarkups.JobFromBlueBeamMarkups();

            Excel.Application oXL = new Excel.Application();
            Excel.Workbooks workbooks = oXL.Workbooks;
            Excel.Workbook workbook = workbooks.Add();
            Excel.Worksheet oSheet = workbook.ActiveSheet;

            
            var listOfStrippedMarks = job.ListOfStrippedMarks;

            int row = 3;
            foreach (int strippedMark in listOfStrippedMarks)
            {
                int loadCounter = 0;
                int noteCounter = 0;
                foreach (Girder girder in job.Girders)
                {
                    if (girder.StrippedNumber == strippedMark)
                    {
                        oSheet.Cells[row, 2] = girder.Mark;
                        oSheet.Cells[row, 3] = girder.Quantity;
                        oSheet.Cells[row, 4] = girder.Description;
                        string[] lengthArray = Regex.Split(girder.strBaseLength, "-");
                        double lengthFt = Convert.ToDouble(lengthArray[0]);
                        double lengthIn = 0.0;
                        if (lengthArray.Length ==2 )
                        {
                            lengthIn = Convert.ToDouble(lengthArray[1]);
                        }
                        oSheet.Cells[row, 5] = lengthFt;
                        oSheet.Cells[row, 6] = lengthIn;
                        int loadRow = row;
                        foreach(string load in girder.Loads)
                        {
                            loadCounter++;

                            string[] loadArray = load.Split(new string[] { " ", "@", "=", "=", ">", "-" }, StringSplitOptions.RemoveEmptyEntries);
                            if (loadArray.Length >= 1)
                            {
                                oSheet.Cells[loadRow, 17] = loadArray[0];
                            }
                            if (loadArray.Length >= 2)
                            {
                                oSheet.Cells[loadRow, 18] = loadArray[1];
                            }
                            if (loadArray.Length >= 3)
                            {
                                oSheet.Cells[loadRow, 19] = loadArray[2];
                            }
                            if (loadArray.Length >= 4)
                            {
                                oSheet.Cells[loadRow, 20] = loadArray[3];
                            }
                            if (loadArray.Length >= 5)
                            {
                                oSheet.Cells[loadRow, 21] = loadArray[4];
                            }
                            if (loadArray.Length >= 6)
                            {
                                oSheet.Cells[loadRow, 22] = loadArray[5];
                            }
                            if (loadArray.Length >= 7)
                            {
                                oSheet.Cells[loadRow, 23] = loadArray[6];
                            }
                            if (loadArray.Length >= 8)
                            {
                                oSheet.Cells[loadRow, 24] = loadArray[7];
                            }
                            if (loadArray.Length >= 9)
                            {
                                oSheet.Cells[loadRow, 25] = loadArray[8];
                            }
                            loadRow++;
                        }

                        int noteRow = row;
                        foreach (string note in girder.Notes)
                        {
                            noteCounter++;
                            oSheet.Cells[noteRow, 27] = note;
                            
                            noteRow++;
                        }

                        goto EndLoop;
                    }
                }
                foreach (Joist joist in job.Joists)
                {
                    if (joist.StrippedNumber == strippedMark)
                    {
                        oSheet.Cells[row, 2] = joist.Mark;
                        oSheet.Cells[row, 3] = joist.Quantity;
                        oSheet.Cells[row, 4] = joist.Description;
                        string[] lengthArray = Regex.Split(joist.strBaseLength, "-");
                        double lengthFt = Convert.ToDouble(lengthArray[0]);
                        double lengthIn = 0.0;
                        if (lengthArray.Length == 2)
                        {
                            lengthIn = Convert.ToDouble(lengthArray[1]);
                        }
                        oSheet.Cells[row, 5] = lengthFt;
                        oSheet.Cells[row, 6] = lengthIn;
                        int loadRow = row;
                        foreach (string load in joist.Loads)
                        {
                            loadCounter++;

                            string[] loadArray = load.Split(new string[] { " ", "@", "=", "=", ">", "-" }, StringSplitOptions.RemoveEmptyEntries);
                            if (loadArray.Length >= 1)
                            {
                                oSheet.Cells[loadRow, 17] = loadArray[0];
                            }
                            if (loadArray.Length >= 2)
                            {
                                oSheet.Cells[loadRow, 18] = loadArray[1];
                            }
                            if (loadArray.Length >= 3)
                            {
                                oSheet.Cells[loadRow, 19] = loadArray[2];
                            }
                            if (loadArray.Length >= 4)
                            {
                                oSheet.Cells[loadRow, 20] = loadArray[3];
                            }
                            if (loadArray.Length >= 5)
                            {
                                oSheet.Cells[loadRow, 21] = loadArray[4];
                            }
                            if (loadArray.Length >= 6)
                            {
                                oSheet.Cells[loadRow, 22] = loadArray[5];
                            }
                            if (loadArray.Length >= 7)
                            {
                                oSheet.Cells[loadRow, 23] = loadArray[6];
                            }
                            if (loadArray.Length >= 8)
                            {
                                oSheet.Cells[loadRow, 24] = loadArray[7];
                            }
                            if (loadArray.Length >= 9)
                            {
                                oSheet.Cells[loadRow, 25] = loadArray[8];
                            }
                            loadRow++;
                        }

                        int noteRow = row;
                        foreach (string note in joist.Notes)
                        {
                            noteCounter++;
                            oSheet.Cells[noteRow, 27] = note;

                            noteRow++;
                        }

                        goto EndLoop;
                    }
                }

                EndLoop:
                row = row + Math.Max(Math.Max(loadCounter, noteCounter), 1);

            }

            oXL.Visible = true;
            Marshal.ReleaseComObject(oSheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(oXL);
            GC.Collect();
        }

    }
}