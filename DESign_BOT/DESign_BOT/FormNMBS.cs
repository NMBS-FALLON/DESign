using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using DESign_BOT;
using DESign_BASE;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DESign_Bot_FS_Tools;

namespace DESign_BOT
{
    
    public partial class FormNMBSHelper : Form
    {
        StringManipulation stringManipulation = new StringManipulation();
        ClassInsertBOMData classInsertBOMData = new ClassInsertBOMData();
        FolderOperations folderOperations = new FolderOperations();
        //JoistDetails joistDetails = new JoistDetails();
        public FormNMBSHelper()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
 
            
        }

        public void btnCreateNewBOM_Click(object sender, EventArgs e)
        {
            labelProgramState.Text = "Select file & hold; this could take several minutes";
            try { 
                  classInsertBOMData.createNMBSBOM3();
                  labelProgramState.Text = "Process complete";
                 
                }
            catch { labelProgramState.Text = "Sorry, there was an issue"; }
            

        }


        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            new FormNMBS_AB().Show();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            classInsertBOMData.NucorBOM_AB();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            classInsertBOMData.Blank_AB();
        }



        private void tBoxWoodReq_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            tBoxWoodReq.Text = "Please wait; this could take several minutes";

            tBoxWoodReq.Text = folderOperations.WoodRequirements();



        }

        private void button3_Click(object sender, EventArgs e)
        {
            tBoxWoodReq.Text = "Please wait; this could take several minutes";
            try
            {
                folderOperations.createTCWidthDocument();
            }
            catch
            {
                tBoxWoodReq.Text = "Process Failed";
            }
            this.Close();
        }

        private void btnQuickTCWidth_Click(object sender, EventArgs e)
        {
            ExtractJoistDetails extractJoistDetails = new ExtractJoistDetails();
            Job job = new Job();
            job = extractJoistDetails.JobFromShoporderJoistDetails();

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            //Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.

            string excelPath = System.IO.Path.GetTempFileName();

            System.IO.File.WriteAllBytes(excelPath, Properties.Resources.DesignTCWidths);

            oWB = oXL.Workbooks.Open(excelPath);

            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            int joistcount = 0;
            foreach (Joist joist in job.Joists)
            {
                joistcount++;
                string excelRow = Convert.ToString(joistcount + 6);
                oSheet.get_Range("A" + excelRow, Missing.Value).Value = joist.Mark;
                oSheet.get_Range("B" + excelRow, Missing.Value).Value = joist.Quantity;
                oSheet.get_Range("C" + excelRow, Missing.Value).Value = joist.Description;
                oSheet.get_Range("D" + excelRow, Missing.Value).Value = stringManipulation.DecimilLengthToHyphen(joist.BaseLength);
                oSheet.get_Range("E" + excelRow, Missing.Value).Value = joist.TCWidth;
            }

        }

        private void btnWoodReqFromJoistDetails_Click(object sender, EventArgs e)
        {
            ExtractJoistDetails extractJoistDetails = new ExtractJoistDetails();
            Job job = new Job();
            job = extractJoistDetails.JobFromShoporderJoistDetails();

            double dblFiveInch = 0.0;
            double dblSevenInch = 0.0;
            double dblEightInch = 0.0;
            double dblNineInch = 0.0;
            double dblElevenInch = 0.0;
            double dblThirteenInch = 0.0;
            
            foreach(Joist joist in job.Joists)
            {
                double qty = Convert.ToDouble(joist.Quantity);
                if (joist.TCWidth == "5") { dblFiveInch = dblFiveInch + qty*(joist.BaseLength + joist.TCXL + joist.TCXR); }
                if (joist.TCWidth == "7") { dblSevenInch = dblSevenInch + qty * (joist.BaseLength + joist.TCXL + joist.TCXR); }
                if (joist.TCWidth == "8") { dblEightInch = dblEightInch + qty * (joist.BaseLength + joist.TCXL + joist.TCXR); }
                if (joist.TCWidth == "9") { dblNineInch = dblNineInch + qty * (joist.BaseLength + joist.TCXL + joist.TCXR); }
                if (joist.TCWidth == "11") { dblElevenInch = dblElevenInch + qty * (joist.BaseLength + joist.TCXL + joist.TCXR); }
                if (joist.TCWidth == "13") { dblThirteenInch = dblThirteenInch + qty * (joist.BaseLength + joist.TCXL + joist.TCXR); }

            }

            string stringFiveInch, stringSevenInch, stringEightInch, stringNineInch, stringElevenInch, stringThirteenInch;
            stringFiveInch = stringSevenInch = stringEightInch = stringNineInch = stringElevenInch = stringThirteenInch = String.Empty;
            if (dblFiveInch != 0)
            {
                stringFiveInch = "5\" = " + Convert.ToString(Convert.ToInt32(dblFiveInch)) + "  lf \r\n";
            }
            if (dblSevenInch != 0)
            {
                stringSevenInch = "7\" = " + Convert.ToString(Convert.ToInt32(dblSevenInch)) + "  lf \r\n";
            }
            if (dblEightInch != 0)
            {
                stringEightInch = "8\" = " + Convert.ToString(Convert.ToInt32(dblEightInch)) + "  lf \r\n";
            }
            if (dblNineInch != 0)
            {
                stringNineInch = "9\" = " + Convert.ToString(Convert.ToInt32(dblNineInch)) + "  lf \r\n";
            }
            if (dblElevenInch != 0)
            {
                stringElevenInch = "11\" = " + Convert.ToString(Convert.ToInt32(dblElevenInch)) + "  lf \r\n";
            }
            if (dblThirteenInch != 0)
            {
                stringThirteenInch = "13\" = " + Convert.ToString(Convert.ToInt32(dblThirteenInch)) + "  lf \r\n";
            }
            /*
            string woodRequirements =

                "5\" = " + Convert.ToString(Convert.ToInt16(fiveInch)) + "  lf \r\n" +
                "7 1/8\" = " + Convert.ToString(Convert.ToInt16(sevenInch)) + "  lf \r\n" +
                "8 1/8\" = " + Convert.ToString(Convert.ToInt16(eightInch)) + "  lf \r\n" +
                "9 1/8\" = " + Convert.ToString(Convert.ToInt16(nineInch)) + "  lf \r\n" +
                "10 1/8\" = " + Convert.ToString(Convert.ToInt16(tenInch)) + "  lf \r\n" +
                "11 1/8\" = " + Convert.ToString(Convert.ToInt16(elevenInch)) + "  lf \r\n";
            */
            string woodRequirements =

                stringFiveInch + stringSevenInch + stringEightInch + stringNineInch + stringElevenInch + stringThirteenInch;

            tBoxWoodReq.Text = woodRequirements;



        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            string reportPath = "";
            System.Windows.Forms.OpenFileDialog openFile = new System.Windows.Forms.OpenFileDialog();
            openFile.Filter = "Excel files|*.xlsm";
            openFile.Title = "Select BOM";
            if (openFile.ShowDialog() == (System.Windows.Forms.DialogResult.OK))
            {
                reportPath = openFile.FileName;
            }

            if (reportPath != "")
            {
                DESign_Bot_FS_Tools.Seperator.getAllBomInfo(reportPath);
            }
            else
            {
                MessageBox.Show("No BOM Selected");
            }
        }
    }
}
