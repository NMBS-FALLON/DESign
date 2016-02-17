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

namespace DESign_BOT
{
    
    public partial class FormNMBSHelper : Form
    {
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
            //joistDetails.JoistsFromJoistDetails();
        }
    }
}
