using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DESign_WordAddIn
{
    public partial class FormSOtoRISA : Form
    {
        public FormSOtoRISA()
        {
            InitializeComponent();
        }
        SOtoRISA sOtoRISA = new SOtoRISA();

        JoistCoverSheet JoistCoverSheet = new JoistCoverSheet();

        StringManipulation StringManipulation = new StringManipulation();

        QueryAngleDataXML QueryAngleDataXML = new QueryAngleDataXML();
        List<CheckBox> cbChosenJoist = new List<CheckBox>();



        List<List<string>> joistData;


        private void FormSOtoRISA_Load(object sender, EventArgs e)
        {


            string clipboard = Clipboard.GetText();

            joistData = JoistCoverSheet.JoistData();

            var labelMarkTitle = new Label();
            var labelChosenJoist = new Label();




            labelMarkTitle.Size = new System.Drawing.Size(60, 15);



            labelMarkTitle.AutoSize = true;
            labelChosenJoist.AutoSize = true;


            labelMarkTitle.Location = new Point(20, 60);
            labelChosenJoist.Location = new Point(85, 60);


            labelMarkTitle.Text = "MARK";
            labelChosenJoist.Text = "SELECT JOIST";


            labelMarkTitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelChosenJoist.Font = new Font("Times New Roman", 9, FontStyle.Bold);

            labelMarkTitle.TextAlign = ContentAlignment.MiddleLeft;
            labelChosenJoist.TextAlign = ContentAlignment.MiddleCenter;



            this.Controls.Add(labelMarkTitle);
            this.Controls.Add(labelChosenJoist);


            List<string> joistMarks = joistData[0];

            int joistDataLength = joistMarks.Count();



            var labelMark = new Label[joistDataLength];
            var cbChooseJoist = new CheckBox[joistDataLength];


            for (var i = 0; i < joistDataLength; i++)
            {
                var labelMarks = new Label();
                var cbChooseThisJoist = new CheckBox();


                int Y = 150 + (i * 25);

                labelMarks.Text = joistMarks[i];
                labelMarks.Location = new Point(20, Y);
                labelMarks.Size = new System.Drawing.Size(50, 25);





                cbChooseThisJoist.Location = new Point(110, Y);
                cbChooseThisJoist.Size = new System.Drawing.Size(20, 20);




                this.Controls.Add(cbChooseThisJoist);
                this.Controls.Add(labelMarks);


                cbChosenJoist.Add(cbChooseThisJoist);



                cbChooseJoist[i] = cbChooseThisJoist;
                labelMark[i] = labelMarks;

            }


        }

        public void showChosenJoists ()
        {
            for (int i=0; i<joistData[0].Count; i++)
            {
                if (cbChosenJoist[i].Checked == true)
                {
                    sOtoRISA.getNodes(joistData[0][i]);
                }
            }
        }

        private void btnExportSOtoRISA_Click(object sender, EventArgs e)
        {
            showChosenJoists();
        }
    }
}
