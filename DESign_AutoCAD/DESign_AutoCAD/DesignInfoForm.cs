using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Text.RegularExpressions;
using DESign_BASE;
using System.Linq.Expressions;

namespace DESign_AutoCAD
{
    public partial class DesignInfoForm : Form
    {

        public (bool AddJoistTcwCrimped, bool AddJoistTcwNonCrimped, bool AddBoltLength, bool AddGirderTcw, bool AddWeights, bool AddTcMaxBridging, bool AddBcMaxBridging, bool useSql, string Plant, string JobNumber) Return { get; set; }
        public DesignInfoForm()
        {
            InitializeComponent();
        }

        private void DesignInfoForm_Load(object sender, EventArgs e)
        {
            cbDataSource.SelectedIndex = 0;
            tbJobNumber.Text = GetJobNumber();
        }
        /*
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
           
            if (keyData == (Keys.Control | Keys.U))
            {
                clbInfoSelect.Items.Add("WEIGHT");
            }
            return base.ProcessCmdKey(ref msg, keyData);
          
        }
        */
 
        private void btnAddInfo_Click(object sender, EventArgs e)
        {



            var addJoistTcwCrimped = clbInfoSelect.CheckedIndices.Contains(0);
            var addJoistTcwNonCrimped = clbInfoSelect.CheckedIndices.Contains(1);
            var addGirderTcw = clbInfoSelect.CheckedIndices.Contains(2);
            var addJoistBoltLength = clbInfoSelect.CheckedIndices.Contains(3);
            var addWeights = clbInfoSelect.CheckedIndices.Contains(4);
            var addTcMaxBridging = clbInfoSelect.CheckedIndices.Contains(5);
            var addBcMaxBridging = clbInfoSelect.CheckedIndices.Contains(6);
            var useSql = cbDataSource.Text == "SQL" ? true : false;
            var plant = cbPLant.Text;
            var jobNumber = tbJobNumber.Text;

            var allInputsReceived = true;

            if (cbPLant.Text == "")
            {
                allInputsReceived = false;
                MessageBox.Show("PLEASE SELECT A PLANT");
            }
            
            if (useSql && tbJobNumber.Text == "")
            {
                allInputsReceived = false;
                MessageBox.Show("PLEASE INPUT A JOB NUMBER");
            }



            this.Return = (addJoistTcwCrimped, addJoistTcwNonCrimped, addJoistBoltLength, addGirderTcw, addWeights, addTcMaxBridging, addBcMaxBridging, useSql, plant, jobNumber);
            if (allInputsReceived)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                this.DialogResult = DialogResult.None;
            }
        }

        public void removeDesignInfo()
        {

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                    (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead, false, true);

                foreach (ObjectId id in btr)
                {
                    Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false, true) as Entity;
                    if (currentEntity == null)
                    {
                        continue;
                    }
                    if (currentEntity.GetType() == typeof(MText))
                    {
                        ((MText)currentEntity).Contents = removeTCWidths(((MText)currentEntity).Contents);
                    }
                    if (currentEntity.GetType() == typeof(DBText))
                    {
                        ((DBText)currentEntity).TextString = removeTCWidths(((DBText)currentEntity).TextString);
                    }
                    if (currentEntity.GetType() == typeof(RotatedDimension))
                    {
                        ((RotatedDimension)currentEntity).DimensionText = removeTCWidths(((RotatedDimension)currentEntity).DimensionText);
                    }
                }
                tr.Commit();
            }

        }

        private string removeTCWidths(string text)
        {
            if (text.Contains("|"))
            {
                text = text.Substring(0, text.IndexOf('|'));
            }

            return text;

        }

        private void clbInfoSelect_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        public static string GetJobNumber()
        {
            var jobNum = "";
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId id in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);

                    var brIds = btr.GetBlockReferenceIds(true, true);

                    foreach (ObjectId brId in brIds)
                    {
                        var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);

                        foreach (ObjectId oId in br.AttributeCollection)
                        {
                            var attRef = (AttributeReference)tr.GetObject(oId, OpenMode.ForRead);

                            if (attRef.Tag.Equals("FILENO", StringComparison.CurrentCultureIgnoreCase))
                            {
                                jobNum = attRef.TextString;
                                break;
                            }
                        }


                    }

                }




            }

            return jobNum;
        }
    }


}
