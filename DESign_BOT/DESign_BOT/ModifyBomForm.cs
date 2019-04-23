using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DESign.BomTools;

namespace DESign_BOT
{
    public partial class ModifyBomForm : Form
    {
        public ModifyBomForm()
        {
            InitializeComponent();
        }

        public DESign.BomTools.SeismicSeperator.SeperatorInfo SeperatorInfo { get; set; }


        private void ModifyBomForm_Load(object sender, EventArgs e)
        {
            clbBomModifications.SetItemChecked(0, true);
            clbBomModifications.SetItemChecked(1, true);
        }

        private void BtnModifyBom_Click(object sender, EventArgs e)
        {
            var seperateSeismic = clbBomModifications.CheckedIndices.Contains(0);
            var checkGirderIp = clbBomModifications.CheckedIndices.Contains(1);

            this.SeperatorInfo = new DESign.BomTools.SeismicSeperator.SeperatorInfo(seperateSeismic, checkGirderIp);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
