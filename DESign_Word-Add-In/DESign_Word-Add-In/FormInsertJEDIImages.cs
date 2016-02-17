using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing.Imaging;
using System.Reflection;

namespace DESign_WordAddIn
{
    public partial class FormInsertJEDIImages : Form
    {
        MSwordOperations msWordOperations = new MSwordOperations();
        ImageOperations imageOperations = new ImageOperations();
        public FormInsertJEDIImages()
        {
            InitializeComponent();
        }
        string jobNumber = "";
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void FormInsertJEDIImages_Load(object sender, EventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            jobNumber = doc.Name.Split('_')[0];
            string docName = doc.Name.Split('.')[0];

            // DirectoryInfo d = new DirectoryInfo(@"\\nmbsfaln-fs\Engr\_MANUALSHOPORDERS\" + jobNumber + "\\");
            DirectoryInfo d = new DirectoryInfo(@"C:\Users\darien.shannon\Documents\_MANUALSHOPORDERS\" + jobNumber + "\\");
            FileInfo[] infos = d.GetFiles();
            int row = 0;
            foreach (FileInfo f in infos)
            {
                if (f.Name != docName && f.Name!=doc.Name && f.Name.Contains(docName)==true)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[row].Cells[1].Value = f.Name;
                    row++;                 
                }

            }

             
        }

        private void btnInsertSelectedFiles_Click(object sender, EventArgs e)
        {

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            int objectsInWord = Globals.ThisAddIn.Application.ActiveDocument.InlineShapes.Count;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[0].Value) == true)
                {
                    
                    string emfPath = @"C:\Users\darien.shannon\Documents\_MANUALSHOPORDERS\" + jobNumber +"\\" + Convert.ToString(row.Cells[1].Value);
                    using (var src = new Metafile(emfPath))
                    using (var bmp = new Bitmap(@"C:\Users\darien.shannon\Documents\_MANUALSHOPORDERS\blank.tif"))
                    using (var gr = Graphics.FromImage(bmp))
                    {

                        gr.DrawImage(src, new Rectangle(0, 0, bmp.Width, bmp.Height));
                        bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        imageOperations.ConvertToBlackAndWhite(bmp);
                        imageOperations.createPDF(bmp, @"C:\Users\darien.shannon\Documents\_MANUALSHOPORDERS\newPDF_" + jobNumber + ".pdf");
                        bmp.Save(@"C:\Users\darien.shannon\Documents\_MANUALSHOPORDERS\" + jobNumber + "\\" + Convert.ToString(row.Cells[1].Value).Split('.')[0] + ".tif", ImageFormat.Bmp);
                        msWordOperations.addSection(Word.WdOrientation.wdOrientLandscape);
                        selection.InlineShapes.AddOLEObject("Paint.Picture", @"C:\Users\darien.shannon\Documents\_MANUALSHOPORDERS\" + jobNumber + "\\" + Convert.ToString(row.Cells[1].Value).Split('.')[0] + ".tif", true);
                        objectsInWord++;

                        Globals.ThisAddIn.Application.ActiveDocument.InlineShapes[objectsInWord].Height = 450.0f;
                        Globals.ThisAddIn.Application.ActiveDocument.InlineShapes[objectsInWord].Width = 684.0f;

                        gr.Dispose();
                        bmp.Dispose();
                        src.Dispose();
                    }
                    System.GC.Collect();

                }
                    
            }
            this.Close();
        }
    }
}
