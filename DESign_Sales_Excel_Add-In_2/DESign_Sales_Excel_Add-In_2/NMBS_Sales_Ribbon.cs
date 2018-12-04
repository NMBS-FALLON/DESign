using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DESign_Sales_Excel_Add_In.Properties;
using DESign_Sales_Excel_Add_In_2.BlueBeam;
using DESign_Sales_Excel_Add_In_2.Worksheet_Values;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace DESign_Sales_Excel_Add_In_2
{
  public partial class NMBS_Sales_Ribbon
  {
    private CommandBarButton button;
    private CommandBar cb;

    private void NMBS_Sales_Ribbon_Load(object sender, RibbonUIEventArgs e)
    {
      var app = Globals.ThisAddIn.Application;
      cb = app.CommandBars["Cell"];
      button =
        cb.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true) as
          CommandBarButton;
      button.Tag = "Sprinkler Form";
      button.Caption = "Add Sprinkler Load";
      button.Style = MsoButtonStyle.msoButtonCaption;
      button.Click += addSprinklerButton_Click;
      app.SheetBeforeRightClick += app_SheetBeforeRightClick;
    }

    private void app_SheetBeforeRightClick(object Sh, Range Target, ref bool Cancel)
    {
      var app = Globals.ThisAddIn.Application;
      var ws = app.ActiveSheet as Worksheet;
      if (ws.Name == "Marks" || ws.Name == "Base Types")
      {
        var rowsNotAllowed = new List<int> {1, 2, 3, 4};
        if (Target.Column == 17 && rowsNotAllowed.Contains(Target.Row) == false)
        {
          button.Visible = true;
          return;
        }
      }

      button.Visible = false;
    }


    private void addSprinklerButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
    {
      new Tools.FormSprinklerLoading().Show();
    }

    private void button1_Click(object sender, RibbonControlEventArgs e)
    {
      var convert_Takeoff_Form = new Convert_Takeoff_Form();
      convert_Takeoff_Form.Show();
    }

    private void btnNewTakeoff_Click(object sender, RibbonControlEventArgs e)
    {
      var excelPath = Path.GetTempFileName();
      File.WriteAllBytes(excelPath, Resources.TAKEOFF_CONCEPT);

      var oXLTemp = Globals.ThisAddIn.Application;
      var workbooks = oXLTemp.Workbooks;
      Workbook workbook;
      workbook = workbooks.Open(excelPath);

      var saveFileDialog = new SaveFileDialog();
      saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
      saveFileDialog.ShowDialog();
      if (saveFileDialog.FileName != "")
      {
        workbook.CheckCompatibility = false;
        workbook.SaveAs(saveFileDialog.FileName);
      }

      Marshal.ReleaseComObject(workbook);
      Marshal.ReleaseComObject(workbooks);
      Marshal.ReleaseComObject(oXLTemp);
      GC.Collect();

      /*            var mainWindow = (System.Windows.Window)System.Windows.Application.LoadComponent(
                                     new System.Uri("/Design_Sales_Excel_Add-in;component/NewTakeoff.xaml", UriKind.Relative));
                  var application = new System.Windows.Application();
                  application.Run(mainWindow);
      */
    }

    private void btnInfo_Click(object sender, RibbonControlEventArgs e)
    {
      MessageBox.Show("SPECIAL BASETYPES:\r\n" +
                      "   - [ALL]   : BRINGS INFO ONTO ALL MARKS\r\n" +
                      "   - [ALL J] : BRINGS INFO ONTO ALL JOISTS\r\n" +
                      "   - [ALL G] : BRINGS INFO ONTO ALL GIRDERS\r\n" +
                      "   - [ALL: {SEQ.}] : ALL MARKS IN SPECIFIED SEQUENCE\r\n" +
                      "   - [ALL J: {SEQ.}] : ALL JOISTS IN SPECIFIED SEQUENCE\r\n" +
                      "   - [ALL G: {SEQ.}] : ALL GIRDERS IN SPECIFIED SEQUENCE\r\n" +
                      "   - NOTE: DO NOT USE [ALL J] OR [ALL G] FOR DESIGNATIONS\r\n" +
                      "RULES FOR SEPERATING SEISMIC:\r\n" +
                      "   - MAKE SURE THERE ARE NO LOADS IN LC3\r\n" +
                      "   - ADD LOADS AND BEND CHECKS SHOULD BE ADDED\r\n" +
                      "     USING THE SPECIAL BASETYPES RATHER THAN VIA\r\n" +
                      "     COVER SHEET NOTES\r\n" +
                      "   - SEPERATION IS ONLY ALLOWED ON ROOFS.\r\n" +
                      "   - IF THE JOIST DESIGNATION LL IS FROM\r\n" +
                      "     SNOW, THE FLAT ROOF SNOW LOAD (Pf)\r\n" +
                      "     MUST BE LESS THAN 30 PSF.\r\n" +
                      "   - FOR SEPERATION TO OCCUR ON GIRDERS,\r\n" +
                      "     THE DESIGNATION MUST BE IN TL/LL FORM\r\n" +
                      "     (i.e. 54G7N12.5/5.8K).\r\n" +
                      "OTHER NOTES:\r\n" +
                      "   - SEQUENCES NEED TO BE BETWEEN THE { & } CHARACTERS\r\n" +
                      "   - SINGLE PITCH GEOMETRY: <20-30>LH10\r\n" +
                      "   - DOUBLE PITCH GEOMETRY: <20-30-20>LH10\r\n" +
                      "   - JOIST DESCRIPTION SHORTHAND NOTATION:\r\n" +
                      "     '+' CHARACTER GETS REPLACED W/ 'K'\r\n" +
                      "     '-' CHARACTER GETS REPLACED W/ 'LH'\r\n" +
                      "     '+-' CHARACTER GETS REPLACED W/ 'KCS'\r\n" +
                      "     FIRST OCCURANCE OF '*' GETS REPLACED W/ 'G'\r\n" +
                      "     SECOND OCCURANCE OF '*' GETS REPLACED W/ 'N'\r\n" +
                      "     THIRD OCCURANCE OF '*' GETS REPLACED W/ 'K'\r\n" +
                      "     EXAMPLES:\r\n" +
                      "        20+5 = 20K5, 32-6 = 32LH6, 48*6*10*2 = 48G6N10K2");
    }


    private void btnJobCheck_Click(object sender, RibbonControlEventArgs e)
    {
      var dialogResult = MessageBox.Show("INCLUDE BLUEBEAM-TAKEOFF CHECK?", "OPTIONS", MessageBoxButtons.YesNo);
      if (dialogResult == DialogResult.Yes)
      {
        var errors = "";

        var thisTakeoff = new Takeoff();
        thisTakeoff = thisTakeoff.ImportTakeoff();
        var blueBeamTakeoff = new Takeoff();
        var extractor = new ExtractBlueBeamMarkups();
        blueBeamTakeoff = extractor.TakeoffFromBB();

        var thistakeoffJoists =
          from s in thisTakeoff.Sequences
          from j in s.Joists
          select j;


        var blueBeamJoists =
          from s in blueBeamTakeoff.Sequences
          from j in s.Joists
          select j;


        var rg = new Regex(@"\d+");

        var bbJoistTups =
          blueBeamJoists
            .GroupBy(x => x.Mark.Text)
            .Select(g => new Tuple<string, int?>(rg.Match(g.Key).Value, g.Sum(x => x.Quantity.Value)));

        foreach (var thisTakeoffJoist in thistakeoffJoists)
        {
          var blueBeamMatchedJoists = bbJoistTups.Where(joist => joist.Item1 == thisTakeoffJoist.Mark.Text);
          if (blueBeamMatchedJoists.Any())
          {
            var blueBeamJoist = blueBeamMatchedJoists.First();
            var blueBeamQty = blueBeamJoist.Item2;
            var thisTakeoffQty = thisTakeoffJoist.Quantity.Value;

            if (blueBeamQty != thisTakeoffQty)
              errors = errors + string.Format(
                         "Mark {0}:  Takeoff Qty = {1}, BlueBeam Qty = {2}.\r\n\r\n",
                         thisTakeoffJoist.Mark.Text, thisTakeoffQty, blueBeamQty);
          }
          else
          {
            errors = errors + string.Format("Takeoff Mark {0} is not in the BlueBeam markups.\r\n\r\n",
                       thisTakeoffJoist.Mark.Text);
          }
        }

        foreach (var bbJoist in bbJoistTups)
        {
          var takeoffMatchedJoists = thistakeoffJoists.Where(toJoist => toJoist.Mark.Text == bbJoist.Item1);
          if (takeoffMatchedJoists.Any() == false)
            errors = errors + string.Format("BlueBeam Mark {0} is not on the takeoff.\r\n\r\n", bbJoist.Item1);
        }


        if (errors == "")
        {
          MessageBox.Show("TAKEOFF MATCHES BLUEBEAM!!!");
        }
        else
        {
          var filePath = Path.GetTempPath() + "Errors.txt";
          File.WriteAllText(filePath, "MISMATCHES:\r\n\r\n\r\n" + errors);
          Process.Start(filePath);

          var dialogResult2 = MessageBox.Show("Would you like to transpose BlueBeam quantities onto Takeoff?",
            "OPTIONS", MessageBoxButtons.YesNo);
          if (dialogResult == DialogResult.Yes) thisTakeoff.AddQuantitiesFromBB(blueBeamTakeoff);
        }
      }

      else if (dialogResult == DialogResult.No)
      {
        var thisTakeoff = new Takeoff();
        thisTakeoff = thisTakeoff.ImportTakeoff();
        MessageBox.Show("DONE.\r\n IF NO ERROR REPORT POPPED UP, YOU ARE ALL GOOD.");
      }
    }
  }
}