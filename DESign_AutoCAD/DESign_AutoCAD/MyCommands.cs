using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Text.RegularExpressions;
using DESign_BASE;
using System.Linq.Expressions;
using System.Linq;

[assembly: CommandClass(typeof(DESign_AutoCAD.MyCommands))]

namespace DESign_AutoCAD
{

    public class MyCommands
    {


        [CommandMethod("MyCommandGroup", "DESIGN", CommandFlags.Modal)]
        public void Design()
        {

            removeDesignInfo();


            bool addJoistTcw = false;
            bool addBoltLength = false;
            bool addGirderTcw = false;
            bool addWeights = false;

            using (var dif = new DesignInfoForm())
            {
                var result = dif.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    (addJoistTcw, addBoltLength, addGirderTcw, addWeights) = dif.Return;
                }
            }

            var weightFactor = 0.02;

            if (addWeights)
            {
                using (var addWeightPercentForm = new WeightFactorForm())
                {
                    var result = addWeightPercentForm.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        weightFactor = addWeightPercentForm.WeightPercentToAdd;
                    }
                }
            }



            Job job = new Job();

            if (addJoistTcw || addBoltLength || addGirderTcw || addWeights)
            {
                job = ExtractJoistDetails.JobFromShoporderJoistDetails();
            }

            if (job.Joists == null && job.Girders == null)
            {
                return;
            }

            var joistInfoList = new List<(string Mark, int quantity, string TcWidth, int BoltSize, double Weight)>();

            foreach (var joist in job.Joists)
            {
                var mark = joist.Mark;
                var tcWidth = joist.TCWidth;
                var bcSize = joist.BC;
                var bcVleg = QueryAngleData.DblVleg(bcSize);
                var isMerchantBc = !bcSize.Contains("A");
                var boltSize = isMerchantBc ?
                                 System.Math.Max((short)3, (short)System.Math.Ceiling(bcVleg + 1)) :
                                 System.Math.Max((short)3, (short)System.Math.Ceiling(bcVleg + (1 - 0.078)));
                joistInfoList.Add((mark, joist.Quantity, tcWidth, boltSize, joist.WeightInLBS));
            }

            var joistTcWidthMajority =
                joistInfoList
                .GroupBy(info => info.TcWidth)
                .Select(group => (TcWidth: group.Key, Sum: group.Sum(info => info.quantity)))
                .OrderByDescending(info => info.Sum)
                .First()
                .TcWidth;

            var boltLengthMajority =
                joistInfoList
                .GroupBy(info => info.BoltSize)
                .Select(group => (BoltSize: group.Key, Sum: group.Sum(info => info.quantity)))
                .OrderByDescending(info => info.Sum)
                .First()
                .BoltSize;

            var marksWithMessages = new List<(string Mark, List<string> Messages)>();


            foreach (var (Mark, Quantity, TcWidth, BoltLength, Weight) in joistInfoList)
                {
                    var messages = new List<string>();
                    if (addJoistTcw && TcWidth != joistTcWidthMajority) { messages.Add("TCW=" + TcWidth); }
                    if (addBoltLength && BoltLength != boltLengthMajority) { messages.Add("BL=" + BoltLength); }
                    if (addWeights) { messages.Add("WT=" + ((int)(System.Math.Ceiling(Weight * (1 + weightFactor) / 10.0) * 10.0)).ToString()); }
                    if (messages.Count != 0)
                    {
                        marksWithMessages.Add((Mark, messages));
                    }
                }

            foreach (var girder in job.Girders)
            {
                var messages = new List<string>();
                if (addGirderTcw) { messages.Add("TCW=" + girder.TCWidth); }
                if (addWeights) { messages.Add("WT=" + ((int)(System.Math.Ceiling(girder.WeightInLBS * (1 + weightFactor) / 10.0) * 10.0)).ToString()); }
                if (messages.Count != 0)
                {
                    marksWithMessages.Add((girder.Mark, messages));
                }
            }


            foreach (var (Mark, Messages) in marksWithMessages)
            {
                var markWithMessage = string.Format("{0} [{1}]", Mark, string.Join(",", Messages));

                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (doc.LockDocument())
                {
                    Database db = doc.Database;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                            (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                        foreach (ObjectId id in btr)
                        {
                            Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                            if (currentEntity == null)
                            {
                                continue;
                            }
                            if (currentEntity.GetType() == typeof(MText))
                            {
                                string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                                if (strippedText == Mark)
                                {
                                    MText mText = (MText)currentEntity;
                                    string mTextContent = mText.Contents;
                                    string newmTextContent = Regex.Replace(mTextContent, Mark, markWithMessage);
                                    ((MText)currentEntity).Contents = newmTextContent;
                                }
                            }
                            if (currentEntity.GetType() == typeof(DBText))
                            {
                                string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                                if (strippedText == Mark)
                                {
                                    DBText dbText = (DBText)currentEntity;
                                    string dbTextString = dbText.TextString;
                                    string newdbTextString = Regex.Replace(dbTextString, Mark, markWithMessage);
                                    ((DBText)currentEntity).TextString = newdbTextString;
                                }
                            }
                            if (currentEntity.GetType() == typeof(RotatedDimension))
                            {
                                string dimText = ((RotatedDimension)currentEntity).DimensionText;

                                if (dimText.Contains(string.Format("-{0}\\X", Mark)) == true ||
                                    dimText.Contains(string.Format("-{0} ", Mark)))
                                {
                                    string replace = string.Format("-{0}", Mark);
                                    string replacement = string.Format("-{0}", markWithMessage);
                                    string newrdText = Regex.Replace(dimText, replace, replacement);
                                    ((RotatedDimension)currentEntity).DimensionText = newrdText;
                                }

                                string[] dimTextArray = Regex.Split(dimText, "-");
                                string mark = dimTextArray[dimTextArray.Length - 1];

                                if (mark == Mark)
                                {
                                    string replace = string.Format("-{0}", Mark);
                                    string replacement = string.Format("-{0}", markWithMessage);
                                    string newrdText = Regex.Replace(dimText, replace, replacement);
                                    ((RotatedDimension)currentEntity).DimensionText = newrdText;
                                }

                            }

                        }
                        tr.Commit();
                    }
                }

            }

            if (addJoistTcw || addBoltLength)
            {
                var majorityMessage = "";
                if (addJoistTcw)
                {
                    majorityMessage += joistTcWidthMajority + "\" majority TC width.\n";
                }
                if (addBoltLength)
                {
                    majorityMessage += boltLengthMajority + "\" majoirty bolt length.\n";
                }
                System.Windows.Forms.MessageBox.Show(majorityMessage);
            }
        }
    
        
/*
        [CommandMethod("MyCommandGroup", "DESign_GIRDER_TCW", CommandFlags.Modal)]
        public void tcWidths_Girders()
        {
            var job = ExtractJoistDetails.JobFromShoporderJoistDetails();
            if (job.Joists == null && job.Girders == null)
            {
                return;
            }
            List<Girder> girders = job.Girders;


            foreach (Girder girder in girders)
            {
                string markWithWidth = girder.Mark + " [TCW=" + girder.TCWidth + "]";
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                        (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                    foreach (ObjectId id in btr)
                    {
                        Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                        if (currentEntity == null)
                        {
                            continue;
                        }
                        if (currentEntity.GetType() == typeof(MText))
                        {
                            string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                            if (strippedText == girder.Mark)
                            {
                                MText mText = (MText)currentEntity;
                                string mTextContent = mText.Contents;
                                string newmTextContent = Regex.Replace(mTextContent, girder.Mark, markWithWidth);
                                ((MText)currentEntity).Contents = newmTextContent;
                            }
                        }
                        if (currentEntity.GetType() == typeof(DBText))
                        {
                            string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                            if (strippedText == girder.Mark)
                            {
                                DBText dbText = (DBText)currentEntity;
                                string dbTextString = dbText.TextString;
                                string newdbTextString = Regex.Replace(dbTextString, girder.Mark, markWithWidth);
                                ((DBText)currentEntity).TextString = newdbTextString;
                            }
                        }
                        if (currentEntity.GetType() == typeof(RotatedDimension))
                        {
                            string dimText = ((RotatedDimension)currentEntity).DimensionText;

                            if (dimText.Contains(string.Format("-{0}\\X", girder.Mark)) == true ||
                                dimText.Contains(string.Format("-{0} ", girder.Mark)))
                            {
                                string replace = string.Format("-{0}", girder.Mark);
                                string replacement = string.Format("-{0}", markWithWidth);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                            string[] dimTextArray = Regex.Split(dimText, "-");
                            string mark = dimTextArray[dimTextArray.Length - 1];

                            if (mark == girder.Mark)
                            {
                                string replace = string.Format("-{0}", girder.Mark);
                                string replacement = string.Format("-{0}", markWithWidth);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                        }

                    }
                    tr.Commit();
                }
            }
        }


        [CommandMethod("MyCommandGroup", "DESign_JOIST_TCW_AND_BL", CommandFlags.Modal)]
        public void tcWidths_joists_with_bolts()
        {
            var job = ExtractJoistDetails.JobFromShoporderJoistDetails();

            if (job.Joists == null && job.Girders == null)
            {
                return;
            }

            var markInfoList = new List<(string Mark, int quantity, string TcWidth, int BoltSize)>();

            foreach (var joist in job.Joists)
            {
                var mark = joist.Mark;
                var tcWidth = joist.TCWidth;
                var bcSize = joist.BC;
                var bcVleg = QueryAngleData.DblVleg(bcSize);
                var isMerchantBc = !bcSize.Contains("A");
                var boltSize = isMerchantBc ?
                                 System.Math.Max((short)3, (short)System.Math.Ceiling(bcVleg + 1)) :
                                 System.Math.Max((short)3, (short)System.Math.Ceiling(bcVleg + (1 - 0.078)));
                markInfoList.Add((mark, joist.Quantity, tcWidth, boltSize));
            }

            var tcWidthMajority =
                markInfoList
                .GroupBy(info => info.TcWidth)
                .Select(group => (TcWidth: group.Key, Sum: group.Sum(info => info.quantity)))
                .OrderByDescending(info => info.Sum)
                .First()
                .TcWidth;

            var boltMajority =
                markInfoList
                .GroupBy(info => info.BoltSize)
                .Select(group => (BoltSize: group.Key, Sum: group.Sum(info => info.quantity)))
                .OrderByDescending(info => info.Sum)
                .First()
                .BoltSize;


            var marksWithMessages = new List<(string Mark, List<string> Messages)>();

            foreach (var (Mark, Quantity, TcWidth, BoltSize) in markInfoList)
            {
                var messages = new List<string>();
                if (TcWidth != tcWidthMajority) { messages.Add("TCW=" + TcWidth); }
                if (BoltSize != boltMajority) { messages.Add("BL=" + BoltSize); }
                if (messages.Count != 0)
                {
                    marksWithMessages.Add((Mark, messages));
                }
            }

            foreach (var (Mark, Messages) in marksWithMessages)
            {
                var markWithMessage = string.Format("{0} [{1}]", Mark, string.Join(",", Messages));

                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                        (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                    foreach (ObjectId id in btr)
                    {
                        Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                        if (currentEntity == null)
                        {
                            continue;
                        }
                        if (currentEntity.GetType() == typeof(MText))
                        {
                            string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                            if (strippedText == Mark)
                            {
                                MText mText = (MText)currentEntity;
                                string mTextContent = mText.Contents;
                                string newmTextContent = Regex.Replace(mTextContent, Mark, markWithMessage);
                                ((MText)currentEntity).Contents = newmTextContent;
                            }
                        }
                        if (currentEntity.GetType() == typeof(DBText))
                        {
                            string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                            if (strippedText == Mark)
                            {
                                DBText dbText = (DBText)currentEntity;
                                string dbTextString = dbText.TextString;
                                string newdbTextString = Regex.Replace(dbTextString, Mark, markWithMessage);
                                ((DBText)currentEntity).TextString = newdbTextString;
                            }
                        }
                        if (currentEntity.GetType() == typeof(RotatedDimension))
                        {
                            string dimText = ((RotatedDimension)currentEntity).DimensionText;

                            if (dimText.Contains(string.Format("-{0}\\X", Mark)) == true ||
                                dimText.Contains(string.Format("-{0} ", Mark)))
                            {
                                string replace = string.Format("-{0}", Mark);
                                string replacement = string.Format("-{0}", markWithMessage);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                            string[] dimTextArray = Regex.Split(dimText, "-");
                            string mark = dimTextArray[dimTextArray.Length - 1];

                            if (mark == Mark)
                            {
                                string replace = string.Format("-{0}", Mark);
                                string replacement = string.Format("-{0}", markWithMessage);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                        }

                    }
                    tr.Commit();
                }

            }

            System.Windows.Forms.MessageBox.Show(
                tcWidthMajority + "\" majority TC width.\n" +
                boltMajority + "\" majority bolt size.");
        }

        /*
        
        [CommandMethod("MyCommandGroup", "TCWIDTHS_JOISTS_EXCLUDE_5_IN", CommandFlags.Modal)]
        public void tcWidths_joists_no_5_in()   
        {
            Job job1 = new Job();
            job1 = joistDetails.JobFromShoporderJoistDetails();
            if (job1.Joists == null && job1.Girders == null)
            {
                return;
            }
            List<Joist> joists = job1.Joists;

            foreach (Joist joist in joists)
            {
                if (joist.TCWidth != "5")
                {
                    string markWithWidth = joist.Mark + " (" + joist.TCWidth + ")";
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    Database db = doc.Database;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                            (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                        foreach (ObjectId id in btr)
                        {
                            Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                            if (currentEntity == null)
                            {
                                continue;
                            }
                            if (currentEntity.GetType() == typeof(MText))
                            {
                                string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                                if (strippedText == joist.Mark)
                                {
                                    MText mText = (MText)currentEntity;
                                    string mTextContent = mText.Contents;
                                    string newmTextContent = Regex.Replace(mTextContent, joist.Mark, markWithWidth);
                                    ((MText)currentEntity).Contents = newmTextContent;
                                }
                            }
                            if (currentEntity.GetType() == typeof(DBText))
                            {
                                string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                                if (strippedText == joist.Mark)
                                {
                                    DBText dbText = (DBText)currentEntity;
                                    string dbTextString = dbText.TextString;
                                    string newdbTextString = Regex.Replace(dbTextString, joist.Mark, markWithWidth);
                                    ((DBText)currentEntity).TextString = newdbTextString;
                                }
                            }
                            if (currentEntity.GetType() == typeof(RotatedDimension))
                            {
                                string dimText = ((RotatedDimension)currentEntity).DimensionText;

                                if (dimText.Contains(string.Format("-{0}\\X", joist.Mark)) == true ||
                                    dimText.Contains(string.Format("-{0} ", joist.Mark)))
                                {
                                    string replace = string.Format("-{0}", joist.Mark);
                                    string replacement = string.Format("-{0}", markWithWidth);
                                    string newrdText = Regex.Replace(dimText, replace, replacement);
                                    ((RotatedDimension)currentEntity).DimensionText = newrdText;
                                }

                                string[] dimTextArray = Regex.Split(dimText, "-");
                                string mark = dimTextArray[dimTextArray.Length - 1];

                                if (mark == joist.Mark)
                                {
                                    string replace = string.Format("-{0}", joist.Mark);
                                    string replacement = string.Format("-{0}", markWithWidth);
                                    string newrdText = Regex.Replace(dimText, replace, replacement);
                                    ((RotatedDimension)currentEntity).DimensionText = newrdText;
                                }

                            }

                        }
                        tr.Commit();
                    }
                }
            }
        }
        
        [CommandMethod("MyCommandGroup", "TCWIDTHS_JOISTS_EXCLUDE_7_IN", CommandFlags.Modal)]
        public void tcWidths_joists_no_7_in()
        {
            Job job1 = new Job();
            job1 = joistDetails.JobFromShoporderJoistDetails();
            if (job1.Joists == null && job1.Girders == null)
            {
                return;
            }
            List<Joist> joists = job1.Joists;

            foreach (Joist joist in joists)
            {
                if (joist.TCWidth != "7")
                {
                    string markWithWidth = joist.Mark + " (" + joist.TCWidth + ")";
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    Database db = doc.Database;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                            (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                        foreach (ObjectId id in btr)
                        {
                            Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                            if (currentEntity == null)
                            {
                                continue;
                            }
                            if (currentEntity.GetType() == typeof(MText))
                            {
                                string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                                if (strippedText == joist.Mark)
                                {
                                    MText mText = (MText)currentEntity;
                                    string mTextContent = mText.Contents;
                                    string newmTextContent = Regex.Replace(mTextContent, joist.Mark, markWithWidth);
                                    ((MText)currentEntity).Contents = newmTextContent;
                                }
                            }
                            if (currentEntity.GetType() == typeof(DBText))
                            {
                                string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                                if (strippedText == joist.Mark)
                                {
                                    DBText dbText = (DBText)currentEntity;
                                    string dbTextString = dbText.TextString;
                                    string newdbTextString = Regex.Replace(dbTextString, joist.Mark, markWithWidth);
                                    ((DBText)currentEntity).TextString = newdbTextString;
                                }
                            }
                            if (currentEntity.GetType() == typeof(RotatedDimension))
                            {
                                string dimText = ((RotatedDimension)currentEntity).DimensionText;

                                if (dimText.Contains(string.Format("-{0}\\X", joist.Mark)) == true ||
                                    dimText.Contains(string.Format("-{0} ", joist.Mark)))
                                {
                                    string replace = string.Format("-{0}", joist.Mark);
                                    string replacement = string.Format("-{0}", markWithWidth);
                                    string newrdText = Regex.Replace(dimText, replace, replacement);
                                    ((RotatedDimension)currentEntity).DimensionText = newrdText;
                                }

                                string[] dimTextArray = Regex.Split(dimText, "-");
                                string mark = dimTextArray[dimTextArray.Length - 1];

                                if (mark == joist.Mark)
                                {
                                    string replace = string.Format("-{0}", joist.Mark);
                                    string replacement = string.Format("-{0}", markWithWidth);
                                    string newrdText = Regex.Replace(dimText, replace, replacement);
                                    ((RotatedDimension)currentEntity).DimensionText = newrdText;
                                }

                            }

                        }
                        tr.Commit();
                    }
                }
            }
        }
        
        [CommandMethod("MyCommandGroup", "TCWIDTHS_JOISTS_AND_GIRDERS", CommandFlags.Modal)]
        public void tcWidths_ALL()
        {
            Job job = new Job();
            job = joistDetails.JobFromShoporderJoistDetails();
            if (job.Joists == null && job.Girders == null)
            {
                return;
            }

            var markInfoList = new List<(string Mark, int quantity, string TcWidth, int BoltSize)>();

            foreach (var joist in job.Joists)
            {
                var mark = joist.Mark;
                var tcWidth = joist.TCWidth;
                var bcSize = joist.BC;
                var bcVleg = QueryAngleData.DblVleg(bcSize);
                var isMerchantBc = !bcSize.Contains("A");
                var boltSize = isMerchantBc ?
                                 System.Math.Max((short)3, (short)System.Math.Ceiling(bcVleg + 1)) :
                                 System.Math.Max((short)3, (short)System.Math.Ceiling(bcVleg + (1 - 0.078)));
                markInfoList.Add((mark, joist.Quantity, tcWidth, boltSize));
            }

            var tcWidthMajority =
                markInfoList
                .GroupBy(info => info.TcWidth)
                .Select(group => (TcWidth: group.Key, Sum: group.Sum(info => info.quantity)))
                .OrderByDescending(info => info.Sum)
                .First()
                .TcWidth;


            var marksWithMessages = new List<(string Mark, List<string> Messages)>();

            foreach (var (Mark, Quantity, TcWidth, BoltSize) in markInfoList)
            {
                var messages = new List<string>();
                if (TcWidth != tcWidthMajority) { messages.Add("TCW=" + TcWidth); }
                if (messages.Count != 0)
                {
                    marksWithMessages.Add((Mark, messages));
                }
            }

            foreach (var (Mark, Messages) in marksWithMessages)
            {
                var markWithMessage = string.Format("{0} [{1}]", Mark, string.Join(",", Messages));

                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                        (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                    foreach (ObjectId id in btr)
                    {
                        Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                        if (currentEntity == null)
                        {
                            continue;
                        }
                        if (currentEntity.GetType() == typeof(MText))
                        {
                            string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                            if (strippedText == Mark)
                            {
                                MText mText = (MText)currentEntity;
                                string mTextContent = mText.Contents;
                                string newmTextContent = Regex.Replace(mTextContent, Mark, markWithMessage);
                                ((MText)currentEntity).Contents = newmTextContent;
                            }
                        }
                        if (currentEntity.GetType() == typeof(DBText))
                        {
                            string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                            if (strippedText == Mark)
                            {
                                DBText dbText = (DBText)currentEntity;
                                string dbTextString = dbText.TextString;
                                string newdbTextString = Regex.Replace(dbTextString, Mark, markWithMessage);
                                ((DBText)currentEntity).TextString = newdbTextString;
                            }
                        }
                        if (currentEntity.GetType() == typeof(RotatedDimension))
                        {
                            string dimText = ((RotatedDimension)currentEntity).DimensionText;

                            if (dimText.Contains(string.Format("-{0}\\X", Mark)) == true ||
                                dimText.Contains(string.Format("-{0} ", Mark)))
                            {
                                string replace = string.Format("-{0}", Mark);
                                string replacement = string.Format("-{0}", markWithMessage);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                            string[] dimTextArray = Regex.Split(dimText, "-");
                            string mark = dimTextArray[dimTextArray.Length - 1];

                            if (mark == Mark)
                            {
                                string replace = string.Format("-{0}", Mark);
                                string replacement = string.Format("-{0}", markWithMessage);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                        }

                    }
                    tr.Commit();
                }

            }

            System.Windows.Forms.MessageBox.Show(
                    tcWidthMajority + "\" majority TC width.\n");


            List<Girder> girders = job.Girders;


            foreach (Girder girder in girders)
            {
                string markWithWidth = girder.Mark + " [TCW=" + girder.TCWidth + "]";
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                        (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                    foreach (ObjectId id in btr)
                    {
                        Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                        if (currentEntity == null)
                        {
                            continue;
                        }
                        if (currentEntity.GetType() == typeof(MText))
                        {
                            string strippedText = Regex.Replace(((MText)currentEntity).Contents, " ", "");
                            if (strippedText == girder.Mark)
                            {
                                MText mText = (MText)currentEntity;
                                string mTextContent = mText.Contents;
                                string newmTextContent = Regex.Replace(mTextContent, girder.Mark, markWithWidth);
                                ((MText)currentEntity).Contents = newmTextContent;
                            }
                        }
                        if (currentEntity.GetType() == typeof(DBText))
                        {
                            string strippedText = Regex.Replace(((DBText)currentEntity).TextString, " ", "");
                            if (strippedText == girder.Mark)
                            {
                                DBText dbText = (DBText)currentEntity;
                                string dbTextString = dbText.TextString;
                                string newdbTextString = Regex.Replace(dbTextString, girder.Mark, markWithWidth);
                                ((DBText)currentEntity).TextString = newdbTextString;
                            }
                        }
                        if (currentEntity.GetType() == typeof(RotatedDimension))
                        {
                            string dimText = ((RotatedDimension)currentEntity).DimensionText;

                            if (dimText.Contains(string.Format("-{0}\\X", girder.Mark)) == true ||
                                dimText.Contains(string.Format("-{0} ", girder.Mark)))
                            {
                                string replace = string.Format("-{0}", girder.Mark);
                                string replacement = string.Format("-{0}", markWithWidth);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                            string[] dimTextArray = Regex.Split(dimText, "-");
                            string mark = dimTextArray[dimTextArray.Length - 1];

                            if (mark == girder.Mark)
                            {
                                string replace = string.Format("-{0}", girder.Mark);
                                string replacement = string.Format("-{0}", markWithWidth);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                        }

                    }
                    tr.Commit();
                }
            }
        } 
    

        [CommandMethod("MyCommandGroup", "DESign_CLEAR", CommandFlags.Modal)]

        */

        public void removeDesignInfo()
        {

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject
                    (SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    Entity currentEntity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
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
            if (text.Contains("["))
            {
                text = text.Substring(0, text.IndexOf('['));
            }
            if (text.Contains("("))
            {
                text = text.Substring(0, text.IndexOf('('));
            }

            return text;

        }
    }


}
