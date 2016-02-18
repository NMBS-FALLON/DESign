using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Text.RegularExpressions;
using DESign_BASE;

[assembly: CommandClass(typeof(DESign_AutoCAD.MyCommands))]

namespace DESign_AutoCAD
{

    public class MyCommands
    {
      
        DESign_BASE.ExtractJoistDetails joistDetails = new ExtractJoistDetails();

        [CommandMethod("MyCommandGroup", "TCWIDTHS_JOISTS", CommandFlags.Modal)]
        public void tcWidths_Joists()
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

                            if (dimText.Contains(string.Format("-{0}\\X",joist.Mark)) == true)
                            {
                                string replace = string.Format("-{0}", joist.Mark);
                                string replacement = string.Format("-{0}", markWithWidth);
                                string newrdText = Regex.Replace(dimText, replace, replacement);
                                ((RotatedDimension)currentEntity).DimensionText = newrdText;
                            }

                            string[] dimTextArray = Regex.Split(dimText, "-");
                            string mark = dimTextArray[dimTextArray.Length - 1];

                            if (mark==joist.Mark)
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

        [CommandMethod("MyCommandGroup", "TCWIDTHS_GIRDERS", CommandFlags.Modal)]
        public void tcWidths_Girders()
        {
            Job job1 = new Job();
            job1 = joistDetails.JobFromShoporderJoistDetails();
            if (job1.Joists == null && job1.Girders == null)
            {
                return;
            }
            List<Girder> girders = job1.Girders;


            foreach (Girder girder in girders)
            {
                string markWithWidth = girder.Mark + " (" + girder.TCWidth + ")";
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

                            if (dimText.Contains(string.Format("-{0}\\X", girder.Mark)) == true)
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

        [CommandMethod("MyCommandGroup", "TEST", CommandFlags.Modal)]
        public void test()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
        }

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

                                if (dimText.Contains(string.Format("-{0}\\X", joist.Mark)) == true)
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

                                if (dimText.Contains(string.Format("-{0}\\X", joist.Mark)) == true)
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

        [CommandMethod("MyCommandGroup", "TCWIDTHS_ALL", CommandFlags.Modal)]
        public void tcWidths_ALL()
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

                            if (dimText.Contains(string.Format("-{0}\\X", joist.Mark)) == true)
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

            List<Girder> girders = job1.Girders;


            foreach (Girder girder in girders)
            {
                string markWithWidth = girder.Mark + " (" + girder.TCWidth + ")";
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

                            if (dimText.Contains(string.Format("-{0}\\X", girder.Mark)) == true)
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

    }
}
