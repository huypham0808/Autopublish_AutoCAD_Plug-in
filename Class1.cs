using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.PlottingServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Diagnostics;

namespace SST_SDD
{
    public class Class1
    {
        [CommandMethod("SDD")]
        public static void PublishPDF()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            short bgPlot = (short)Application.GetSystemVariable("BACKGROUNDPLOT");
            Application.SetSystemVariable("BACKGROUNDPLOT", 0);

            // Get the layout ObjectId List
            List<ObjectId> layoutIds = GetLayoutIds(db);

            string dwgFileName = (string)Application.GetSystemVariable("DWGNAME");
            string dwgPath = (string)Application.GetSystemVariable("DWGPREFIX");

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                DsdEntryCollection collection = new DsdEntryCollection();

                foreach (ObjectId layoutId in layoutIds)
                {
                    Layout layout = trans.GetObject(layoutId, OpenMode.ForRead) as Layout;

                    if (layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;
                    DsdEntry entry = new DsdEntry();
                    entry.DwgName = dwgPath + dwgFileName;
                    entry.Layout = layout.LayoutName;
                    entry.Title = "Layout_" + layout.LayoutName;
                    entry.NpsSourceDwg = entry.DwgName;
                    entry.Nps = GetCurrentNps();

                    collection.Add(entry);
                }
                //Lay ten file loai bo .dwg
                dwgFileName = dwgFileName.Substring(0, dwgFileName.Length - 4);
                DsdData dsdData = new DsdData();
                dsdData.SheetType = SheetType.MultiPdf;
                dsdData.ProjectPath = dwgPath;
                dsdData.DestinationName = dsdData.ProjectPath + dwgFileName + ".pdf";              
                try
                {
                    if (System.IO.File.Exists(dsdData.DestinationName))
                    System.IO.File.Delete(dsdData.DestinationName);
                }
                catch(Autodesk.AutoCAD.Runtime.Exception ex)
                {                    
                    ed.WriteMessage("Please close the current PDF file before printing.");
                    return;
                }
                dsdData.SetDsdEntryCollection(collection);
                string dsdFile = dsdData.ProjectPath + dwgFileName + ".dsd";
                dsdData.WriteDsd(dsdFile);
                System.IO.StreamReader sr = new System.IO.StreamReader(dsdFile);
                string str = sr.ReadToEnd();
                sr.Close();
                str = str.Replace("PromptForDwfName=TRUE", "PromptForDwfName=FALSE");

                System.IO.StreamWriter sw = new System.IO.StreamWriter(dsdFile);
                sw.Write(str);
                sw.Close();

                dsdData.ReadDsd(dsdFile);
                System.IO.File.Delete(dsdFile);

                PlotConfig plotConfig = PlotConfigManager.SetCurrentConfig("DWG To PDF.pc3");
                Autodesk.AutoCAD.Publishing.Publisher publisher = Autodesk.AutoCAD.ApplicationServices.Application.Publisher;
                publisher.AboutToBeginPublishing += new Autodesk.AutoCAD.Publishing.AboutToBeginPublishingEventHandler(Publisher_AboutToBeginPublishing);
                publisher.PublishExecute(dsdData, plotConfig);
                
                trans.Commit();
                Process.Start(dsdData.DestinationName);
            }
            Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);
            
        }
        private static List<ObjectId> GetLayoutIds(Database db)
        {
            List<ObjectId> layoutIds = new List<ObjectId>();

            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDic = Tx.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false) as DBDictionary;

                foreach (DBDictionaryEntry entry in layoutDic)
                {
                    layoutIds.Add(entry.Value);
                }
            }

            return layoutIds;
        }
        static void Publisher_AboutToBeginPublishing(object sender, Autodesk.AutoCAD.Publishing.AboutToBeginPublishingEventArgs e)
        {
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nAboutToBeginPublishing!!");
        }
        public static string GetCurrentNps()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayoutManager layManager = LayoutManager.Current;

                ObjectId layoutId = layManager.GetLayoutId(layManager.CurrentLayout);
                Layout layoutObj = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                PlotSettings plotSettings = tr.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForRead) as PlotSettings;
                string currentPlotStyleName = layoutObj.CurrentStyleSheet;

                tr.Commit();

                return currentPlotStyleName;
            }
        }
    }
}
