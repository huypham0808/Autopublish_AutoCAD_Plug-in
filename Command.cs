using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Publishing;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Diagnostics;
[assembly: CommandClass(typeof(AutoPrint.Class1))]
namespace AutoPrint
{
    public class Class1
    {
        public bool Visible { get; private set; }

        [CommandMethod("PCCOLD")]
        public static void PublishPDF()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            short bgPlot = (short)Application.GetSystemVariable("BACKGROUNDPLOT");
            Application.SetSystemVariable("BACKGROUNDPLOT", 0);

            // Get the layout ObjectId List
            

            string dwgFileName = (string)Application.GetSystemVariable("DWGNAME");
            string dwgPath = (string)Application.GetSystemVariable("DWGPREFIX");

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                List<ObjectId> layoutIds = GetLayoutIds(db, trans);
                DsdEntryCollection collection = new DsdEntryCollection();

                foreach (ObjectId layoutId in layoutIds)
                {
                    Layout layout = trans.GetObject(layoutId, OpenMode.ForRead) as Layout;

                    if (layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;
                    using(DsdData dsdFileData = new DsdData())
                    {
                        dsdFileData.DestinationName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\";
                    }
                    DsdEntry entry = new DsdEntry();
                    entry.DwgName = dwgPath + dwgFileName;
                    entry.Layout = layout.LayoutName;
                    entry.Title = layout.LayoutName;
                    entry.NpsSourceDwg = entry.DwgName;
                    //entry.Nps = "STN plot style";
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
                    MessageBox.Show("Please close the current PDF file","AutoCAD", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                dsdData.SetDsdEntryCollection(collection);
                string logFilePath = dsdData.LogFilePath;
                logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + dwgFileName +".txt";

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

        private static List<ObjectId> GetLayoutIds(Database db, Transaction tr)
        {
            List<ObjectId> layoutIds = new List<ObjectId>();

            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDic = Tx.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false) as DBDictionary;
                int tabOrder = 0;
                foreach (DBDictionaryEntry entry in layoutDic)
                {
                    layoutIds.Add(entry.Value);                  
                }
               
            }
            layoutIds = layoutIds.OrderBy(id => ((Layout)tr.GetObject(id, OpenMode.ForRead)).TabOrder).ToList();
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
        //NEW VERSION 05/05/2024
        [CommandMethod("PCC", CommandFlags.Modal)]
        public void AutoPublishLayout()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            string mypath, pdfFile, dsdFile, txtFile, dir = string.Empty;
    
            mypath = doc.Name; //full path with autocad file name + extension
            pdfFile = Path.ChangeExtension(mypath,"pdf");
            dsdFile = Path.ChangeExtension(mypath, "dsd");
            txtFile = Path.ChangeExtension(mypath, "txt");
            dir = Path.GetDirectoryName(mypath);

            if (File.Exists(pdfFile)) File.Delete(pdfFile);
            if (File.Exists(dsdFile)) File.Delete(dsdFile);
            if (File.Exists(txtFile)) File.Delete(txtFile);
            if (File.Exists(dir + "\\plot.log")) File.Delete(dir + "\\plot.log");

            using (var lc = doc.LockDocument())
            {
                var myList = new List<Layout>();
                LayoutManager lm = LayoutManager.Current;
                using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                {
                    db.UpdateExt(true);
                    var lays = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    foreach (var item in lays)
                    {
                        //if (item.Key.ToUpper() == "Model") continue;
                        var lay1 = tr.GetObject(lm.GetLayoutId(item.Key), OpenMode.ForRead) as Layout;
                        if (lay1.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;
                        myList.Add(lay1);
                    }                 
                }
                myList.Sort((a, b) => a.TabOrder.CompareTo(b.TabOrder));
                using (DsdEntryCollection dsdDwgFiles = new DsdEntryCollection())
                {
                    foreach (var lay in myList)
                    {
                        using (DsdEntry dsdDwgFile1 = new DsdEntry())
                        {
                            // Set the file name and layout
                            dsdDwgFile1.DwgName = mypath;
                            dsdDwgFile1.Layout = lay.LayoutName;
                            dsdDwgFile1.Title = lay.LayoutName;

                            // Set the page setup override
                            dsdDwgFile1.Nps = GetCurrentNps();
                            dsdDwgFile1.NpsSourceDwg = dsdDwgFile1.DwgName;

                            dsdDwgFiles.Add(dsdDwgFile1);
                        }
                    }
                    myList.Clear();



                    // Set the properties for the DSD file and then write it out
                    using (DsdData dsdFileData = new DsdData())
                    {
                        // Set the target information for publishing
                        dsdFileData.DestinationName = pdfFile;
                        dsdFileData.ProjectPath = dir;
                        dsdFileData.SheetType = SheetType.MultiPdf;

                        // Set the drawings that should be added to the publication
                        dsdFileData.SetDsdEntryCollection(dsdDwgFiles);
                        // Set the general publishing properties
                        dsdFileData.LogFilePath = txtFile;

                        // Create the DSD file
                        dsdFileData.WriteDsd(dsdFile);

                        try
                        {
                            // Publish the specified drawing files in the DSD file, and
                            // honor the behavior of the BACKGROUNDPLOT system variable

                            using (DsdData dsdDataFile = new DsdData())
                            {
                                System.IO.StreamReader sr = new System.IO.StreamReader(dsdFile);
                                string str = sr.ReadToEnd();
                                sr.Close();
                                str = str.Replace("PromptForDwfName=TRUE", "PromptForDwfName=FALSE");
                                System.IO.StreamWriter sw = new System.IO.StreamWriter(dsdFile);
                                sw.Write(str);
                                sw.Close();
                                dsdDataFile.ReadDsd(dsdFile);

                                // Get the DWG to PDF.pc3 and use it as a 
                                // device override for all the layouts
                                PlotConfig acPlCfg = PlotConfigManager.SetCurrentConfig("DWG to PDF.PC3");
                                Autodesk.AutoCAD.Publishing.Publisher publisher = Autodesk.AutoCAD.ApplicationServices.Application.Publisher;
                                publisher.AboutToBeginPublishing += new Autodesk.AutoCAD.Publishing.AboutToBeginPublishingEventHandler(Publisher_AboutToBeginPublishing);
                                publisher.PublishExecute(dsdDataFile, acPlCfg);
                                Process.Start(dsdDataFile.DestinationName);
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception es)
                        {
                            System.Windows.Forms.MessageBox.Show(es.Message);
                        }
                    }
                }
                if (File.Exists(dsdFile)) File.Delete(dsdFile);
                if (File.Exists(txtFile)) File.Delete(txtFile);
                if (File.Exists(dir + "\\plot.log")) File.Delete(dir + "\\plot.log");
            }        
        }
    }
}
