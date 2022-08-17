using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace BL
{
    public class DataExportCSV
    {

        public void GetListAttributes()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            Database db = HostApplicationServices.WorkingDatabase;

            Transaction tr = db.TransactionManager.StartTransaction();

            int count1 = 0;
            // Start the transaction

            try
            {

                // Build a filter list so that only

                // block references are selected


                TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };

                SelectionFilter filter = new SelectionFilter(filList);

                PromptSelectionOptions opts = new PromptSelectionOptions();

                opts.MessageForAdding = Utility.Constraint.SelectBlockReference;

                PromptSelectionResult res = ed.GetSelection(opts, filter);

                // Do nothing if selection is unsuccessful

                if (res.Status != PromptStatus.OK)

                    return;

                SelectionSet selSet = res.Value;


                ObjectId[] idArray = selSet.GetObjectIds();
                List<BillsOfmaterial> billsOfmaterials = new List<BillsOfmaterial>();

                foreach (ObjectId blkId in idArray)
                {

                    BillsOfmaterial billsOfmaterial = new BillsOfmaterial();

                    BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);

                    //ed.WriteMessage("\nBlock: " + btr.Name); code commented by ravi 

                    btr.Dispose();


                    AttributeCollection attCol = blkRef.AttributeCollection;

                    int count = 0;


                    foreach (ObjectId attId in attCol)
                    {

                        AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);

                        string str = ("\n  Attribute Tag: " + attRef.Tag + "\n    Attribute String: " + attRef.TextString);

                        if (attRef.Tag == "PARTNUMBER")
                            billsOfmaterial.PARTNUMBER = attRef.TextString;
                        if (attRef.Tag == "VARIANT_2_CODE")
                            billsOfmaterial.VARIANT_2_CODE = attRef.TextString;
                        if (attRef.Tag == "INFO")
                            billsOfmaterial.INFO = attRef.TextString;
                        //if (count == 3)
                        //    billsOfmaterial.Quantity = count1;

                        if (attRef.Tag == "LABEL")
                            billsOfmaterial.PanelLABEL = attRef.TextString;

                        var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                        var brclass = RXObject.GetClass(typeof(BlockReference));

                        var blocks = modelSpace.Cast<ObjectId>().Where(id => id.ObjectClass == brclass).
                         Select(Id => (BlockReference)tr.GetObject(Id, OpenMode.ForRead)).
                         GroupBy(br => ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name);

                        blocks.Count();
                        foreach (var group in blocks)
                        {

                            var istrue = group.Any(p => p.ObjectId == blkId);
                            if (istrue)
                            {
                                billsOfmaterial.Quantity = group.Count();
                            }

                            count1 = group.Count();

                        }

                        // string str = ("\n  Attribute Tag: " + attRef.Tag + "\n    Attribute String: " + attRef.TextString);

                        //   ed.WriteMessage(str); 

                        count++;
                    }
                    billsOfmaterials.Add(billsOfmaterial);
                }
                //------------------------
                //billsOfmaterials.GroupBy(p => p.PARTNUMBER);


                var t1 = billsOfmaterials.Where(x => !string.IsNullOrEmpty(x.VARIANT_2_CODE)).ToList();
               
                var aa = billsOfmaterials.GroupBy(x => x.PARTNUMBER);
                List<BillsOfmaterial> billsOfmaterials_test = new List<BillsOfmaterial>();
                foreach (var item in aa)
                {
                    var bb = item.Where(p => p.PARTNUMBER == item.Key).GroupBy(x => x.VARIANT_2_CODE);

                    foreach (var a in bb)
                    {
                        BillsOfmaterial cc = new BillsOfmaterial();
                        cc.PARTNUMBER = item.Key;
                        cc.VARIANT_2_CODE = a.Key;

                        cc.Quantity = string.IsNullOrEmpty(a.Key) ? a.FirstOrDefault().Quantity - billsOfmaterials.Where(x => !string.IsNullOrEmpty(x.VARIANT_2_CODE) && x.PARTNUMBER==item.Key).Count() : a.Count();
                        if (cc.Quantity == 0)
                            cc.Quantity = 1;
                       
                        billsOfmaterials_test.Add(cc);
                    }
                }


                var aa1 = billsOfmaterials.GroupBy(x => x.PARTNUMBER);
                //List<BillsOfmaterial> billsOfmaterials_test = new List<BillsOfmaterial>();
                foreach (var item1 in aa1)
                {
                    var bb1 = item1.Where(p => p.PARTNUMBER == item1.Key && p.Quantity != item1.FirstOrDefault().Quantity).GroupBy(x => x.VARIANT_2_CODE);

                    foreach (var a1 in bb1)
                    {
                        BillsOfmaterial cc1 = new BillsOfmaterial();
                        cc1.PARTNUMBER = item1.Key;
                        cc1.VARIANT_2_CODE = a1.Key;

                        cc1.Quantity = string.IsNullOrEmpty(a1.Key) ? a1.FirstOrDefault().Quantity - billsOfmaterials.Where(x => !string.IsNullOrEmpty(x.VARIANT_2_CODE) && x.PARTNUMBER == item1.Key).Count() : a1.Count();
                        if (cc1.Quantity == 0)
                            cc1.Quantity = 1;

                        billsOfmaterials_test.Add(cc1);
                    }
                }



                ExportCsv(billsOfmaterials_test);
                // ExportCsv(billsOfmaterials);
                tr.Commit();

            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
            finally
            {
                tr.Dispose();
            }
        }

        public static void ExportCsv<T>(List<T> genericList)
        {
            try
            {
                var sb = new StringBuilder();

                var saveFileDialog = new Microsoft.Win32.SaveFileDialog()
                {
                    DefaultExt = "*.csv",
                    Filter = "Text Files (*.csv)|*.csv",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                };
                var lines = new List<string>();
                var result = saveFileDialog.ShowDialog();
                string path = saveFileDialog.FileName;
                if (result != null && result == true)
                {

                    File.WriteAllLines(path, lines.ToArray());
                }

                var finalPath = path; //Path.Combine(basePath, fileName + Utility.Constraint.CSVFileExtension);
                var header = "";
                var info = typeof(T).GetProperties();

                foreach (var prop in typeof(T).GetProperties())
                {
                    header += prop.Name + ", ";
                }

                header = header.Substring(0, header.Length - 2);
                sb.AppendLine(header);
                TextWriter sw = new StreamWriter(finalPath, true);
                sw.Write(sb.ToString());
                sw.Close();

                foreach (var obj in genericList)
                {
                    sb = new StringBuilder();
                    var line = "";
                    foreach (var prop in info)
                    {
                        line += prop.GetValue(obj, null) + ", ";
                    }
                    line = line.Substring(0, line.Length - 2);
                    sb.AppendLine(line);
                    TextWriter sw1 = new StreamWriter(finalPath, true);
                    sw1.Write(sb.ToString());
                    sw1.Close();
                }
            }
            catch (System.Exception ex)
            {
                throw;
            }
        }
    }

}
