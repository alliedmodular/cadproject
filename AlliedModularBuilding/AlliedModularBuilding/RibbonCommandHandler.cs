using System;
using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using BL;
using NuGet.Protocol.Plugins;

namespace AlliedModularBuilding
{
    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            DataExportCSV dataExportCSV = new DataExportCSV();
            try
            {
                Document doc = acadApp.DocumentManager.MdiActiveDocument;

                if (parameter is RibbonButton)
                {
                    RibbonButton button = parameter as RibbonButton;

                    if (button.Id == Utility.Constraint.ButtonBOMID)
                    {
                        dataExportCSV.GetListAttributes();
            
        
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw;
            }
        }

            public void GetListAttributes()
            {

                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

                Database db = HostApplicationServices.WorkingDatabase;

                Transaction tr = db.TransactionManager.StartTransaction();


                // Start the transaction

                try
                {


                    // Build a filter list so that only

                    // block references are selected

                    TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };

                    SelectionFilter filter = new SelectionFilter(filList);

                    PromptSelectionOptions opts = new PromptSelectionOptions();

                    opts.MessageForAdding = "Select block references: ";

                    PromptSelectionResult res = ed.GetSelection(opts, filter);


                    // Do nothing if selection is unsuccessful

                    if (res.Status != PromptStatus.OK)

                        return;

                    SelectionSet selSet = res.Value;

                    ObjectId[] idArray = selSet.GetObjectIds();

                    foreach (ObjectId blkId in idArray)
                    {
                        BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);

                        ed.WriteMessage("\nBlock: " + btr.Name);

                        btr.Dispose();

                        AttributeCollection attCol = blkRef.AttributeCollection;

                        foreach (ObjectId attId in attCol)
                        {
                            AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);

                            string str = ("\n  Attribute Tag: " + attRef.Tag + "\n    Attribute String: " + attRef.TextString);

                            ed.WriteMessage(str);
                        }
                    }

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
      

    }
}
