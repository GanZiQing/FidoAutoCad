using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using FidoAutoCad.Forms;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Exception = System.Exception;

using System.Windows;

[assembly: CommandClass(typeof(FidoAutoCad.AutoCADCommands))]
namespace FidoAutoCad
{
    public class AutoCADCommands
    {
        [CommandMethod("AdskGreeting")]
        public void AdskGreeting()
        {
            string msgText = "test inline";
            try
            {
                // Get the current document and database, and start a transaction
                Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (acDoc.LockDocument())
                {
                    // Starts a new transaction with the Transaction Manager
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        // Open the Block table record for read
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                     OpenMode.ForRead) as BlockTable;

                        // Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                        OpenMode.ForWrite) as BlockTableRecord;

                        /* Creates a new MText object and assigns it a location,
                        text value and text style */
                        using (MText objText = new MText())
                        {
                            // Specify the insertion point of the MText object
                            objText.Location = new Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0);

                            // Set the text string for the MText object
                            objText.Contents = msgText;

                            // Set the text style for the MText object
                            objText.TextStyleId = acCurDb.Textstyle;

                            // Appends the new MText object to model space
                            acBlkTblRec.AppendEntity(objText);

                            // Appends to new MText object to the active transaction
                            acTrans.AddNewlyCreatedDBObject(objText, true);
                        }

                        // Saves the changes to the database and closes the transaction
                        acTrans.Commit();
                        acDoc.Editor.Regen();
                    }
                }
                
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

        }

        public void PrintText(string msgText)
        {
            try
            {
                // Get the current document and database, and start a transaction
                Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (acDoc.LockDocument())
                {
                    // Starts a new transaction with the Transaction Manager
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        // Open the Block table record for read
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                     OpenMode.ForRead) as BlockTable;

                        // Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                        OpenMode.ForWrite) as BlockTableRecord;

                        /* Creates a new MText object and assigns it a location,
                        text value and text style */
                        using (MText objText = new MText())
                        {
                            // Specify the insertion point of the MText object
                            objText.Location = new Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0);

                            // Set the text string for the MText object
                            objText.Contents = msgText;

                            // Set the text style for the MText object
                            objText.TextStyleId = acCurDb.Textstyle;

                            // Appends the new MText object to model space
                            acBlkTblRec.AppendEntity(objText);

                            // Appends to new MText object to the active transaction
                            acTrans.AddNewlyCreatedDBObject(objText, true);
                        }

                        // Saves the changes to the database and closes the transaction
                        acTrans.Commit();
                        acDoc.Editor.Regen();
                    }
                }

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

        }

        //[CommandMethod("FidoCad")]
        public void FidoCadApp()
        {
            try
            {
                FidoAutoCadMain fidoForm = new FidoAutoCadMain(this);
                fidoForm.Show();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message,"Error"); }
        }

        #region FidoDock
        static FidoAutocadDock fidoUserControl;
        static PaletteSet fidoPalletSet;
        [CommandMethod("Fido")]
        public void FidoDock()
        {
            try
            {
                if (fidoUserControl == null)
                {
                    fidoUserControl = new FidoAutocadDock(this);
                    fidoPalletSet = new PaletteSet("Fido AutoCAD", new Guid("87374E16-C0DB-4F3F-9271-7A71ED921566"));
                    var controlSize = fidoUserControl.Size;
                    fidoPalletSet = new PaletteSet("Fido AutoCAD");
                    fidoPalletSet.Add("Fido Form", fidoUserControl);
                    fidoPalletSet.DockEnabled = (DockSides.Left | DockSides.Right);
                    fidoPalletSet.Visible = true;
                    fidoPalletSet.Size = Size.Add(controlSize, new Size(100,0));
                }
                else
                {
                    fidoPalletSet.Visible = !fidoPalletSet.Visible;
                }
                    
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion
        #region Shortcut functions
        [CommandMethod("FidoGetProperties", CommandFlags.UsePickSet)]
        public void FidoGetProperties()
        {
            try
            {
                if (fidoUserControl == null) { throw new Exception("Initialise fido palette first"); }
                fidoUserControl.GetPropButton_Click(null, null);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

    }
}
