using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal.PropertyInspector;
using Autodesk.AutoCAD.Runtime;
using FidoAutoCad.SharedForms;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static FidoAutoCad.CommonUtilities;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDB = Autodesk.AutoCAD.DatabaseServices;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;
using static FidoAutoCad.AcadSharedFunctions;

namespace FidoAutoCad.Forms
{
    public partial class FidoAutocadDock : UserControl
    {
        #region Init
        AutoCADCommands parent;
        Dictionary<string, object> attributes = new Dictionary<string, object>();
        public FidoAutocadDock(AutoCADCommands parent)
        {
            InitializeComponent();
            this.parent = parent;
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            CreateAttributes();

            // Hook to detacher
            AcadApp.QuitWillStart += detachExcel_Auto;
        }
        Excel.Application excelApp = null;
        private void CreateAttributes()
        {
            roundTypeComboBox.Text = "nearest";

            AttributeTextBox tbAtt = new AttributeTextBox("roundOptions_AC", dispRoundOpt, true);
            tbAtt.type = "double";
            tbAtt.SetDefaultValue("1");
            attributes.Add(tbAtt.attName, tbAtt);

            tbAtt = new AttributeTextBox("distConvOptions_AC", dispDistConv, true);
            tbAtt.type = "double";
            attributes.Add(tbAtt.attName, tbAtt);

            tbAtt = new AttributeTextBox("areaConvOptions_AC", dispAreaConv, true);
            tbAtt.type = "double";
            attributes.Add(tbAtt.attName, tbAtt);
                        
            //label1.Focus(); // Cannot figure out how to not select the combo box upon load lol
        }
        #endregion

        #region GetProperties
        public void GetPropButton_Click(object sender, EventArgs e)
        {
            GetProperties(printProperties);
        }
        public void GetProperties(PrintProperties printProperties)
        {
            try
            {
                #region Check Excel
                if (excelApp == null) { throw new Exception("No instance of excel attached"); }
                else if (excelApp.ActiveWorkbook == null) { throw new Exception("No active workbook found"); }
                #endregion

                #region Get Rounding and Conversion Factors
                {
                    (double roundFactor, double distConvFactor, double areaConvFactor) = GetConversionFactors();
                    printProperties.SetConversionFactors(roundFactor, distConvFactor, areaConvFactor);
                }
                #endregion

                #region Autocad Transaction
                (Document acDoc, Editor editor, Database acDb) = GetAcadDoc();
                using (acDoc.LockDocument())
                {
                    // Starts a new transaction with the Transaction Manager
                    using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                    {
                        PromptSelectionResult selectionResult = editor.SelectImplied();
                        if (selectionResult.Status != PromptStatus.OK || selectionResult.Value.Count == 0) { selectionResult = editor.GetSelection(); }

                        if (selectionResult.Status == PromptStatus.OK && selectionResult.Value.Count > 0)
                        {
                            SelectionSet selectedObjs = selectionResult.Value;
                            printProperties.InitPropertyArrays(selectedObjs.Count);

                            int objectNum = 0;
                            foreach (SelectedObject selObj in selectedObjs)
                            {
                                dynamic dynEnt = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;

                                foreach (KeyValuePair<string, ListItem> propertyEntry in printProperties.properties)
                                {
                                    AcadPropertyType acadProp = (AcadPropertyType)propertyEntry.Value;
                                    if (!acadProp.isActive) { continue; }
                                    object arrayEntry = "Unfilled";
                                    switch (propertyEntry.Key)
                                    {
                                        case "Type":
                                            {
                                                try
                                                {
                                                    arrayEntry = (string)dynEnt.AcadObject.EntityName;
                                                }
                                                catch { arrayEntry = "NA"; }
                                                break;
                                            }

                                        case "Length":
                                            {
                                                try
                                                {
                                                    arrayEntry = (double)dynEnt.Length;
                                                }
                                                catch { arrayEntry = "NA"; }
                                                break;
                                            }
                                        case "Area":
                                            {
                                                try
                                                {
                                                    arrayEntry = (double)dynEnt.Area;
                                                }
                                                catch { arrayEntry = "NA"; }
                                                break;
                                            }
                                        default:
                                            throw new Exception($"Property type {propertyEntry.Key} not found");
                                    }

                                    acadProp.contentArray[objectNum] = arrayEntry;
                                }
                                objectNum++;
                            }
                        }
                        else { return; }
                    }
                }
                #endregion

                #region Convert and Round Values
                printProperties.ConvertAndRoundAll();
                #endregion

                #region Write to Excel
                printProperties.PrintToRange(excelApp, null);
                excelApp.ScreenUpdating = true;
                MessageBox.Show("Completed", "Fido AutoCAD");
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        public (Document acDoc, Editor editor, Database acDb) GetAcadDoc()
        {
            Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor editor = acDoc.Editor;
            Database acDb = acDoc.Database;
            try
            {
                acDoc = AcadApp.DocumentManager.MdiActiveDocument;
                editor = acDoc.Editor;
                acDb = acDoc.Database;
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable to attach active autocad document ${ex.Message}");
            }
            return (acDoc, editor, acDb);
        }
        #endregion

        #region Print Properties
        public PrintProperties printProperties = new PrintProperties();
        private void showPrintOptions_Click(object sender, EventArgs e)
        {
            ListSelector optionsForm = new ListSelector("Properties Option", "Select properties to output", "Available properties", "Selected properties");
            optionsForm.AddTracker(ref printProperties.properties);
            optionsForm.ShowDialog();
        }
        
        private (double roundFactor, double distConvFactor, double areaConvFactor) GetConversionFactors()
        {
            double roundFactor = double.NaN;
            double distConvFactor = double.NaN;
            double areaConvFactor = double.NaN;
            if (dispRoundOpt.Text != "")
            {
                roundFactor = ((AttributeTextBox)attributes["roundOptions_AC"]).GetDoubleFromTextBox();
            }
            if (dispDistConv.Text != "")
            {
                distConvFactor = ((AttributeTextBox)attributes["distConvOptions_AC"]).GetDoubleFromTextBox();
            }
            if (dispAreaConv.Text != "")
            {
                areaConvFactor = ((AttributeTextBox)attributes["areaConvOptions_AC"]).GetDoubleFromTextBox();
            }
            return (roundFactor, distConvFactor, areaConvFactor);
        }

        #endregion

        #region Get Coordinates - old
        private void getUcsButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                CheckIfExcelIsAttached();
                #endregion

                #region Autocad Transaction
                List<double[,]> allCoords = new List<double[,]>();
                List<string> allTypes = new List<string>();
                
                (Document acDoc, Editor editor, Database acDb) = GetAcadDoc();
                using (acDoc.LockDocument())
                {
                    try
                    {
                        using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                        {
                            ObjectId[] selectedObjIds = null;
                            try { selectedObjIds = GetAcadSelectedObjectID(acTrans, editor, skipLockCheck.Checked); }
                            catch (Exception ex)
                            {
                                editor.WriteMessage("Fido: Operation terminated. " + ex.Message + Environment.NewLine); return;
                            }

                            int skippedInvalidCount = 0;
                            foreach (ObjectId objId in selectedObjIds)
                            {
                                dynamic dynEnt = acTrans.GetObject(objId, OpenMode.ForRead) as Entity;
                                double[,] coord = GetCoordinates(dynEnt);
                                
                                if (skipInvalidCheck.Checked && double.IsNaN(coord[0,0]))
                                {
                                    skippedInvalidCount++;
                                    continue;
                                }

                                allCoords.Add(coord);
                                allTypes.Add(dynEnt.AcadObject.ObjectName);
                            }
                            editor.WriteMessage($"Fido: {skippedInvalidCount} invalid object(s)." + Environment.NewLine);
                        }
                    }
                    catch (Exception ex) { throw new Exception("Error in autocad transaction" + ex.Message); }
                }
                #endregion

                #region Form print object
                if (allCoords.Count == 0) { throw new Exception("Nothing found to print"); }
                // Get number of rows
                int totalRows = 0;
                foreach (double[,] coord in allCoords)
                {
                    totalRows += coord.GetLength(0);
                }
                object[,] printObject = new object[totalRows, 5];

                int currentRowNum = 0;
                int objectNum = 0;

                foreach (double[,] coord in allCoords)
                {
                    // Column 1 - object number and type
                    printObject[currentRowNum,0] = $"{objectNum + 1}.{allTypes[objectNum]}";
                    objectNum += 1;

                    for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                    {
                        // Column 2 - number
                        printObject[currentRowNum, 1] = coordRowNum + 1;

                        // Column 3 to 5 - coordinates
                        printObject[currentRowNum, 2] = coord[coordRowNum, 0];
                        printObject[currentRowNum, 3] = coord[coordRowNum, 1];
                        if (coord.GetLength(1) > 2)
                        {
                            printObject[currentRowNum, 4] = coord[coordRowNum, 2];
                        }
                        currentRowNum += 1;
                    }
                }

                ReplaceNaN(ref printObject);
                #endregion

                #region Write to Excel
                editor.WriteMessage($"Fido: Begin write to excel." + Environment.NewLine);
                WriteObjectToExcelRange(excelApp, null, 0, 0, true, printObject);
                editor.WriteMessage($"Fido: Coordinates written to excel. Operation completed." + Environment.NewLine);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }

        }
        #endregion

        #region Get Coordinates or Mid Points
        private void getCoordinatesButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                CheckIfExcelIsAttached();
                (double roundFactor, double distConvFactor, double areaConvFactor) = GetConversionFactors();
                #endregion

                #region Get Coordinates in WCS
                List<double[,]> allCoords = new List<double[,]>();
                List<string> allTypes = new List<string>();

                (Document acDoc, Editor editor, Database acDb) = GetAcadDoc();
                using (acDoc.LockDocument())
                {
                    try
                    {
                        using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                        {
                            ObjectId[] selectedObjIds = null;
                            try { selectedObjIds = GetAcadSelectedObjectID(acTrans, editor, skipLockCheck.Checked); }
                            catch (Exception ex)
                            {
                                editor.WriteMessage("Fido: Operation terminated. " + ex.Message + Environment.NewLine); return;
                            }

                            int skippedInvalidCount = 0;
                            foreach (ObjectId objId in selectedObjIds)
                            {
                                dynamic dynEnt = acTrans.GetObject(objId, OpenMode.ForRead) as Entity;
                                double[,] coord = GetCoordinates(dynEnt);

                                if (skipInvalidCheck.Checked && double.IsNaN(coord[0, 0]))
                                {
                                    skippedInvalidCount++;
                                    continue;
                                }

                                allCoords.Add(coord);
                                allTypes.Add(dynEnt.AcadObject.ObjectName);
                            }
                            editor.WriteMessage($"Fido: {skippedInvalidCount} invalid object(s)." + Environment.NewLine);
                        }
                    }
                    catch (Exception ex) { throw new Exception("Error in autocad transaction" + ex.Message); }
                }
                #endregion

                #region Transform in UCS
                if (allCoords.Count == 0) { throw new Exception("Nothing found to print"); }
                if (translateByUcsCheck.Checked) { TransformListOfCoordsByCs(ref allCoords, editor); }
                #endregion

                #region Convert Distances
                if (!double.IsNaN(distConvFactor))
                {
                    foreach (double[,] coord in allCoords)
                    {
                        for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                        {
                            for (int coordColNum = 0; coordColNum < coord.GetLength(1); coordColNum++)
                            {
                                coord[coordRowNum, coordColNum] = coord[coordRowNum, coordColNum] * distConvFactor;
                            }
                        }
                    }
                }
                #endregion

                #region Round Values
                if (!double.IsNaN(roundFactor))
                {
                    foreach (double[,] coord in allCoords)
                    {
                        for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                        {
                            for (int coordColNum = 0; coordColNum < coord.GetLength(1); coordColNum++)
                            {
                                coord[coordRowNum, coordColNum] = RoundDouble(coord[coordRowNum, coordColNum], roundFactor);
                            }
                        }
                    }
                }
                #endregion

                #region Form print object
                // Get number of rows
                (int totalRows, _) = TwoDArrayFunctions.GetSizeOfConcat(allCoords);
                object[,] printObject = new object[totalRows, 5];

                int currentRowNum = 0;
                int objectNum = 0;

                foreach (double[,] coord in allCoords)
                {
                    // Column 1 - object number and type
                    printObject[currentRowNum, 0] = $"{objectNum + 1}.{allTypes[objectNum]}";
                    objectNum += 1;

                    for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                    {
                        // Column 2 - number
                        printObject[currentRowNum, 1] = coordRowNum + 1;

                        // Column 3 to 5 - coordinates
                        printObject[currentRowNum, 2] = coord[coordRowNum, 0];
                        printObject[currentRowNum, 3] = coord[coordRowNum, 1];
                        if (coord.GetLength(1) > 2)
                        {
                            printObject[currentRowNum, 4] = coord[coordRowNum, 2];
                        }
                        currentRowNum += 1;
                    }
                }

                ReplaceNaN(ref printObject);
                #endregion

                #region Write to Excel
                editor.WriteMessage($"Fido: Begin write to excel." + Environment.NewLine);
                WriteObjectToExcelRange(excelApp, null, 0, 0, true, printObject);
                editor.WriteMessage($"Fido: Coordinates written to excel. Operation completed." + Environment.NewLine);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void getMidPointButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                CheckIfExcelIsAttached();
                (double roundFactor, double distConvFactor, double areaConvFactor) = GetConversionFactors();
                #endregion

                #region Get Coordinates in WCS
                List<double[,]> allCoordsList = new List<double[,]>();
                List<string> allTypes = new List<string>();

                (Document acDoc, Editor editor, Database acDb) = GetAcadDoc();
                using (acDoc.LockDocument())
                {
                    try
                    {
                        using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                        {
                            ObjectId[] selectedObjIds = null;
                            try { selectedObjIds = GetAcadSelectedObjectID(acTrans, editor, skipLockCheck.Checked); }
                            catch (Exception ex)
                            {
                                editor.WriteMessage("Fido: Operation terminated. " + ex.Message + Environment.NewLine); return;
                            }

                            int skippedInvalidCount = 0;
                            foreach (ObjectId objId in selectedObjIds)
                            {
                                dynamic dynEnt = acTrans.GetObject(objId, OpenMode.ForRead) as Entity;
                                double[,] coord = GetCoordinates(dynEnt);

                                if (skipInvalidCheck.Checked && double.IsNaN(coord[0, 0]))
                                {
                                    skippedInvalidCount++;
                                    continue;
                                }

                                allCoordsList.Add(coord);
                                allTypes.Add(dynEnt.AcadObject.ObjectName);
                            }
                            editor.WriteMessage($"Fido: {skippedInvalidCount} invalid object(s)." + Environment.NewLine);
                        }
                    }
                    catch (Exception ex) { throw new Exception("Error in autocad transaction" + ex.Message); }
                }
                #endregion

                #region Transform in UCS
                if (allCoordsList.Count == 0) { throw new Exception("Nothing found to print"); }
                if (translateByUcsCheck.Checked) { TransformListOfCoordsByCs(ref allCoordsList, editor); }
                #endregion

                #region Convert To Mid Points
                List<double[,]> midPointsList = new List<double[,]>();

                foreach (double[,] coords in allCoordsList)
                {
                    #region Handle No Mid Points
                    if (coords.GetLength(0) == 0)
                    {
                        double[,] emptyMidPoint = new double[1, coords.GetLength(1)];
                        for (int colNum = 0; colNum < coords.GetLength(1); colNum++)
                        {
                            emptyMidPoint[0, colNum] = double.NaN;
                        }
                        midPointsList.Add(emptyMidPoint);
                        continue;
                    }
                    #endregion


                    double[,] midPoints = new double[coords.GetLength(0) - 1, coords.GetLength(1)];
                    for (int rowNum = 0; rowNum < coords.GetLength(0) - 1; rowNum++)
                    {
                        double[] point1 = TwoDArrayFunctions.GetSingleRow(coords, rowNum);
                        double[] point2 = TwoDArrayFunctions.GetSingleRow(coords, rowNum + 1);
                        double[] midPoint = findMidPoint(point1, point2);

                        TwoDArrayFunctions.WriteSingleRow(ref midPoints, midPoint, rowNum);
                    }
                    midPointsList.Add(midPoints);
                }
                #endregion

                #region Convert Distances
                if (!double.IsNaN(distConvFactor))
                {
                    foreach (double[,] coord in midPointsList)
                    {
                        for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                        {
                            for (int coordColNum = 0; coordColNum < coord.GetLength(1); coordColNum++)
                            {
                                coord[coordRowNum, coordColNum] = coord[coordRowNum, coordColNum] * distConvFactor;
                            }
                        }
                    }
                }
                #endregion

                #region Round Values
                if (!double.IsNaN(roundFactor))
                {
                    foreach (double[,] coord in midPointsList)
                    {
                        for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                        {
                            for (int coordColNum = 0; coordColNum < coord.GetLength(1); coordColNum++)
                            {
                                coord[coordRowNum, coordColNum] = RoundDouble(coord[coordRowNum, coordColNum], roundFactor);
                            }
                        }
                    }
                }
                #endregion

                #region Form print object
                // Get number of rows
                (int totalRows, _) = TwoDArrayFunctions.GetSizeOfConcat(allCoordsList);
                object[,] printObject = new object[totalRows, 5];

                int currentRowNum = 0;
                int objectNum = 0;

                foreach (double[,] coord in midPointsList)
                {
                    // Column 1 - object number and type
                    printObject[currentRowNum, 0] = $"{objectNum + 1}.{allTypes[objectNum]}";
                    objectNum += 1;

                    for (int coordRowNum = 0; coordRowNum < coord.GetLength(0); coordRowNum++)
                    {
                        // Column 2 - number
                        printObject[currentRowNum, 1] = coordRowNum + 1;

                        // Column 3 to 5 - coordinates
                        printObject[currentRowNum, 2] = coord[coordRowNum, 0];
                        printObject[currentRowNum, 3] = coord[coordRowNum, 1];
                        if (coord.GetLength(1) > 2)
                        {
                            printObject[currentRowNum, 4] = coord[coordRowNum, 2];
                        }
                        currentRowNum += 1;
                    }
                }

                ReplaceNaN(ref printObject);
                #endregion

                #region Write to Excel
                editor.WriteMessage($"Fido: Begin write to excel." + Environment.NewLine);
                WriteObjectToExcelRange(excelApp, null, 0, 0, true, printObject);
                editor.WriteMessage($"Fido: Coordinates written to excel. Operation completed." + Environment.NewLine);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region ACAD Shared Outdated
        //private double[,] GetCoordinates(dynamic dynEnt)
        //{
        //    double[,] formattedObject;
        //    #region Get Coordinates
        //    double[] coord;
        //    bool is3D = true;
        //    switch (dynEnt)
        //    {
        //        case Autodesk.AutoCAD.DatabaseServices.Line line:
        //            // handle Line
        //            coord = new double[6];
        //            coord[0] = line.StartPoint.X;
        //            coord[1] = line.StartPoint.Y;
        //            coord[2] = line.StartPoint.Z;
        //            coord[3] = line.EndPoint.X;
        //            coord[4] = line.EndPoint.Y;
        //            coord[5] = line.EndPoint.Z;
        //            break;
        //        case Polyline poly:
        //            is3D = false;
        //            coord = dynEnt.AcadObject.Coordinates;
        //            break;
        //        case Polyline2d poly2d:
        //            is3D = false;
        //            coord = dynEnt.AcadObject.Coordinates;
        //            break;
        //        case Polyline3d poly3d:
        //            coord = dynEnt.AcadObject.Coordinates;
        //            break;
        //        case Circle circle:
        //            coord = dynEnt.AcadObject.Center;
        //            break;
        //        default:
        //            is3D = false;
        //            coord = new double[] { double.NaN, double.NaN };
        //            break;
        //            //throw new Exception($"AutoCAD object type {dynEnt.AcadObject.ObjectName} does not support this function");
        //    }
        //    #endregion

        //    #region Format Coordinates
        //    if (is3D)
        //    {
        //        formattedObject = new double[coord.Length / 3, 3];
        //        for (int rowNum = 0; rowNum < formattedObject.GetLength(0); rowNum++)
        //        {
        //            formattedObject[rowNum, 0] = coord[rowNum * 3];
        //            formattedObject[rowNum, 1] = coord[rowNum * 3 + 1];
        //            formattedObject[rowNum, 2] = coord[rowNum * 3 + 2];
        //        }
        //    }
        //    else
        //    {
        //        formattedObject = new double[coord.Length / 2, 2];
        //        for (int rowNum = 0; rowNum < formattedObject.GetLength(0); rowNum++)
        //        {
        //            formattedObject[rowNum, 0] = coord[rowNum * 2];
        //            formattedObject[rowNum, 1] = coord[rowNum * 2 + 1];
        //        }
        //    }
        //    #endregion

        //    return formattedObject;
        //}
        //private SelectionSet GetAcadSelectionSet(Editor editor)
        //{
        //    PromptSelectionResult selectionResult = editor.SelectImplied();
        //    if (selectionResult.Status != PromptStatus.OK || selectionResult.Value.Count == 0) { selectionResult = editor.GetSelection(); }

        //    if (selectionResult.Status == PromptStatus.OK && selectionResult.Value.Count > 0)
        //    {
        //        SelectionSet selectedObjs = selectionResult.Value;
        //        return selectedObjs;
        //    }
        //    else { throw new Exception("No object selected"); }
        //}

        //private ObjectId[] GetAcadSelectedObjectID(Transaction acTrans, Editor editor, bool skipLockedObj = false)
        //{
        //    SelectionSet selectedObjs = GetAcadSelectionSet(editor);
            
        //    List<ObjectId> listofObjectId = new List<ObjectId>();
        //    int skippedObj = 0;
        //    foreach (SelectedObject selObj in selectedObjs)
        //    {
        //        dynamic dynEnt = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
        //        if (dynEnt == null) { continue; }

        //        if (skipLockedObj)
        //        {
        //            LayerTableRecord ltr = (LayerTableRecord)acTrans.GetObject(dynEnt.LayerId, OpenMode.ForRead);
        //            if (ltr.IsLocked) { skippedObj += 1; continue; }
        //        }
        //        listofObjectId.Add(selObj.ObjectId);
        //    }
            
        //    if (skippedObj > 0)
        //    {
        //        editor.WriteMessage($"Fido: {skippedObj} object(s) on locked layer(s) skipped." + Environment.NewLine);
        //    }

        //    return listofObjectId.ToArray();
        //}

        //private double[] TranformPointByCs(double[] ogCoord, Matrix3d coordinateSys)
        //{
        //    double[] newCoord = new double[3];
        //    Point3d ogPoint;
        //    if (ogCoord.Length == 3)
        //    {
        //        ogPoint = new Point3d(ogCoord[0], ogCoord[1], ogCoord[2]);
        //    }
        //    else if (ogCoord.Length == 2)
        //    {
        //        ogPoint = new Point3d(ogCoord[0], ogCoord[1], 0);
        //    }
        //    else
        //    {
        //        throw new Exception($"Unexpected number of values for coordinate provided. Provided: {ogCoord.Length}, expected: 2 or 3");
        //    }

        //    Point3d ucsPoint = ogPoint.TransformBy(coordinateSys);
        //    newCoord[0] = ucsPoint.X;
        //    newCoord[1] = ucsPoint.Y;
        //    newCoord[2] = ucsPoint.Z;
        //    return newCoord;
        //}

        //private double[] TranformPointByCs(double[] ogCoord, Editor editor, bool wcsToUcs = true)
        //{
        //    Matrix3d transfMatrix;
        //    if (wcsToUcs)
        //    {
        //        transfMatrix = editor.CurrentUserCoordinateSystem.Inverse();
        //    }
        //    else
        //    {
        //        transfMatrix = editor.CurrentUserCoordinateSystem;
        //    }
        //    double[] newCoord = TranformPointByCs(ogCoord, transfMatrix);
        //    return newCoord;
        //}

        //private void TransformListOfCoordsByCs(ref List<double[,]> ogCoords, Editor editor, bool wcsToUcs = true)
        //{
        //    List<double[,]> newCoords = new List<double[,]>();

        //    Matrix3d transfMatrix;
        //    if (wcsToUcs)
        //    {
        //        transfMatrix = editor.CurrentUserCoordinateSystem.Inverse();
        //    }
        //    else
        //    {
        //        transfMatrix = editor.CurrentUserCoordinateSystem;
        //    }
                
        //    foreach (double[,] coords in ogCoords)
        //    {
        //        double[,] newCoord = new double[coords.GetLength(0), coords.GetLength(1)];
        //        for (int rowNum = 0; rowNum < coords.GetLength(0); rowNum++)
        //        {
        //            Point3d oldPoint;
        //            if (coords.GetLength(1) == 3)
        //            {
        //                oldPoint = new Point3d(coords[rowNum, 0], coords[rowNum, 1], coords[rowNum, 2]);
        //            }
        //            else
        //            {
        //                oldPoint = new Point3d(coords[rowNum, 0], coords[rowNum, 1], 0);
        //            }

        //            Point3d ucsPoint = oldPoint.TransformBy(transfMatrix);
        //            newCoord[rowNum, 0] = ucsPoint.X;
        //            newCoord[rowNum, 1] = ucsPoint.Y;
        //            if (coords.GetLength(1) == 3)
        //            {
        //                newCoord[rowNum, 2] = ucsPoint.Z;
        //            }
        //        }
        //        newCoords.Add(newCoord);
        //    }
        //    ogCoords = newCoords;
        //}

        //private double[] findMidPoint(double[] point1, double[] point2)
        //{
        //    if (point1.Length != point2.Length) { throw new Exception("Points must have the same number of coordinates"); }
        //    double[] midPoint = new double[point1.Length];
        //    for (int i = 0; i < point1.Length; i++)
        //    {
        //        midPoint[i] = (point1[i] + point2[i]) / 2;
        //    }
        //    return midPoint;
        //}
        #endregion


    }
    #region Property Classes
    public class PrintProperties
    {
        // Defines what properties need to be printed
        public Dictionary<string, ListItem> properties = new Dictionary<string, ListItem>();
        private string[] allPropertyNames = new string[] { "Type", "Length", "Area" };
        public PrintProperties(bool initialStatus = true)
        {
            int selectedOrderNum = 1;
            for (int initialOrderNum = 0; initialOrderNum < allPropertyNames.Length; initialOrderNum++)
            {
                if (initialStatus)
                {
                    properties.Add(allPropertyNames[initialOrderNum], new AcadPropertyType(allPropertyNames[initialOrderNum], initialOrderNum, selectedOrderNum));
                    selectedOrderNum++;
                }
                else
                {
                    properties.Add(allPropertyNames[initialOrderNum], new AcadPropertyType(allPropertyNames[initialOrderNum], initialOrderNum, false));
                }

            }
        }
        public void InitPropertyArrays(int arraySize)
        {
            foreach (string propertyName in allPropertyNames)
            {
                AcadPropertyType property = (AcadPropertyType)properties[propertyName];
                if (property.isActive)
                {
                    property.contentArray = new object[arraySize];
                }
            }
        }

        #region Print
        public void PrintToRange(Excel.Application excelApp, Range printRange)
        {
            List<AcadPropertyType> propertiesToPrint = new List<AcadPropertyType>();
            // Loop through to find number of acive
            foreach (string propertyName in properties.Keys)
            {
                AcadPropertyType property = (AcadPropertyType)properties[propertyName];
                if (property.isActive) { propertiesToPrint.Add(property); }
            }

            // Sort
            AcadPropertyType[] propertiesToPrintArray = propertiesToPrint.OrderBy(property => property.selectedOrderNum).ToArray();

            // Create print object
            object[,] printObject = new object[propertiesToPrint[0].contentArray.Length, propertiesToPrint.Count];
            for (int rowNum = 0; rowNum < propertiesToPrint[0].contentArray.Length; rowNum++)
            {
                for (int colNum = 0; colNum < propertiesToPrintArray.Length; colNum++)
                {
                    AcadPropertyType propertyType = propertiesToPrintArray[colNum];
                    var contentArray = propertyType.contentArray;
                    var contentToPrint = contentArray[rowNum];
                    printObject[rowNum, colNum] = contentToPrint;
                }
            }

            // Print to Excel
            WriteObjectToExcelRange(excelApp, printRange, 0, 0, true, printObject);
        }
        #endregion

        #region Convert and round Values
        double roundFactor;
        double distConvFactor;
        double areaConvFactor;
        public void SetConversionFactors(double roundFactor, double distConvFactor, double areaConvFactor)
        {
            this.roundFactor = roundFactor;
            this.distConvFactor = distConvFactor;
            this.areaConvFactor = areaConvFactor;
        }
        public void ConvertAndRoundAll()
        {
            foreach (string propertyName in properties.Keys)
            {
                AcadPropertyType property = (AcadPropertyType)properties[propertyName];
                if (property.isActive) { 
                    property.ConvertAndRoundContent(roundFactor, distConvFactor, areaConvFactor);
                }
            }
        }
        #endregion
    }

    public class AcadPropertyType : ListItem
    {
        // Represents a Acad property type
        // Each array contains the values of that property type for all objects
        public object[] contentArray { get; set; }
        public AcadPropertyType(string name, int initialOrderNum, bool status = true) : base(name, initialOrderNum, status) { }
        public AcadPropertyType(string name, int initialOrderNum, int selectedOrderNum) : base(name, initialOrderNum, selectedOrderNum) { }

        public void ConvertAndRoundContent(double roundFactor, double distConvFactor, double areaConvFactor)
        {
            if (name == "Type") { return; }
            
            #region Convert Values
            if (name == "Length") 
            {
                if (!double.IsNaN(distConvFactor))
                {
                    for (int i = 0; i < contentArray.Length; i++)
                    {
                        if (contentArray[i] is double val)
                        {
                            if (!double.IsNaN(val))
                                val *= distConvFactor;

                            contentArray[i] = val;
                        }
                    }
                }
            }
            else if (name == "Area")
            {
                if (!double.IsNaN(areaConvFactor)) 
                {
                    for (int i = 0; i < contentArray.Length; i++)
                    {
                        if (contentArray[i] is double val)
                        {
                            if (!double.IsNaN(val))
                                val *= areaConvFactor;

                            contentArray[i] = val;
                        }
                    }
                }
            }
            #endregion

            #region Round Values
            if (!double.IsNaN(roundFactor))
            {
                for (int i = 0; i < contentArray.Length; i++)
                {
                    contentArray[i] = RoundDouble((double)contentArray[i], roundFactor);
                }
            }
            #endregion
        }
    }

    public class AcadObject
    {
        // Represents an AutoCAD object type

        #region Init
        public string name { get; set; }
        public int order { get; set; }
        public ObjectId objectId { get; set; }
        public ObjectType objectType { get; set; }
        public enum ObjectType
        {
            Line,
            Polyline,
            Polyline2d,
            Polyline3d,
            Circle
        }

        public AcadObject(string name, int order, ObjectId objectId)
        {
            this.name = name;
            this.order = order;
            this.objectId = objectId;
        }
        #endregion



        #region Coordinates
        double[,] coord;
        public double[,] GetObjectCoordinates(Transaction acTrans)
        {
            dynamic dynEnt = acTrans.GetObject(objectId, OpenMode.ForRead) as Entity;
            string type = dynEnt.AcadObject.ObjectName;
            SetObjectType(type);
            coord = GetCoordinates(dynEnt);
            return coord;
        }

        public void SetObjectType(string typeString)
        {
            switch (typeString)
            {
                case "Line":
                    this.objectType = ObjectType.Line;
                    break;
                case "Polyline":
                    this.objectType = ObjectType.Polyline;
                    break;
                case "Polyline2d":
                    this.objectType = ObjectType.Polyline2d;
                    break;
                case "Polyline3d":
                    this.objectType = ObjectType.Polyline3d;
                    break;
                case "Circle":
                    this.objectType = ObjectType.Circle;
                    break;
                default:
                    throw new Exception($"Unknown object type: {typeString}");
            }
        }
        #endregion


        public void test()
        {
            //dynamic dynEnt = acTrans.GetObject(objId, OpenMode.ForRead) as Entity;
            //double[,] coord = GetCoordinates(dynEnt);

            //if (skipInvalidCheck.Checked && double.IsNaN(coord[0, 0]))
            //{
            //    skippedInvalidCount++;
            //    continue;
            //}

            //allCoords.Add(coord);
            //allTypes.Add(dynEnt.AcadObject.ObjectName);
        }

        private double[,] GetCoordinates(dynamic dynEnt)
        {
            double[,] formattedObject;
            #region Get Coordinates
            double[] coord;
            bool is3D = true;
            switch (dynEnt)
            {
                case Autodesk.AutoCAD.DatabaseServices.Line line:
                    // handle Line
                    coord = new double[6];
                    coord[0] = line.StartPoint.X;
                    coord[1] = line.StartPoint.Y;
                    coord[2] = line.StartPoint.Z;
                    coord[3] = line.EndPoint.X;
                    coord[4] = line.EndPoint.Y;
                    coord[5] = line.EndPoint.Z;
                    break;
                case Polyline poly:
                    is3D = false;
                    coord = dynEnt.AcadObject.Coordinates;
                    break;
                case Polyline2d poly2d:
                    is3D = false;
                    coord = dynEnt.AcadObject.Coordinates;
                    break;
                case Polyline3d poly3d:
                    coord = dynEnt.AcadObject.Coordinates;
                    break;
                case Circle circle:
                    coord = dynEnt.AcadObject.Center;
                    break;
                default:
                    is3D = false;
                    coord = new double[] { double.NaN, double.NaN };
                    break;
                    //throw new Exception($"AutoCAD object type {dynEnt.AcadObject.ObjectName} does not support this function");
            }
            #endregion

            #region Format Coordinates
            if (is3D)
            {
                formattedObject = new double[coord.Length / 3, 3];
                for (int rowNum = 0; rowNum < formattedObject.GetLength(0); rowNum++)
                {
                    formattedObject[rowNum, 0] = coord[rowNum * 3];
                    formattedObject[rowNum, 1] = coord[rowNum * 3 + 1];
                    formattedObject[rowNum, 2] = coord[rowNum * 3 + 2];
                }
            }
            else
            {
                formattedObject = new double[coord.Length / 2, 2];
                for (int rowNum = 0; rowNum < formattedObject.GetLength(0); rowNum++)
                {
                    formattedObject[rowNum, 0] = coord[rowNum * 2];
                    formattedObject[rowNum, 1] = coord[rowNum * 2 + 1];
                }
            }
            #endregion

            return formattedObject;
        }
    }
    #endregion
}

namespace FidoAutoCad
{
    #region Acad Shared Functions
    public static class AcadSharedFunctions
    {
        public static double[,] GetCoordinates(dynamic dynEnt)
        {
            double[,] formattedObject;
            #region Get Coordinates
            double[] coord;
            bool is3D = true;
            switch (dynEnt)
            {
                case Autodesk.AutoCAD.DatabaseServices.Line line:
                    // handle Line
                    coord = new double[6];
                    coord[0] = line.StartPoint.X;
                    coord[1] = line.StartPoint.Y;
                    coord[2] = line.StartPoint.Z;
                    coord[3] = line.EndPoint.X;
                    coord[4] = line.EndPoint.Y;
                    coord[5] = line.EndPoint.Z;
                    break;
                case Polyline poly:
                    is3D = false;
                    coord = dynEnt.AcadObject.Coordinates;
                    break;
                case Polyline2d poly2d:
                    is3D = false;
                    coord = dynEnt.AcadObject.Coordinates;
                    break;
                case Polyline3d poly3d:
                    coord = dynEnt.AcadObject.Coordinates;
                    break;
                case Circle circle:
                    coord = dynEnt.AcadObject.Center;
                    break;
                default:
                    is3D = false;
                    coord = new double[] { double.NaN, double.NaN };
                    break;
                    //throw new Exception($"AutoCAD object type {dynEnt.AcadObject.ObjectName} does not support this function");
            }
            #endregion

            #region Format Coordinates
            if (is3D)
            {
                formattedObject = new double[coord.Length / 3, 3];
                for (int rowNum = 0; rowNum < formattedObject.GetLength(0); rowNum++)
                {
                    formattedObject[rowNum, 0] = coord[rowNum * 3];
                    formattedObject[rowNum, 1] = coord[rowNum * 3 + 1];
                    formattedObject[rowNum, 2] = coord[rowNum * 3 + 2];
                }
            }
            else
            {
                formattedObject = new double[coord.Length / 2, 2];
                for (int rowNum = 0; rowNum < formattedObject.GetLength(0); rowNum++)
                {
                    formattedObject[rowNum, 0] = coord[rowNum * 2];
                    formattedObject[rowNum, 1] = coord[rowNum * 2 + 1];
                }
            }
            #endregion

            return formattedObject;
        }
        public static SelectionSet GetAcadSelectionSet(Editor editor)
        {
            PromptSelectionResult selectionResult = editor.SelectImplied();
            if (selectionResult.Status != PromptStatus.OK || selectionResult.Value.Count == 0) { selectionResult = editor.GetSelection(); }

            if (selectionResult.Status == PromptStatus.OK && selectionResult.Value.Count > 0)
            {
                SelectionSet selectedObjs = selectionResult.Value;
                return selectedObjs;
            }
            else { throw new Exception("No object selected"); }
        }

        public static ObjectId[] GetAcadSelectedObjectID(Transaction acTrans, Editor editor, bool skipLockedObj = false)
        {
            SelectionSet selectedObjs = GetAcadSelectionSet(editor);

            List<ObjectId> listofObjectId = new List<ObjectId>();
            int skippedObj = 0;
            foreach (SelectedObject selObj in selectedObjs)
            {
                dynamic dynEnt = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                if (dynEnt == null) { continue; }

                if (skipLockedObj)
                {
                    LayerTableRecord ltr = (LayerTableRecord)acTrans.GetObject(dynEnt.LayerId, OpenMode.ForRead);
                    if (ltr.IsLocked) { skippedObj += 1; continue; }
                }
                listofObjectId.Add(selObj.ObjectId);
            }

            if (skippedObj > 0)
            {
                editor.WriteMessage($"Fido: {skippedObj} object(s) on locked layer(s) skipped." + Environment.NewLine);
            }

            return listofObjectId.ToArray();
        }

        public static double[] TranformPointByCs(double[] ogCoord, Matrix3d coordinateSys)
        {
            double[] newCoord = new double[3];
            Point3d ogPoint;
            if (ogCoord.Length == 3)
            {
                ogPoint = new Point3d(ogCoord[0], ogCoord[1], ogCoord[2]);
            }
            else if (ogCoord.Length == 2)
            {
                ogPoint = new Point3d(ogCoord[0], ogCoord[1], 0);
            }
            else
            {
                throw new Exception($"Unexpected number of values for coordinate provided. Provided: {ogCoord.Length}, expected: 2 or 3");
            }

            Point3d ucsPoint = ogPoint.TransformBy(coordinateSys);
            newCoord[0] = ucsPoint.X;
            newCoord[1] = ucsPoint.Y;
            newCoord[2] = ucsPoint.Z;
            return newCoord;
        }

        public static double[] TranformPointByCs(double[] ogCoord, Editor editor, bool wcsToUcs = true)
        {
            Matrix3d transfMatrix;
            if (wcsToUcs)
            {
                transfMatrix = editor.CurrentUserCoordinateSystem.Inverse();
            }
            else
            {
                transfMatrix = editor.CurrentUserCoordinateSystem;
            }
            double[] newCoord = TranformPointByCs(ogCoord, transfMatrix);
            return newCoord;
        }

        public static void TransformListOfCoordsByCs(ref List<double[,]> ogCoords, Editor editor, bool wcsToUcs = true)
        {
            List<double[,]> newCoords = new List<double[,]>();

            Matrix3d transfMatrix;
            if (wcsToUcs)
            {
                transfMatrix = editor.CurrentUserCoordinateSystem.Inverse();
            }
            else
            {
                transfMatrix = editor.CurrentUserCoordinateSystem;
            }

            foreach (double[,] coords in ogCoords)
            {
                double[,] newCoord = new double[coords.GetLength(0), coords.GetLength(1)];
                for (int rowNum = 0; rowNum < coords.GetLength(0); rowNum++)
                {
                    Point3d oldPoint;
                    if (coords.GetLength(1) == 3)
                    {
                        oldPoint = new Point3d(coords[rowNum, 0], coords[rowNum, 1], coords[rowNum, 2]);
                    }
                    else
                    {
                        oldPoint = new Point3d(coords[rowNum, 0], coords[rowNum, 1], 0);
                    }

                    Point3d ucsPoint = oldPoint.TransformBy(transfMatrix);
                    newCoord[rowNum, 0] = ucsPoint.X;
                    newCoord[rowNum, 1] = ucsPoint.Y;
                    if (coords.GetLength(1) == 3)
                    {
                        newCoord[rowNum, 2] = ucsPoint.Z;
                    }
                }
                newCoords.Add(newCoord);
            }
            ogCoords = newCoords;
        }

        public static double[] findMidPoint(double[] point1, double[] point2)
        {
            if (point1.Length != point2.Length) { throw new Exception("Points must have the same number of coordinates"); }
            double[] midPoint = new double[point1.Length];
            for (int i = 0; i < point1.Length; i++)
            {
                midPoint[i] = (point1[i] + point2[i]) / 2;
            }
            return midPoint;
        }
    }
    #endregion
}
