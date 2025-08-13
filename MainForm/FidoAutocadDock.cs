using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using FidoAutoCad.SharedForms;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using static FidoAutoCad.AcadSharedFunctions;
using static FidoAutoCad.CommonUtilities;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;

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
                    //(double roundFactor, double distConvFactor, double areaConvFactor, RoundingMode roundMode) = GetConversionFactors();
                    //printProperties.SetConversionFactors(roundFactor, distConvFactor, areaConvFactor);
                    printProperties.GetConversionFactorFromForm(this);
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

        public (double roundFactor, double distConvFactor, double areaConvFactor, RoundMode roundMode) GetConversionFactors()
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

            #region Round Mode
            string roundModeString = roundTypeComboBox.Text.ToLower();
            RoundMode roundMode;
            switch (roundModeString)
            {
                case "nearest":
                    roundMode = RoundMode.Nearest;
                    break;
                case "ceiling":
                    roundMode = RoundMode.Ceiling;
                        break;
                case "floor":
                    roundMode = RoundMode.Floor;
                    break;
                case "ceiling (abs)":
                    roundMode = RoundMode.CeilingAbs;
                    break;
                case "floor (abs)":
                    roundMode = RoundMode.FloorAbs;
                    break;
                default:
                    throw new Exception($"Round mode {roundModeString} is undefined");
            }
            #endregion
            return (roundFactor, distConvFactor, areaConvFactor, roundMode);
        }
        #endregion

        #region Get Coordinates or Mid Points
        private void getMidPointButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                CheckIfExcelIsAttached();
                //(double roundFactor, double distConvFactor, double areaConvFactor, RoundingMode roundMode) = GetConversionFactors();
                ConversionFactors cF = new ConversionFactors(this);
                #endregion

                #region Get Coordinates in WCS
                List<AcadObject> acadObjects = new List<AcadObject>();

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

                            #region Refractored Version
                            int objCount = 0;
                            int skippedInvalidCount = 0;
                            foreach (ObjectId objId in selectedObjIds)
                            {
                                AcadObject acadObj = new AcadObject(objId);
                                double[,] coord = acadObj.GetObjectCoordinatesFromAcad(acTrans);

                                if (skipInvalidCheck.Checked && double.IsNaN(coord[0, 0]))
                                {
                                    skippedInvalidCount++;
                                    continue;
                                }
                                acadObj.order = objCount;
                                objCount += 1;
                                acadObjects.Add(acadObj);
                            }
                            #endregion
                            editor.WriteMessage($"Fido: {skippedInvalidCount} invalid object(s)." + Environment.NewLine);
                        }
                    }
                    catch (Exception ex) { throw new Exception("Error in autocad transaction" + ex.Message); }
                }
                #endregion

                if (acadObjects.Count == 0) { throw new Exception("Nothing found to print"); }
                int totalPrintRows = 0;

                foreach (AcadObject acadObj in acadObjects)
                {
                    #region Transform in UCS
                    if (translateByUcsCheck.Checked)
                    {
                        TransformCoordByCs(ref acadObj.coords, editor);
                    }
                    #endregion

                    #region Find Mid Points
                    acadObj.CalculateMidPoints();
                    #endregion

                    //#region Convert Distances
                    //if (!double.IsNaN(cF.distConvFactor))
                    //{
                    //    acadObj.ConvertMidPoints(cF);
                    //}
                    //#endregion

                    //#region Round Values
                    //if (!double.IsNaN(cF.roundFactor))
                    //{
                    //    acadObj.RoundMidPoints(cF);
                    //}
                    //#endregion

                    acadObj.ConvertAndRoundMidPoints(cF);

                    #region Get Number of Rows
                    totalPrintRows += acadObj.midPoints.GetLength(0);
                    #endregion
                }

                #region Form Print Object
                int objectNum = 0;
                List<object[,]> printArrays = new List<object[,]>();
                foreach (AcadObject acadObj in acadObjects)
                {
                    // Column 1 - object number and type
                    object[,] printArray = acadObj.CreateMidPointPrintObj();
                    printArrays.Add(printArray);
                    objectNum += 1;
                }
                object[,] printObject = TwoDArrayFunctions.ConcatArrays(printArrays);
                #endregion

                #region Write to Excel
                editor.WriteMessage($"Fido: Begin write to excel." + Environment.NewLine);
                WriteObjectToExcelRange(excelApp, null, 0, 0, true, printObject);
                editor.WriteMessage($"Fido: Coordinates written to excel. Operation completed." + Environment.NewLine);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getCoordinatesButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                CheckIfExcelIsAttached();
                //(double roundFactor, double distConvFactor, double areaConvFactor, RoundingMode roundMode) = GetConversionFactors();
                ConversionFactors cF = new ConversionFactors(this);
                #endregion

                #region Get Coordinates in WCS
                List<AcadObject> acadObjects = new List<AcadObject>();

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

                            #region Refractored Version
                            int objCount = 0;
                            int skippedInvalidCount = 0;
                            foreach (ObjectId objId in selectedObjIds)
                            {
                                AcadObject acadObj = new AcadObject(objId);
                                double[,] coord = acadObj.GetObjectCoordinatesFromAcad(acTrans);

                                if (skipInvalidCheck.Checked && double.IsNaN(coord[0, 0]))
                                {
                                    skippedInvalidCount++;
                                    continue;
                                }
                                acadObj.order = objCount;
                                objCount += 1;
                                acadObjects.Add(acadObj);
                            }
                            #endregion
                            editor.WriteMessage($"Fido: {skippedInvalidCount} invalid object(s)." + Environment.NewLine);
                        }
                    }
                    catch (Exception ex) { throw new Exception("Error in autocad transaction" + ex.Message); }
                }
                #endregion

                if (acadObjects.Count == 0) { throw new Exception("Nothing found to print"); }
                int totalPrintRows = 0;

                foreach (AcadObject acadObj in acadObjects)
                {
                    #region Transform in UCS
                    if (translateByUcsCheck.Checked)
                    {
                        TransformCoordByCs(ref acadObj.coords, editor);
                    }
                    #endregion

                    //#region Convert Distances
                    //if (!double.IsNaN(cF.distConvFactor))
                    //{
                    //    acadObj.ConvertCoordinates(cF);
                    //}
                    //#endregion

                    //#region Round Values
                    //if (!double.IsNaN(cF.roundFactor))
                    //{
                    //    acadObj.RoundCoordinates(cF);
                    //}
                    //#endregion

                    acadObj.ConvertAndRoundCoordinates(cF);

                    #region Get Number of Rows
                    totalPrintRows += acadObj.coords.GetLength(0);
                    #endregion
                }

                #region Form Print Object
                int objectNum = 0;
                List<object[,]> printArrays = new List<object[,]>();
                foreach (AcadObject acadObj in acadObjects)
                {
                    // Column 1 - object number and type
                    object[,] printArray = acadObj.CreateCoordPrinttObj();
                    printArrays.Add(printArray);
                    objectNum += 1;
                }
                object[,] printObject = TwoDArrayFunctions.ConcatArrays(printArrays);
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
        //double roundFactor;
        //double distConvFactor;
        //double areaConvFactor;
        //public void SetConversionFactors(double roundFactor, double distConvFactor, double areaConvFactor)
        //{
        //    this.roundFactor = roundFactor;
        //    this.distConvFactor = distConvFactor;
        //    this.areaConvFactor = areaConvFactor;
        //}
        ConversionFactors conversionFactors;
        public void SetConversionFactors(double roundFactor, double distConvFactor, double areaConvFactor)
        {
            conversionFactors = new ConversionFactors(roundFactor, distConvFactor, areaConvFactor);
        }

        public void GetConversionFactorFromForm(FidoAutocadDock parentForm)
        {
            conversionFactors = new ConversionFactors(parentForm);
        }
        public void ConvertAndRoundAll()
        {
            foreach (string propertyName in properties.Keys)
            {
                AcadPropertyType property = (AcadPropertyType)properties[propertyName];
                if (property.isActive)
                {
                    property.ConvertAndRoundContent(conversionFactors);
                }
            }
        }
        #endregion
    }

    public class ConversionFactors
    {
        // Represents conversion factors for rounding and unit conversion
        public double roundFactor { get; set; }
        public double distConvFactor { get; set; }
        public double areaConvFactor { get; set; }
        public RoundMode roundMode { get; set; } = RoundMode.Nearest;
        public ConversionFactors(double roundFactor, double distConvFactor, double areaConvFactor)
        {
            this.roundFactor = roundFactor;
            this.distConvFactor = distConvFactor;
            this.areaConvFactor = areaConvFactor;
        }

        public ConversionFactors()
        {
            this.roundFactor = double.NaN;
            this.distConvFactor = double.NaN;
            this.areaConvFactor = double.NaN;
        }
        public ConversionFactors(FidoAutocadDock parentForm)
        {
            (this.roundFactor, this.distConvFactor, this.areaConvFactor, this.roundMode) = parentForm.GetConversionFactors();
        }

        #region Modify Array
        public void ConvertAndRoundArray(ref double[,] array, bool isDist)
        {
            if (isDist)
            {
                if (!(double.IsNaN(distConvFactor)))
                {
                    TwoDArrayFunctions.MultiplyArray(ref array, distConvFactor);
                }
            }
            else // Is Area
            {
                if (!(double.IsNaN(areaConvFactor)))
                {
                    TwoDArrayFunctions.MultiplyArray(ref array, areaConvFactor);
                }
            }

            if (!(double.IsNaN(roundFactor)))
            {
                TwoDArrayFunctions.RoundArray(ref array, roundFactor, roundMode);
            }               
        }

        public void ConvertAndRoundArray(ref double[] array, bool isDist)
        {
            if (isDist)
            {
                if (!(double.IsNaN(distConvFactor)))
                {
                    TwoDArrayFunctions.MultiplyArray(ref array, distConvFactor);
                }
            }
            else // Is Area
            {
                if (!(double.IsNaN(areaConvFactor)))
                {
                    TwoDArrayFunctions.MultiplyArray(ref array, areaConvFactor);
                }
            }

            if (!(double.IsNaN(roundFactor)))
            {
                TwoDArrayFunctions.RoundArray(ref array, roundFactor, roundMode);
            }
        }

        public void ConvertAndRoundArray(ref object[] array, bool isDist)
        {
            if (isDist)
            {
                if (!(double.IsNaN(distConvFactor)))
                {
                    TwoDArrayFunctions.MultiplyArray(ref array, distConvFactor);
                }
            }
            else // Is Area
            {
                if (!(double.IsNaN(areaConvFactor)))
                {
                    TwoDArrayFunctions.MultiplyArray(ref array, areaConvFactor);
                }
            }

            if (!(double.IsNaN(roundFactor)))
            {
                TwoDArrayFunctions.RoundArray(ref array, roundFactor, roundMode);
            }
        }
        #endregion
    }

    public class AcadPropertyType : ListItem
    {
        // Represents a Acad property type
        // Each array contains the values of that property type for all objects
        public object[] contentArray;
        public AcadPropertyType(string name, int initialOrderNum, bool status = true) : base(name, initialOrderNum, status) { }
        public AcadPropertyType(string name, int initialOrderNum, int selectedOrderNum) : base(name, initialOrderNum, selectedOrderNum) { }

        public void ConvertAndRoundContent(ConversionFactors cF)
        {
            if (name == "Type") { return; }
            #region Old
            //#region Convert Values
            //if (name == "Length")
            //{
            //    if (!double.IsNaN(cF.distConvFactor))
            //    {
            //        for (int i = 0; i < contentArray.Length; i++)
            //        {
            //            if (contentArray[i] is double val)
            //            {
            //                if (!double.IsNaN(val))
            //                    val *= cF.distConvFactor;

            //                contentArray[i] = val;
            //            }
            //        }
            //    }
            //}
            //else if (name == "Area")
            //{
            //    if (!double.IsNaN(cF.areaConvFactor))
            //    {
            //        for (int i = 0; i < contentArray.Length; i++)
            //        {
            //            if (contentArray[i] is double val)
            //            {
            //                if (!double.IsNaN(val))
            //                    val *= cF.areaConvFactor;

            //                contentArray[i] = val;
            //            }
            //        }
            //    }
            //}
            //#endregion

            //#region Round Values
            //if (!double.IsNaN(cF.roundFactor))
            //{
            //    for (int i = 0; i < contentArray.Length; i++)
            //    {
            //        if (!(contentArray[i] is double)) { continue; }
            //        contentArray[i] = RoundDouble((double)contentArray[i], cF.roundFactor);
            //    }
            //}
            //#endregion
            #endregion

            if (name == "Length")
            {
                cF.ConvertAndRoundArray(ref contentArray, true);
            }
            else if (name == "Area")
            {
                cF.ConvertAndRoundArray(ref contentArray, false);
            }
    }
    }
    #endregion

    #region Acad Object
    public class AcadObject
    {
        // Represents an AutoCAD object type

        #region Init
        public string name { get; set; }
        public int order { get; set; }
        public ObjectId objectId { get; set; }


        public AcadObject(string name, int order, ObjectId objectId)
        {
            this.name = name;
            this.order = order;
            this.objectId = objectId;
        }

        public AcadObject(ObjectId objectId)
        {
            this.objectId = objectId;
        }
        #endregion

        #region Coordinates and Object Type
        public double[,] coords;
        public bool isClosed;
        public double[,] GetObjectCoordinatesFromAcad(Transaction acTrans)
        {
            dynamic dynEnt = acTrans.GetObject(objectId, OpenMode.ForRead) as Entity;
            type = dynEnt.AcadObject.ObjectName;
            //SetObjectType(type);
            coords = GetCoordinates(dynEnt);
            try
            {
                isClosed = dynEnt.AcadObject.Closed;
            }
            catch { isClosed = false; }

            return coords;
        }
        #endregion

        #region Type
        //public ObjectType objectType { get; set; }
        //public enum ObjectType
        //{
        //    Line,
        //    Polyline,
        //    Polyline2d,
        //    Polyline3d,
        //    Circle,
        //    Others
        //}
        string type;
        //public void SetObjectType(string typeString)
        //{
        //    switch (typeString)
        //    {
        //        // these are wrong and need to be fixed
        //        case "AcDbLine":
        //            this.objectType = ObjectType.Line;
        //            break;
        //        case "AcDbPolyline":
        //            this.objectType = ObjectType.Polyline;
        //            break;
        //        case "AcDbPolyline2d":
        //            this.objectType = ObjectType.Polyline2d;
        //            break;
        //        case "AcDb3dPolyline":
        //            this.objectType = ObjectType.Polyline3d;
        //            break;
        //        case "AcDbCircle":
        //            this.objectType = ObjectType.Circle;
        //            break;
        //        default:
        //            throw new Exception($"Unknown object type: {typeString}");
        //    }
        //}

        public string GetObjectTypeFromAcad(Transaction acTrans)
        {
            dynamic dynEnt = acTrans.GetObject(objectId, OpenMode.ForRead) as Entity;
            type = dynEnt.AcadObject.ObjectName;
            //SetObjectType(type);
            return type;
        }
        #endregion

        #region Mid Point
        public double[,] midPoints;
        public double[,] CalculateMidPoints()
        {
            #region Checks
            if (coords == null) { throw new Exception("No coordinates provided"); }

            if (coords.GetLength(0) < 2)
            {
                midPoints = new double[1, coords.GetLength(1)];
                for (int colNum = 0; colNum < coords.GetLength(1); colNum++)
                {
                    midPoints[0, colNum] = double.NaN;
                }
                return midPoints;
            }
            #endregion

            #region Calculate Size
            if (isClosed) { midPoints = new double[coords.GetLength(0), coords.GetLength(1)]; }
            else { midPoints = new double[coords.GetLength(0) - 1, coords.GetLength(1)]; }
            #endregion

            #region Calculate Mid Points
            for (int rowNum = 0; rowNum < coords.GetLength(0) - 1; rowNum++)
            {
                double[] point1 = TwoDArrayFunctions.GetSingleRow(coords, rowNum);
                double[] point2 = TwoDArrayFunctions.GetSingleRow(coords, rowNum + 1);
                double[] midPoint = findMidPoint(point1, point2);

                TwoDArrayFunctions.WriteSingleRow(ref midPoints, midPoint, rowNum);
            }

            // Add last point if closed
            if (isClosed)
            {
                double[] point1 = TwoDArrayFunctions.GetSingleRow(coords, coords.GetLength(0) - 1);
                double[] point2 = TwoDArrayFunctions.GetSingleRow(coords, 0);
                double[] midPoint = findMidPoint(point1, point2);

                TwoDArrayFunctions.WriteSingleRow(ref midPoints, midPoint, coords.GetLength(0) - 1);
            }
            #endregion
            return midPoints;
        }

        public object[,] CreateMidPointPrintObj()
        {
            return CreatePrintObj(midPoints);
            //if (midPoints == null) { throw new Exception("No mid points calculated"); } 
            //object[,] printObject = new object[midPoints.GetLength(0), 5];

            //printObject[0, 0] = $"{order}.{type}";

            //for (int coordRowNum = 0; coordRowNum < midPoints.GetLength(0); coordRowNum++)
            //{
            //    // Column 2 - numbers
            //    printObject[coordRowNum, 1] = coordRowNum + 1;
            //}

            //// Column 3 to 5 - coordinates
            //TwoDArrayFunctions.WriteArrayIntoArray(ref printObject, midPoints, 0, 2);
            //return printObject;
        }
        public object[,] CreateCoordPrinttObj()
        {
            return CreatePrintObj(coords);
        }
        public object[,] CreatePrintObj(double[,] sourceObject)
        {

            if (sourceObject == null) { throw new Exception("No mid points calculated"); }
            object[,] printObject = new object[sourceObject.GetLength(0), 5];

            printObject[0, 0] = $"{order}.{type}";

            for (int coordRowNum = 0; coordRowNum < sourceObject.GetLength(0); coordRowNum++)
            {
                // Column 2 - numbers
                printObject[coordRowNum, 1] = coordRowNum + 1;
            }

            // Column 3 to 5 - coordinates
            TwoDArrayFunctions.WriteArrayIntoArray(ref printObject, sourceObject, 0, 2);
            ReplaceNaN(ref printObject);
            return printObject;
        }
        #endregion

        #region Rounding and Conversion
        //public void ConvertCoordinates(ConversionFactors cF)
        //    => TwoDArrayFunctions.MultiplyArray(coords, cF.distConvFactor);

        //public void ConvertMidPoints(ConversionFactors cF)
        //    => TwoDArrayFunctions.MultiplyArray(midPoints, cF.distConvFactor);

        //public void RoundCoordinates(ConversionFactors cF)
        //    => TwoDArrayFunctions.RoundArray(coords, cF.roundFactor);

        //public void RoundMidPoints(ConversionFactors cF)
        //    => TwoDArrayFunctions.RoundArray(midPoints, cF.roundFactor);

        public void ConvertAndRoundCoordinates(ConversionFactors cF)
        {
            cF.ConvertAndRoundArray(ref coords, true);
        }

        public void ConvertAndRoundMidPoints(ConversionFactors cF)
        {
            cF.ConvertAndRoundArray(ref midPoints, true);
        }

        #endregion
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

        public static double[] TranformPointByCs(ref double[] ogCoord, Editor editor, bool wcsToUcs = true)
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

        public static void TransformCoordByCs(ref List<double[,]> ogCoords, Editor editor, bool wcsToUcs = true)
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

        public static void TransformCoordByCs(ref double[,] ogCoords, Editor editor, bool wcsToUcs = true)
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

            double[,] newCoords = new double[ogCoords.GetLength(0), ogCoords.GetLength(1)];
            for (int rowNum = 0; rowNum < ogCoords.GetLength(0); rowNum++)
            {
                Point3d oldPoint;
                if (ogCoords.GetLength(1) == 3)
                {
                    oldPoint = new Point3d(ogCoords[rowNum, 0], ogCoords[rowNum, 1], ogCoords[rowNum, 2]);
                }
                else
                {
                    oldPoint = new Point3d(ogCoords[rowNum, 0], ogCoords[rowNum, 1], 0);
                }

                Point3d ucsPoint = oldPoint.TransformBy(transfMatrix);
                newCoords[rowNum, 0] = ucsPoint.X;
                newCoords[rowNum, 1] = ucsPoint.Y;
                if (ogCoords.GetLength(1) == 3)
                {
                    newCoords[rowNum, 2] = ucsPoint.Z;
                }
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
