using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using FidoAutoCad.SharedForms;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDB = Autodesk.AutoCAD.DatabaseServices;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;

using static FidoAutoCad.CommonUtilities;
using Autodesk.AutoCAD.Internal.PropertyInspector;

namespace FidoAutoCad.Forms
{
    public partial class FidoAutocadDock : UserControl
    {
        AutoCADCommands parent;
        public FidoAutocadDock(AutoCADCommands parent)
        {
            InitializeComponent();
            this.parent = parent;
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        }
        Excel.Application excelApp = null;


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

                #region Get the current document and database, and start a transaction
                Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor editor = acDoc.Editor;
                Database acCurDb = acDoc.Database;
                try
                {
                    acDoc = AcadApp.DocumentManager.MdiActiveDocument;
                    editor = acDoc.Editor;
                    acCurDb = acDoc.Database;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to attach active autocad document ${ex.Message}");
                }
                #endregion

                #region Autocad Transaction
                //string[] acadTypes = null;
                //double[] lengths = null;
                //double[] areas = null;

                using (acDoc.LockDocument())
                {
                    // Starts a new transaction with the Transaction Manager
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
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
                                //// Type
                                //if (printProperties.properties["Type"].isActive)
                                //{
                                //    try
                                //    {
                                //        string acadType = dynEnt.AcadObject.EntityName;
                                //        acadTypes[i] = acadType;
                                //    }
                                //    catch { acadTypes[i] = "NA"; }
                                //}

                                //// Length
                                //if (printProperties.properties["Length"].isActive)
                                //{
                                //    try
                                //    {
                                //        double length = dynEnt.Length;
                                //        lengths[i] = length;
                                //    }
                                //    catch { lengths[i] = double.NaN; }
                                //}

                                //// Area
                                //if (printProperties.properties["Area"].isActive)
                                //{
                                //    try
                                //    {
                                //        double area = dynEnt.Area;
                                //        areas[i] = area;
                                //    }
                                //    catch { areas[i] = double.NaN; }
                                //}
                                objectNum++;
                            }
                        }
                        else { return; }
                    }
                }
                #endregion

                #region Write to Excel
                //WriteToExcelRangeAsCol(excelApp, null, 0, 0, true, acadTypes, ConvertDoubleArrayToStringArray(lengths, 0), ConvertDoubleArrayToStringArray(areas, 0));
                // To refractor to convert to object and print, or to make properties just a separate class
                //int colNum = 0;
                //if (printProperties.propertyStatus["Type"])
                //{
                //    WriteToExcelRangeAsCol(excelApp, null, 0, colNum, false, acadTypes);
                //    colNum += 1;
                //}

                //// Length
                //if (printProperties.propertyStatus["Length"])
                //{
                //    WriteToExcelRangeAsCol(excelApp, null, 0, colNum, false, ConvertDoubleArrayToStringArray(lengths, 0));
                //    colNum += 1;
                //}

                //// Area
                //if (printProperties.propertyStatus["Area"])
                //{
                //    WriteToExcelRangeAsCol(excelApp, null, 0, colNum, false, ConvertDoubleArrayToStringArray(areas, 0));
                //    colNum += 1;
                //}
                printProperties.PrintToRange(excelApp, null);
                excelApp.ScreenUpdating = true;
                MessageBox.Show("Completed", "Fido AutoCAD");
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }

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
        #endregion

        #region Test Zone
        private void testButton_Click(object sender, EventArgs e)
        {

        }
        #endregion

    }
    #region Property Classes
    public class PrintProperties
    {
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
            //object[,] printObject = new object[propertiesToPrint.Count, propertiesToPrint[0].contentArray.Length];
            //for (int colNum = 0; colNum < propertiesToPrintArray.Length; colNum++)
            //{
            //    for (int rowNum = 0; rowNum < propertiesToPrint[0].contentArray.Length; rowNum++)
            //    {
            //        var testVar = propertiesToPrint[colNum].contentArray[rowNum];
            //        printObject[colNum, rowNum] = propertiesToPrint[colNum].contentArray[rowNum];
            //    }
            //}
            object[,] printObject = new object[propertiesToPrint[0].contentArray.Length, propertiesToPrint.Count];
            //object[,] printObject = new object[propertiesToPrint.Count, propertiesToPrint[0].contentArray.Length];
            for (int rowNum = 0; rowNum < propertiesToPrint[0].contentArray.Length ; rowNum++)
            {
                for (int colNum = 0; colNum < propertiesToPrintArray.Length; colNum++)
                {
                    AcadPropertyType propertyType = propertiesToPrintArray[colNum];
                    var contentArray = propertyType.contentArray;
                    var contentToPrint = contentArray[rowNum];
                    printObject[rowNum, colNum] = contentToPrint;
                    //printObject[rowNum, colNum] = propertiesToPrint[rowNum].contentArray[colNum];
                }
            }


            // Print to Excel
            WriteObjectToExcelRange(excelApp, printRange, 0, 0, true, printObject);
        }
        #endregion
    }

    public class AcadPropertyType : ListItem
    {
        public object[] contentArray { get; set; }
        public AcadPropertyType(string name, int initialOrderNum, bool status = true) : base(name, initialOrderNum, status) { }
        public AcadPropertyType(string name, int initialOrderNum, int selectedOrderNum) : base(name, initialOrderNum, selectedOrderNum) { }
    }
    #endregion
}
