using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FidoAutoCad.SharedForms
{
    public class AttributeTextBox
    {
        #region Initialisation
        public string attName;
        public System.Windows.Forms.TextBox textBox;
        public string type = "string";
        public string defaultValue;
        public string currentValue = null;
        public AttributeTextBox(string attName, TextBox textBox, bool isBasicValue = false)
        {
            this.attName = attName;
            this.textBox = textBox;

            if (isBasicValue)
            {
                SubscribeToTextBoxEvents();
            }

            RefreshTextBox();
        }

        public AttributeTextBox(string attName, TextBox textBox, string defaultValue, bool isBasicValue = false)
        {
            this.attName = attName;
            this.textBox = textBox;

            if (isBasicValue)
            {
                SubscribeToTextBoxEvents();
            }
            SetDefaultValue(defaultValue);
        }

        protected void SubscribeToTextBoxEvents()
        {
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        protected void textBox_LostFocus(object sender, EventArgs e)
        {
            PerformVerification();
            SetValueFromTextBox();
        }

        protected void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
        }
        #endregion

        #region Additional Verification

        private void PerformVerification()
        {
            if (textBox.Text == "")
            {
                return;
            }

            switch (type)
            {
                case "string":
                    break;
                case "int":
                    CheckIsInt();
                    break;
                case "double":
                    CheckIsDouble();
                    break;
                case "filename":
                    CheckIsFileName();
                    break;
                case "partial filepath":
                    CheckIsPartialFilePath();
                    break;
                default:
                    throw new Exception("Verification type not found");
            }
        }

        private bool CheckIsInt()
        {
            try
            {
                Convert.ToInt32(textBox.Text);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Value '{textBox.Text}' for {attName} cannot be converted to integer\n\n" + ex.Message, "Error");
                RefreshTextBox();
                return false;
            }
        }

        private bool CheckIsDouble()
        {
            bool success = double.TryParse(textBox.Text, out double value);
            if (!success)
            {
                MessageBox.Show($"Value '{textBox.Text}' for {attName} cannot be converted to number (double)", "Error");
                RefreshTextBox();
                return false;
            }
            return true;
        }

        private bool CheckIsFileName()
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidFileNameChars());
            bool success = true;
            string invalidCharsInString = "";
            foreach (char c in textBox.Text)
            {
                if (invalidChars.Contains(c))
                {
                    invalidCharsInString += c + " ";
                    success = false;
                }
            }

            if (!success)
            {
                MessageBox.Show($"'{textBox.Text}' contains invalid characters: \n" + invalidCharsInString, "Error");
                RefreshTextBox();
                return false;
            }
            return true;
        }

        private bool CheckIsPartialFilePath()
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidPathChars());
            bool success = true;
            foreach (char c in textBox.Text)
            {
                if (invalidChars.Contains(c))
                {
                    success = false;
                    break;
                }
            }

            if (!success)
            {
                string invalidString = "";
                foreach (char character in invalidChars)
                {
                    invalidString += character;
                }
                MessageBox.Show($"'{textBox.Text}' contains invalid characters: \n{invalidString}", "Error");
                RefreshTextBox();
                return false;
            }
            return true;
        }

        #endregion

        #region Default Value
        public void SetDefaultValue(string value)
        {
            defaultValue = value;
            RefreshTextBox();
        }
        #endregion

        #region Get and Set Values for Properties
        //public (bool, DocumentProperty) GetValueFromProp()
        //{
        //    // This always gets value from what is saved in the document. 
        //    // First output is boolean stating if property has been set or not 
        //    // Second output returns the DocumentProperty type
        //    DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
        //    foreach (DocumentProperty prop in AllCustProps)
        //    {
        //        if (prop.Name == attName)
        //        {
        //            return (true, prop);
        //        }
        //    }
        //    return (false, null);
        //}

        //public bool SetValue(object AttValue, bool showMsg = false)
        //{
        //    try
        //    {
        //        DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
        //        (bool exist, DocumentProperty prop) = GetValueFromProp();
        //        if (exist)
        //        {
        //            if (prop.Value.ToString() == AttValue.ToString())
        //            {
        //                return true;
        //            }
        //            prop.Value = AttValue;
        //            if (showMsg)
        //            {
        //                MessageBox.Show(attName + " updated to " + AttValue.ToString());
        //            }
        //        }
        //        else
        //        {
        //            AllCustProps.Add(attName, false, MsoDocProperties.msoPropertyTypeString, AttValue.ToString());
        //            if (showMsg)
        //            {
        //                MessageBox.Show(attName + " added as " + AttValue.ToString());
        //            }
        //        }
        //        RefreshTextBox();
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Error adding value {AttValue} for {attName}: \n\n{ex.Message}", "Error");
        //        RefreshTextBox();
        //        return false;
        //    }
        //}

        public void RefreshTextBox()
        {
            if (currentValue != null)
            {
                textBox.Text = currentValue;
            }
            else
            {
                if (defaultValue != null)
                {
                    textBox.Text = defaultValue;
                    //SetValueFromTextBox();
                }
                else
                {
                    textBox.Clear();
                }
            }
        }

        protected virtual void SetValueFromTextBox(bool showMsg = false)
        {
            currentValue = textBox.Text;
        }
        #endregion

        #region Get Values for TextBox
        public double GetDoubleFromTextBox()
        {
            bool check = double.TryParse(textBox.Text, out double doubleValue);
            if (!check)
            {
                throw new Exception($"Unable to parse {textBox.Text} into number for {attName}");
            }
            else
            {
                return doubleValue;
            }
        }

        public float GetFloatFromTextBox()
        {
            bool check = float.TryParse(textBox.Text, out float floatValue);
            if (!check)
            {
                throw new Exception($"Unable to parse {textBox.Text} into number for {attName}");
            }
            else
            {
                return floatValue;
            }
        }

        public int GetIntFromTextBox()
        {
            bool check = int.TryParse(textBox.Text, out int value);
            if (!check)
            {
                throw new Exception($"Unable to parse {textBox.Text} into integer for {attName}");
            }
            else
            {
                return value;
            }
        }
        #endregion

        #region Import Export
        //public void ResetValue()
        //{
        //    (bool gotValue, DocumentProperty thisProp) = GetValueFromProp();
        //    if (gotValue)
        //    {
        //        thisProp.Delete();
        //    }
        //    RefreshTextBox();
        //    if (defaultValue != null)
        //    {
        //        textBox.Text = defaultValue;
        //    }
        //}

        //public virtual bool ImportValue(string value)
        //{
        //    try
        //    {
        //        textBox.Text = value;
        //        SetValueFromTextBox(false);
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}
        #endregion
    }


}
