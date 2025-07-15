using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ListBox = System.Windows.Forms.ListBox;

namespace FidoAutoCad.SharedForms
{
    public partial class ListSelector : Form
    {
        public ListSelector()
        {
            InitializeComponent();
            CancelButton = MyCancelButton;
        }
        public ListSelector(string formTitle, string description, string leftBoxLabel, string rightBoxLabel)
        {
            InitializeComponent();
            CancelButton = MyCancelButton;
            Text = formTitle;
            descriptionLabel.Text = description;
            LeftLabel.Text = leftBoxLabel;
            RightLabel.Text = rightBoxLabel;
        }

        #region Initializer
        Dictionary<string, ListItem> tracker;
        public void AddTracker(ref Dictionary<string, ListItem> tracker)
        {
            this.tracker = tracker;
            Load += new EventHandler((object sender, EventArgs e) => RenderFromTracker());
        }
        private void RenderFromTracker()
        {
            List<ListItem> activeNames = new List<ListItem>();
            List<ListItem> inactiveNames = new List<ListItem>();

            foreach (ListItem item in tracker.Values)
            {
                if (item.isActive) { activeNames.Add(item); }
                else { inactiveNames.Add(item); }
            }

            var testVar = activeNames.OrderBy(x => x.selectedOrderNum).ToList();
            var testVar2 = inactiveNames.OrderBy(x => x.selectedOrderNum).ToList();

            foreach (ListItem item in testVar) 
            { 
                RightListBox.Items.Add(item.name); 
            }
            foreach (ListItem item in testVar2) 
            { 
                LeftListBox.Items.Add(item.name); 
            }

            //foreach (KeyValuePair<string, ListItem> item in tracker)
            //{
            //    if (item.Value.isActive) { RightListBox.Items.Add(item.Key); }
            //    else { LeftListBox.Items.Add(item.Key); }
            //}
        }
        #endregion

        #region Move Sheet Buttons
        #region Helper Functions
        private void MoveAllItems(ListBox source, ListBox destination)
        {
            // Check if source and destination are not null
            if (source == null || destination == null)
                throw new ArgumentNullException("ListBoxes cannot be null.");

            // Add all items from source to destination
            destination.Items.AddRange(source.Items);
            // Clear all items from the source listBox
            source.Items.Clear();
        }

        private void MoveSelectedItems(ListBox source, ListBox destination)
        {
            // Check if source and destination are not null
            if (source == null || destination == null)
                throw new ArgumentNullException("ListBoxes cannot be null.");

            // Collect the selected items in an array
            var selectedItems = new object[source.SelectedItems.Count];
            source.SelectedItems.CopyTo(selectedItems, 0);

            // Add selected items to the destination
            destination.Items.AddRange(selectedItems);

            // Remove selected items from the source
            foreach (var item in selectedItems)
            {
                source.Items.Remove(item);
            }
        }
        #endregion

        private void MoveAllRight_Click(object sender, EventArgs e)
        {
            MoveAllItems(LeftListBox, RightListBox);
        }

        private void MoveSelectionRight_Click(object sender, EventArgs e)
        {
            MoveSelectedItems(LeftListBox, RightListBox);
        }

        private void MoveSelectionLeft_Click(object sender, EventArgs e)
        {
            MoveSelectedItems(RightListBox, LeftListBox);
        }

        private void MoveAllLeft_Click(object sender, EventArgs e)
        {
            MoveAllItems(RightListBox, LeftListBox);
        }
        #endregion

        #region Confirmation Buttons

        #region Helper Functions
        private List<string> GetListBoxItems(ListBox listBox)
        {
            List<string> items = new List<string>();
            foreach (var item in listBox.Items)
            {
                items.Add(item.ToString());
            }
            return items;
        }

        private List<string> GetListBoxSelectedItems(ListBox listBox)
        {
            List<string> items = new List<string>();
            foreach (var item in listBox.SelectedItems)
            {
                items.Add(item.ToString());
            }
            return items;
        }
        #endregion
        private void ConfirmationButton_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (string item in LeftListBox.Items) //Unselected items
                {
                    tracker[item].isActive = false;
                    tracker[item].selectedOrderNum = -1;
                }

                int itemNum = 1;
                foreach (string item in RightListBox.Items) //Selected items
                {
                    tracker[item].isActive = true;
                    tracker[item].selectedOrderNum = itemNum;
                    itemNum++;
                }
                Close();
            }
            catch (Exception ex) { MessageBox.Show($"{ex.Message}", "Error"); }

        }

        private void MyCancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion
    }
    public class ListItem
    {
        public string name { get; set; }
        public bool isActive { get; set; }
        public int initialOrderNum { get; set; }
        public int selectedOrderNum { get; set; }
        public ListItem(string name, int initialOrderNum, bool status = true)
        {
            this.name = name;
            this.initialOrderNum = initialOrderNum;
            this.isActive = status;
        }
        public ListItem(string name, int initialOrderNum, int selectedOrderNum)
        {
            this.name = name;
            this.initialOrderNum = initialOrderNum;
            this.isActive = true;
            this.selectedOrderNum = selectedOrderNum;
        }
    }
}
