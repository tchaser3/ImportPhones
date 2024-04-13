/* Title:           Import Phones
 * Date:            4-04-19
 * Author:          Terry Holmes
 * 
 * Description:     This is used to import phone information */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEventLogDLL;
using NewEmployeeDLL;
using PhonesDLL;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportPhones
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        PhonesClass ThePhoneClass = new PhonesClass();

        //setting update data
        public static ImportPhonesDataSet TheImportPhonesDataSet = new ImportPhonesDataSet();
        FindEmployeeByLastNameDataSet TheFindEmployeeByLastNameDataSet = new FindEmployeeByLastNameDataSet();

        public static int gintTransactionID;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strExtension;
            int intExtension;
            string strFirstName;
            string strLastName;
            string strMACAddress;
            string strDIDNumber;
            int intRecordsReturned;
            int intEmployeeCounter;
            int intEmployeeID;
            int intWarehouseID;

            try
            {
                TheImportPhonesDataSet.importphones.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 2; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strExtension = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2).ToUpper();
                    intExtension = Convert.ToInt32(strExtension);
                    strFirstName = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2).ToUpper();
                    strLastName = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2).ToUpper();
                    strMACAddress = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2).ToUpper();
                    strDIDNumber = Convert.ToString((range.Cells[intCounter, 5] as Excel.Range).Value2).ToUpper();

                    if ((intExtension > 1099) && (intExtension < 1200))
                    {
                        intWarehouseID = 100000;
                    }
                    else if ((intExtension > 1199) && (intExtension < 1300))
                    {
                        intWarehouseID = 2014;
                    }
                    else if ((intExtension > 1299) && (intExtension < 1400))
                    {
                        intWarehouseID = 1343;
                    }
                    else if ((intExtension > 1399) && (intExtension < 1500))
                    {
                        intWarehouseID = 2122;
                    }
                    else
                    {
                        intWarehouseID = -1;
                    }


                    TheFindEmployeeByLastNameDataSet = TheEmployeeClass.FindEmployeesByLastNameKeyWord(strLastName);

                    intRecordsReturned = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName.Rows.Count - 1;
                    intEmployeeID = -1;

                    if(intRecordsReturned > -1)
                    {
                        for(intEmployeeCounter = 0; intEmployeeCounter <= intRecordsReturned; intEmployeeCounter++)
                        {
                            if(strFirstName == TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intEmployeeCounter].FirstName)
                            {
                                intEmployeeID = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intEmployeeCounter].EmployeeID;
                            }
                        }
                    }

                    ImportPhonesDataSet.importphonesRow NewPhoneRow = TheImportPhonesDataSet.importphones.NewimportphonesRow();

                    NewPhoneRow.DID = strDIDNumber;
                    NewPhoneRow.EmployeeID = intEmployeeID;
                    NewPhoneRow.Extension = intExtension;
                    NewPhoneRow.FirstName = strFirstName;
                    NewPhoneRow.LastName = strLastName;
                    NewPhoneRow.MACAddress = strMACAddress;
                    NewPhoneRow.WarehouseID = intWarehouseID;

                    TheImportPhonesDataSet.importphones.Rows.Add(NewPhoneRow);
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportPhonesDataSet.importphones;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Phones // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void DgrResults_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataGrid dataGrid;
            DataGridRow selectedRow;
            DataGridCell TransactionID;
            string strTransactionID;

            try
            {
                if (dgrResults.SelectedIndex > -1)
                {
                    //setting local variable
                    dataGrid = dgrResults;
                    selectedRow = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(dataGrid.SelectedIndex);
                    TransactionID = (DataGridCell)dataGrid.Columns[0].GetCellContent(selectedRow).Parent;
                    strTransactionID = ((TextBlock)TransactionID.Content).Text;
                    gintTransactionID = Convert.ToInt32(strTransactionID);

                    EditPhone EditPhone = new EditPhone();
                    EditPhone.ShowDialog();
                    
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Phones // Grid View Selection " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError;
            int intExtension;
            int intEmployeeID;
            string strDIDNumber;
            int intWarehouseID;
            string strMACAddress;

            try
            {
                intNumberOfRecords = TheImportPhonesDataSet.importphones.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intExtension = TheImportPhonesDataSet.importphones[intCounter].Extension;
                    intEmployeeID = TheImportPhonesDataSet.importphones[intCounter].EmployeeID;
                    strDIDNumber = TheImportPhonesDataSet.importphones[intCounter].DID;
                    intWarehouseID = TheImportPhonesDataSet.importphones[intCounter].WarehouseID;
                    strMACAddress = TheImportPhonesDataSet.importphones[intCounter].MACAddress;

                    blnFatalError = ThePhoneClass.InsertPhone(intExtension, strDIDNumber, intEmployeeID, intWarehouseID, strMACAddress);

                    if (blnFatalError == true)
                        throw new Exception();
                }

                TheMessagesClass.InformationMessage("Phones Have Been Entered");
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Phones // Main window // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            } 

        }
    }
}
