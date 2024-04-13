/* Title:           Edit Phones
 * Date:            4-4-19
 * Author:          Terry Holmes
 * 
 * Description:     This will allow the user to edit a phone */

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
using System.Windows.Shapes;
using NewEventLogDLL;
using NewEmployeeDLL;

namespace ImportPhones
{
    /// <summary>
    /// Interaction logic for EditPhone.xaml
    /// </summary>
    public partial class EditPhone : Window
    {
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();

        //setting up the data
        ComboEmployeeDataSet TheComboEmployeeDataSet = new ComboEmployeeDataSet();

        int gintSelectedIndex;
        int gintEmployeeID;
        string gstrFirstName;
        string gstrLastName;

        public EditPhone()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;

            try
            {
                intNumberOfRecords = MainWindow.TheImportPhonesDataSet.importphones.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if (MainWindow.TheImportPhonesDataSet.importphones[intCounter].TransactionID == MainWindow.gintTransactionID)
                    {
                        txtExtension.Text = Convert.ToString(MainWindow.TheImportPhonesDataSet.importphones[intCounter].Extension);
                        gintSelectedIndex = intCounter;
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Phones // Edit Phone // Window Loaded " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void TxtEnterLastName_TextChanged(object sender, TextChangedEventArgs e)
        {
            string strLastName;
            int intLength;
            int intNumberOfRecords;
            int intCounter;

            try
            {
                strLastName = txtEnterLastName.Text;
                cboSelectEmployee.Items.Clear();
                cboSelectEmployee.Items.Add("Select Employee");
                intLength = strLastName.Length;

                if(intLength > 2)
                {
                    TheComboEmployeeDataSet = TheEmployeeClass.FillEmployeeComboBox(strLastName);

                    intNumberOfRecords = TheComboEmployeeDataSet.employees.Rows.Count - 1;

                    if(intNumberOfRecords > -1)
                    {
                        for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                        {
                            cboSelectEmployee.Items.Add(TheComboEmployeeDataSet.employees[intCounter].FullName);
                        }
                    }

                    cboSelectEmployee.SelectedIndex = 0;
                }
            }
            catch (Exception Ex)
            {
                TheMessagesClass.ErrorMessage(Ex.ToString());

                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Phones // Edit Phone // Enter Last Name " + Ex.Message);
            }
        }

        private void CboSelectEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;

            intSelectedIndex = cboSelectEmployee.SelectedIndex - 1;

            if(intSelectedIndex > -1)
            {
                gintEmployeeID = TheComboEmployeeDataSet.employees[intSelectedIndex].EmployeeID;
                gstrFirstName = TheComboEmployeeDataSet.employees[intSelectedIndex].FirstName;
                gstrLastName = TheComboEmployeeDataSet.employees[intSelectedIndex].LastName;
            }
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.TheImportPhonesDataSet.importphones[gintSelectedIndex].EmployeeID = gintEmployeeID;
            MainWindow.TheImportPhonesDataSet.importphones[gintSelectedIndex].FirstName = gstrFirstName;
            MainWindow.TheImportPhonesDataSet.importphones[gintSelectedIndex].LastName = gstrLastName;

            Close();
        }
    }
}
