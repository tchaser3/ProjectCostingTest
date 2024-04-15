/* Title:       Project Costing Test
 * Date:        2-6-20
 * Author:      Terry Holmes
 * 
 * Description: This is used to compute the costing */

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
using Microsoft.Win32;

namespace ProjectCostingTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();

        FindEmployeeProjectTotalHoursDataSet aFindProjectTotalHoursDataSet;
        FindEmployeeProjectTotalHoursDataSetTableAdapters.FindEmployeeProjectTotalHoursTableAdapter aFindProjectTotalHoursTableAdapter;
        FindEmployeeProjectTotalHoursDataSet TheFindProjectTotalHoursDataSet;
        ProjectTotalsDataSet TheProjectTotalsDataSet = new ProjectTotalsDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }
        private FindEmployeeProjectTotalHoursDataSet FindEmployeeProjectTotalHours()
        {
            aFindProjectTotalHoursDataSet = new FindEmployeeProjectTotalHoursDataSet();
            aFindProjectTotalHoursTableAdapter = new FindEmployeeProjectTotalHoursDataSetTableAdapters.FindEmployeeProjectTotalHoursTableAdapter();
            aFindProjectTotalHoursTableAdapter.Fill(aFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours);

            return aFindProjectTotalHoursDataSet;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            TheFindProjectTotalHoursDataSet = FindEmployeeProjectTotalHours();

            dgrResults.ItemsSource = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            string strAssignedProjectID;
            decimal decCoaxCost;
            decimal decCoaxHours;
            decimal decTechOpsCost;
            decimal decTechOpsHours;
            decimal decUndergroundCosts;
            decimal decUndergroundHours;
            decimal decFiberCosts;
            decimal decFiberHours;
            int intSecondCounter;
            int intSecondNumberOfRecords;

            try
            {
                intNumberOfRecords = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours.Rows.Count - 1;
                intSecondNumberOfRecords = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    decCoaxCost = 0;
                    decFiberCosts = 0;
                    decTechOpsCost = 0;
                    decUndergroundCosts = 0;
                    decCoaxHours = 0;
                    decFiberHours = 0;
                    decTechOpsHours = 0;
                    decUndergroundHours = 0;
                    strAssignedProjectID = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intCounter].AssignedProjectID;

                    for(intSecondCounter = 0; intSecondCounter <= intSecondNumberOfRecords; intSecondCounter++)
                    {
                        if(strAssignedProjectID == TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].AssignedProjectID)
                        {
                            if (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].Department == "COAX")
                            {
                                decCoaxCost += (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours * Convert.ToDecimal(49.35));
                                decCoaxHours += TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours;
                            }
                            else if(TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].Department == "TECH OPS")
                            {
                                decTechOpsCost += (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours * Convert.ToDecimal(51.39));
                                decTechOpsHours = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours;
                            }
                            else if (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].Department == "FIBER")
                            {
                                decFiberCosts += (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours * Convert.ToDecimal(51.39));
                                decFiberHours = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours;
                            }
                            else if (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].Department == "UNDERGROUND")
                            {
                                decUndergroundCosts += (TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours * Convert.ToDecimal(51.39));
                                decUndergroundHours = TheFindProjectTotalHoursDataSet.FindEmployeeProjectTotalHours[intSecondCounter].TotalHours;
                            }

                            intCounter++;
                        }

                    }

                    ProjectTotalsDataSet.projecttotalsRow NewProjectRow = TheProjectTotalsDataSet.projecttotals.NewprojecttotalsRow();

                    NewProjectRow.Coax = decCoaxCost;
                    NewProjectRow.CoaxHours = decCoaxHours;
                    NewProjectRow.ProjectID = strAssignedProjectID;
                    NewProjectRow.Fiber = decFiberCosts;
                    NewProjectRow.FiberHours = decFiberHours;
                    NewProjectRow.Techops = decTechOpsCost;
                    NewProjectRow.TechOpsHours = decTechOpsHours;
                    NewProjectRow.Underground = decUndergroundCosts;
                    NewProjectRow.UndergroundHours = decUndergroundHours;
                    NewProjectRow.TotalLaborCosts = (decCoaxCost + decFiberCosts + decTechOpsCost + decUndergroundCosts);

                    TheProjectTotalsDataSet.projecttotals.Rows.Add(NewProjectRow);

                    intCounter--;

                }

                dgrResults.ItemsSource = TheProjectTotalsDataSet.projecttotals;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Project Total Test // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheProjectTotalsDataSet.projecttotals.Rows.Count;
                intColumnNumberOfRecords = TheProjectTotalsDataSet.projecttotals.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheProjectTotalsDataSet.projecttotals.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheProjectTotalsDataSet.projecttotals.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Project Costing Test // Export To Excel " + ex.Message);

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
