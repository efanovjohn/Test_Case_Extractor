using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using System.ComponentModel;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Windows.Interop;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using Microsoft.VisualBasic;
//using OfficeOpenXml;
using System.IO;
using Microsoft.TeamFoundation.Server;


namespace Test_Case_Extractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private TfsTeamProjectCollection _tfs=null;
        private WorkItemStore _store = null;
        ITestManagementTeamProject _testproject = null;
        ITestPlan TP = null;
        
        
        

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_connect_Click(object sender, RoutedEventArgs e)
        {
            this._tfs = null;
            Sel_TPlan.Items.Clear();
            treeView_suite.Items.Clear();
            TFS_Textbox.Text = null;
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();
            
            if (tpp.SelectedTeamProjectCollection != null)
            {
                
                this._tfs = tpp.SelectedTeamProjectCollection;
                
                ITestManagementService test_service = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));
                _store = (WorkItemStore)_tfs.GetService(typeof(WorkItemStore));
               
                TFS_Textbox.Text = this._tfs.Name;

                    string proj_name = tpp.SelectedProjects[0].Name;
                    _testproject = test_service.GetTeamProject(proj_name);
                    if (_testproject != null)
                    {
                        TFS_Textbox.Text = TFS_Textbox.Text + "\\" + _testproject.ToString();
                        GetTestPlans(_testproject);
                    }
                    else
                        MessageBox.Show("Please select a valid Team Project");

                
            }
        }

        void GetTestPlans(ITestManagementTeamProject _testproject)
        {
            
            Sel_TPlan.Visibility = Visibility.Visible;
            Lbl_TPlan.Visibility=Visibility.Visible;
            Lbl_TSuites.Visibility = Visibility.Visible;
            Gen_Btn.Visibility = Visibility.Visible;

            treeView_suite.Background = System.Windows.Media.Brushes.White;
            treeView_suite.BorderBrush = System.Windows.Media.Brushes.Black;



            foreach (ITestPlan tp in _testproject.TestPlans.Query("Select * from TestPlan"))
            {
                string t_plan = tp.Name  +" <ID: " + tp.Id.ToString() + " >";
                Sel_TPlan.Items.Add(t_plan);
            }
                    
                  


        }

     

        void GetTestCases(ITestCaseCollection testcases, ExcelWorksheet oSheet)
        {
            
            int i = 2;
            
            foreach (ITestCase Testcase in testcases)
            {
                int j = 1;
                string str1 = null;
                string str2 = null;
                foreach (ITestAction action in Testcase.Actions)
                {
                    ISharedStep shared_step = null;
                    ISharedStepReference shared_ref = action as ISharedStepReference;
                    if (shared_ref != null)
                    {
                        shared_step = shared_ref.FindSharedStep();
                        foreach (ITestAction shr_action in shared_step.Actions)
                        {
                            var test_step = shr_action as ITestStep;
                            str1 = str1 + j.ToString() + "." + ((test_step.Title.ToString().Length ==0)? "<<Not Recorded>>" : test_step.Title.ToString()) + System.Environment.NewLine;
                            str2 = str2 + j.ToString() + "." + ((test_step.ExpectedResult.ToString().Length ==0) ? "<<Not Recorded>>" : test_step.ExpectedResult.ToString()) + System.Environment.NewLine;
                            j++;
                        }
                       
                    }
                    else
                    {
                        var test_step = action as ITestStep;
                        str1 = str1 + j.ToString() + "." + ((test_step.Title.ToString().Length ==0) ? "<<Not Recorded>>" : test_step.Title.ToString()) + System.Environment.NewLine;
                        str2 = str2 + j.ToString() + "." + ((test_step.ExpectedResult.ToString().Length ==0) ? "<<Not Recorded>>" : test_step.ExpectedResult.ToString()) + System.Environment.NewLine;
                        j++;
                    }
                }
              



                string result = null;
                string tot_result = null;
                foreach (ITestPoint test_point in TP.QueryTestPoints(string.Format("Select * from TestPoint where TestCaseId = {0} ", Testcase.Id)))
                {
                    ITestCaseResult tc_res = test_point.MostRecentResult;
                    if (tc_res != null)
                    {
                        if (tc_res.Outcome.ToString().Equals("None"))
                           result= tc_res.TestConfigurationName +  " : In Progress" + System.Environment.NewLine;
                        else
                            result = tc_res.TestConfigurationName + " : " + tc_res.Outcome.ToString() + System.Environment.NewLine;
                    }
                    else
                        result = test_point.ConfigurationName + " : Active" + System.Environment.NewLine;

                    tot_result = tot_result + result;
                }
                
           
               Query query = new Query(_store,string.Format("SELECT [Target].[System.Id] FROM WorkItemLinks WHERE ([Source].[System.Id] = {0}) and ([Source].[System.WorkItemType] = 'Test Case')  And ([Target].[System.WorkItemType] = 'Bug')mode(MustContain)", Testcase.Id));
              
                WorkItemLinkInfo[] workItemLinkInfoArray = null;
                if (query.IsLinkQuery)
               {

                   workItemLinkInfoArray = query.RunLinkQuery();

               }

               else
               {

                   throw new Exception("Run link query fail. Query passed is not a link query");

               }
                string bug_list = null;
                for (int k = 0; k < workItemLinkInfoArray.Length; k++)
                {
                    if (workItemLinkInfoArray[k].LinkTypeId != 0)
                        bug_list = bug_list + workItemLinkInfoArray[k].TargetId.ToString() + System.Environment.NewLine;
                }

              

                oSheet.Cells[i, 1].Value = Testcase.Id.ToString();
                oSheet.Cells[i, 2].Value = Testcase.Title.ToString();
                oSheet.Cells[i, 3].Value = str1;
                oSheet.Cells[i, 4].Value = str2;
                oSheet.Cells[i, 8].Value = Testcase.WorkItem.CreatedBy.ToString();
                oSheet.Cells[i, 5].Value = Testcase.Description.ToString() + " ";
                oSheet.Cells[i, 9].Value = tot_result;
                oSheet.Cells[i, 10].Value = bug_list;
                

                
                
                i++;
            }
        }

      
        void Access_Excel(ITestCaseCollection testcases)
        {
            try
            {

               
                
                FileInfo new_file = new FileInfo(File_Name.Text);
              
               // FileInfo template = new FileInfo(System.Windows.Forms.Application.StartupPath + "\\Resources\\Test_Case_Template.xlsx");
                FileInfo template = new FileInfo(Directory.GetCurrentDirectory() + "\\Test_Case_Template.xlsx");
                using (ExcelPackage xlpackage = new ExcelPackage(new_file,template))
                {
                    ExcelWorksheet worksheet = xlpackage.Workbook.Worksheets["Test Case"];
                   
                    GetTestCases(testcases, worksheet);
                    xlpackage.Save();

                    MessageBox.Show("File has been saved at " + File_Name.Text);
            }


                             
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");

            }
        }

        private void Sel_TPlan_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Sel_TPlan.SelectedValue != null)
            {
                string TP_Name = Sel_TPlan.SelectedValue.ToString();
                treeView_suite.Items.Clear();
                int ID = get_ID(TP_Name);
                GetTestSuites_Plan(ID);
            }
            



        }


        void  GetTestSuites_Plan(int ID)
        {

             TP = _testproject.TestPlans.Find(ID);
            if (TP != null)              
            { 
            if(TP.RootSuite != null)
                
                {
               
                    TreeViewItem root = new TreeViewItem();
                    root.Header = TP.RootSuite.Title.ToString() + " <ID: " + TP.RootSuite.Id.ToString() + " >";
                    treeView_suite.Items.Add(root);
                                     
                    GetSuiteEntries(TP.RootSuite.SubSuites, root);

                    
                }
            }

        }

    

        int get_ID(string name)
        {
            
            int ID= Convert.ToInt32(name.Substring((name.LastIndexOf("<ID: ") + 5), (name.LastIndexOf(" >") - (name.LastIndexOf("<ID: ") + 5))));
            
            return ID;
        }


        void GetSuiteEntries(ITestSuiteCollection suite_entries, TreeViewItem root)
        {
            foreach (ITestSuiteBase suite in suite_entries)
            {
                
               if(suite!=null)
               {
                TreeViewItem suite_tree = new TreeViewItem();
                suite_tree.Header = suite.Title.ToString() + " <ID: " + suite.Id.ToString() + " >";
                root.Items.Add(suite_tree);
             
              
                if (suite.TestSuiteType == TestSuiteType.StaticTestSuite)
                {
                    IStaticTestSuite suite1 = suite as IStaticTestSuite;
                   if(suite1!=null && (suite1.SubSuites.Count > 0))
                   {
                      
                       GetSuiteEntries(suite1.SubSuites, suite_tree);
                   }
               }
              }

            }
        }

      
        private void Gen_Btn_Click(object sender, RoutedEventArgs e)
        {
            if (File_Name.Text == null || File_Name.Text.Length.Equals(0))
            {
                MessageBox.Show("Please Enter a valid file path");
            }
            else
            {
                ITestCaseCollection test_cases = null;
                if (treeView_suite.SelectedValue != null)
                {


                    int suite_ID = get_ID(treeView_suite.SelectedValue.ToString());
                    ITestSuiteBase suite = _testproject.TestSuites.Find(suite_ID);
                    test_cases = suite.AllTestCases;
                    if (test_cases != null)
                        Access_Excel(test_cases);

                }

                else
                {
                    MessageBox.Show("Please select a test suite");
                }
            }
                                      
        }

      

        private void button1_Click(object sender, RoutedEventArgs e)
        {

            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            saveFileDialog1.Filter = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                File_Name.Text = saveFileDialog1.FileName;
            }

            else
            {
                MessageBox.Show("Please choose a valid filename");
            }

        }

        private void treeView_suite_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

      
      

    }
}
