using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System.Windows.Forms;
using System.IO;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Clash;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;


namespace NavisPlugin
{
    [Plugin("NavisPlugin", "CONN", DisplayName = "Navis Plugin :)")]
    public class ClAddin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            //MessageBox.Show("Hello World", "Execute", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //return 0;
            
            string grName = "";
            string grGuid = "";
            string grStatus = "";

            string connstring = Properties.Settings.Default.ConnectionString;

            SqlConnection conn = new SqlConnection(connstring);
            //string delstring = "TRUNCATE TABLE tblClashResult";
            //SqlCommand cmdDelete = new SqlCommand(delstring, conn);
            //cmdDelete.CommandType = CommandType.Text;

            //try
            //{
            //    conn.Open();
            //    cmdDelete.ExecuteNonQuery();
            //    MessageBox.Show("All Data deleted");

            //}
            //catch (SqlException e)
            //{
            //    MessageBox.Show(e.Message.ToString(), "Error Message");
            //}
            //finally
            //{
            //    conn.Close();
            //}


            MessageBox.Show("will dump clash to c:\\sqlite\\dumpClash.txt");
            StreamWriter sw = File.CreateText("c:\\sqlite\\dumpClash.txt");
            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
            
            DocumentClash documentClash = document.GetClash();
            DocumentClashTests oDCT = documentClash.TestsData;

            foreach (ClashTest test in oDCT.Tests)
            {
                sw.WriteLine("Connection String:" + connstring);
                sw.WriteLine("***Test: " + test.DisplayName + "***");
                //sw.WriteLine(" Status: " + test.Status.ToString());
                //sw.WriteLine(" Test Type: " + test.TestType.ToString());
                if (test.LastRun != null)
                {
                    sw.WriteLine(" Last Run: " + test.LastRun.Value.ToShortDateString());
                }
                //sw.WriteLine(" tolerance: " + test.Tolerance);
                //sw.WriteLine(" comments: " + test.Comments);
                //sw.WriteLine(" Simulation Type: " + test.SimulationType.ToString());
                //sw.WriteLine(" Simulation Step: " + test.SimulationStep.ToString());
                sw.WriteLine("    ---Results---");

                foreach (SavedItem issue in test.Children)
                {
                    ClashResultGroup group = issue as ClashResultGroup;
                    if (null != group)
                    {
                        sw.WriteLine(" test result group: " + group.DisplayName);
                        sw.WriteLine(" group status: " + group.Status.ToString());

                        grName = group.DisplayName;
                        grGuid = group.Guid.ToString();
                        grStatus = group.Status.ToString();

                        foreach (SavedItem issue1 in group.Children)
                        {
                            ClashResult rt1 = issue1 as ClashResult;
                            if (null != rt1)
                                
                                writeClashResult(test, rt1, sw, conn, grName, grGuid, grStatus);
                        }
                    }
                    ClashResult rt = issue as ClashResult;
                    
                    if (null != rt)
                        writeClashResult(test, rt, sw, conn, grName, grGuid, grStatus);

                }
            }
            MessageBox.Show("done");
            sw.Close();
            return 0;
        }

        private void writeClashResult(ClashTest test, ClashResult rt, StreamWriter sw, SqlConnection conn, string gName, string gGuid, string gStatus)
        {

            //if (rt.Status.ToString() == "Approved")
            //{
            sw.WriteLine("Clash Test: " + test.DisplayName);
            sw.WriteLine("  clash Name: " + rt.DisplayName);
            sw.WriteLine("  clash Guid: " + rt.Guid.ToString());

            SqlCommand cmdClashData = new SqlCommand("dbo.Procedure", conn);
            cmdClashData.CommandType = CommandType.StoredProcedure;

            cmdClashData.Parameters.Add("@testName", SqlDbType.NVarChar, 100).Value = test.DisplayName;
            cmdClashData.Parameters.Add("@testGuid", SqlDbType.NVarChar, 50).Value = test.Guid.ToString();

            if(test.LastRun != null)
            {
                cmdClashData.Parameters.Add("@testLastRunDate", SqlDbType.Date).Value = test.LastRun.Value.ToShortDateString();
                cmdClashData.Parameters.Add("@testLastRunTime", SqlDbType.Time).Value = test.LastRun.Value.ToShortTimeString();
            }
            
            cmdClashData.Parameters.Add("@testStatus", SqlDbType.NVarChar, 50).Value = test.Status.ToString();
            cmdClashData.Parameters.Add("@testType", SqlDbType.NVarChar, 50).Value = test.TestType.ToString();
            cmdClashData.Parameters.Add("@testTolerance", SqlDbType.Float, 53).Value = test.Tolerance;

            cmdClashData.Parameters.Add("@groupName", SqlDbType.NVarChar, 50).Value = gName;
            cmdClashData.Parameters.Add("@groupGuid", SqlDbType.NVarChar, 50).Value = gGuid;
            cmdClashData.Parameters.Add("@groupStatus", SqlDbType.NVarChar, 50).Value = gStatus;
            
            cmdClashData.Parameters.Add("@clashName", SqlDbType.NVarChar, 50).Value = rt.DisplayName;
            cmdClashData.Parameters.Add("@clashGuid", SqlDbType.NVarChar, 50).Value = rt.Guid.ToString();

            if (rt.CreatedTime != null)
            {
                sw.WriteLine("  Created Time: " + rt.CreatedTime.ToString());
                cmdClashData.Parameters.Add("@clashTime", SqlDbType.SmallDateTime).Value = rt.CreatedTime.Value.ToShortDateString();
            }
            
            cmdClashData.Parameters.Add("@clashStatus", SqlDbType.NVarChar, 50).Value = rt.Status.ToString();
            cmdClashData.Parameters.Add("@clashApprovedBy", SqlDbType.NVarChar, 50).Value = rt.ApprovedBy;
            cmdClashData.Parameters.Add("@clashApprovedTime", SqlDbType.SmallDateTime).Value = rt.ApprovedTime.Value.ToShortDateString();
            cmdClashData.Parameters.Add("@clashX", SqlDbType.Float, 53).Value = rt.Center.X;
            cmdClashData.Parameters.Add("@clashY", SqlDbType.Float, 53).Value = rt.Center.Y;
            cmdClashData.Parameters.Add("@clashZ", SqlDbType.Float, 53).Value = rt.Center.Z;

            try
            {
                conn.Open();

                cmdClashData.ExecuteNonQuery();
                
            }
            catch(SqlException e)
            {
                MessageBox.Show(e.Message.ToString(), "Error Message");
            }
            finally
            {
                conn.Close();
            }

                
                sw.WriteLine("  Status: " + rt.Status.ToString());
                sw.WriteLine("  Approved By: " + rt.ApprovedBy);
                sw.WriteLine("  Approved Date: " + rt.ApprovedTime.ToString());
                //if (rt.Center != null)
                //    sw.WriteLine(" Centre[{0}, {1}, {2}]: ", rt.Center.X, rt.Center.Y, rt.Center.Z);
                //sw.WriteLine(" Simulation Type: " + rt.SimulationType.ToString());
            //}

        }
    }
}
