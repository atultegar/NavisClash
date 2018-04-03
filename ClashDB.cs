using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System.IO;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Clash;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;

namespace NavisClash
{
    [Plugin("NavisClash", "CONN", DisplayName = "Clash to DB")]
    
    public class ClashDB:AddInPlugin
    {
        
        public override int Execute(params string[] parameters)
        {
            MessageBox.Show("Hello World!", "Execute", MessageBoxButtons.OK, MessageBoxIcon.Information);

            string connstr = "Data Source = (localdb)\\ProjectsV13; Initial Catalog = ClashDB; Integrated Security = True; Connect Timeout = 30; Encrypt = False; TrustServerCertificate = True; ApplicationIntent = ReadWrite; MultiSubnetFailover = False";
            //SQL Connection
            SqlConnection conn = new SqlConnection(connstr);

            //SQL Command - Stored Procedure
            SqlCommand cmdClashData = new SqlCommand("dbo.createClashData", conn);
            cmdClashData.CommandType = CommandType.StoredProcedure;

            //Create SQL Parameters
            
            cmdClashData.Parameters.Add(new SqlParameter("@testName", SqlDbType.NVarChar, 100));
            cmdClashData.Parameters.Add(new SqlParameter("@testStatus", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@testType", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@testLastRun", SqlDbType.SmallDateTime));
            cmdClashData.Parameters.Add(new SqlParameter("@testTolerance", SqlDbType.Float, 53));
            cmdClashData.Parameters.Add(new SqlParameter("@groupName", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@groupStatus", SqlDbType.NVarChar, 50));
            

            //Navisworks Document Clash Data
            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
            DocumentClash documentClash = document.GetClash();
            DocumentClashTests oDCT = documentClash.TestsData;

            //Loop for getting clash test results
            foreach (ClashTest test in oDCT.Tests)
            {
                cmdClashData.Parameters["@testName"].Value = test.DisplayName;
                
                cmdClashData.Parameters["@testStatus"].Value = test.Status.ToString();
                cmdClashData.Parameters["@testType"].Value = test.TestType.ToString();
                if (test.LastRun != null)
                {
                    cmdClashData.Parameters["@testLastRun"].Value = test.LastRun.Value.ToShortDateString();
                }
                cmdClashData.Parameters["@testTolerance"].Value = test.Tolerance;

                conn.Open();

                cmdClashData.ExecuteNonQuery();

                foreach (SavedItem issue in test.Children)
                {
                    ClashResultGroup group = issue as ClashResultGroup;
                    if (null != group)
                    {
                        cmdClashData.Parameters["@groupName"].Value = group.DisplayName;
                        cmdClashData.Parameters["@groupStatus"].Value = group.Status.ToString();

                        cmdClashData.ExecuteNonQuery();

                        foreach (SavedItem issue1 in group.Children)
                        {
                            ClashResult rt1 = issue1 as ClashResult;
                            if (null != rt1)
                                writeClashResult(rt1, cmdClashData);

                        }
                    }

                    ClashResult rt = issue as ClashResult;
                    if (null != rt)
                        writeClashResult(rt, cmdClashData);
                }

            }

            MessageBox.Show("Done");
            conn.Close();
            //string sql = "INSERT INTO [ClashDB].[Table] ";
            //SqlCommand cmdAddClashData = new SqlCommand
            return 0;
        }

        private void writeClashResult(ClashResult rt, SqlCommand cmdClashData)
        {
            cmdClashData.Parameters.Add(new SqlParameter("@clashGuid", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@clashName", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@clashTime", SqlDbType.SmallDateTime));
            cmdClashData.Parameters.Add(new SqlParameter("@clashStatus", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@clashApprovedBy", SqlDbType.NVarChar, 50));
            cmdClashData.Parameters.Add(new SqlParameter("@clashApprovedOn", SqlDbType.SmallDateTime));

            cmdClashData.Parameters["@clashGuid"].Value = rt.Guid.ToString();
            cmdClashData.Parameters["@clashName"].Value = rt.DisplayName;

            if (rt.CreatedTime != null)
                cmdClashData.Parameters["@clashTime"].Value = rt.CreatedTime.Value.ToShortDateString();

            cmdClashData.Parameters["@clashStatus"].Value = rt.Status.ToString();
            cmdClashData.Parameters["@clashApprovedBy"].Value = rt.ApprovedBy;
            cmdClashData.Parameters["@clashApprovedOn"].Value = rt.ApprovedTime.Value.ToShortTimeString();

            cmdClashData.ExecuteNonQuery();

        }
    }
}
