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

namespace NavisApprove
{
    [Plugin("NavisApprove", "CONN", DisplayName = "Navis Approve Clashes")]
    public class ClassAddin : AddInPlugin
    {
        static Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
        static DocumentClash documentClash = document.GetClash();
        static DocumentClashTests oDCT = documentClash.TestsData;
        public override int Execute(params string[] parameters)
        {
            string grName = "";
            string grGuid = "";
            //string grStatus = "";

            string connstring = Properties.Settings.Default.ConnString;
            SqlConnection conn = new SqlConnection(connstring);
                      

            foreach (ClashTest test in oDCT.Tests)
            {
                foreach (SavedItem issue in test.Children)
                {
                    ClashResultGroup group = issue as ClashResultGroup;
                    if (null != group)
                    {
                        foreach (SavedItem issue1 in group.Children)
                        {
                            ClashResult rt1 = issue as ClashResult;
                            if (null != rt1)
                            {
                                writeClashResult(test, rt1, conn, grName, grGuid);
                            }
                        }
                    }
                }
            }
            MessageBox.Show("Clash Result Status Updated", "Execute");
            return 0;
        }

        private void writeClashResult(ClashTest test, ClashResult rt, SqlConnection conn, string grName, string grGuid)
        {
            //SqlCommand cmdApproveClash = new SqlCommand("dbo.selectApproved", conn);
            //cmdApproveClash.CommandType = CommandType.StoredProcedure;

            //cmdApproveClash.Parameters.Add("@clashGuid", SqlDbType.NVarChar, 50).Value = rt1.Guid.ToString();
            //cmdApproveClash.Parameters.Add("@clashStatus", SqlDbType.NVarChar, 50);

            string rtGuid = rt.Guid.ToString();
            string sqlQuery = $"SELECT ClashGuid, ClashStatus, ClashApprovedBy, ClashApprovedTime FROM tblClashResult WHERE ClashGuid={rtGuid}";
            SqlCommand myCommand = new SqlCommand(sqlQuery, conn);
            try
            {
                conn.Open();
                // Obtain a data reader ExecuteReader().
                using (SqlDataReader myDataReader = myCommand.ExecuteReader())
                {
                    while (myDataReader.Read())
                    {
                        if (myDataReader["ClashStatus"].ToString() == "Approved")
                        {
                            string rtApprovedBy = $"{myDataReader["ClashApprovedBy"]}";
                            oDCT.TestsEditResultStatus(rt, ClashResultStatus.Approved);
                            oDCT.TestsEditResultApprovedBy(rt, rtApprovedBy);

                        }
                    }
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message.ToString(), "Error Message");
            }
            finally
            {
                conn.Close();
            }
            
        }
    }
}