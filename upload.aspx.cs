using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Data_Warehouse_Web_Components.DTO;
using System.DirectoryServices;
using System.IO;
using System.Data.OleDb;
using System.Data;
using EntitySpaces.Interfaces;
using System.Diagnostics;
using System.Globalization;

namespace Data_Warehouse_Web_Components.RDW
{
    public partial class eT : System.Web.UI.Page
    {
        public string xlsDataSource
        {
            get { return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes'"; }
        }

        public string xlsxDataSurce
        {
            get { return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes'"; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack)
            {
                lblErrorMsg.Text = lblReport.Text = string.Empty;
            }
            else
            {
                btnSubmit.Visible = false;
            }
        }

        #region - Internal Methods -

        private void BindDataGrid(List<KDTO> bindobj)
        {
            try
            {
                grdK.DataSource = bindobj; //.OrderBy(p => p.MainCatSortId).ThenBy(p => p.SubCatSortId);
                grdK.DataBind();

                grdExcelData.DataSource = null;
                grdExcelData.DataBind();
            }
            catch (Exception ex)
            {
                lblErrorMsg.Text = ex.Message;
            }
        }

        #endregion - Internal Methods -

        #region - Events -

        
        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                if (ViewState["Excel Data"] != null)
                {
                    DataTable dt = (DataTable)ViewState["Excel Data"];// grdExcelData.DataSource;

                    SubmitExcelData(dt);

                    ViewState["Excel Data"] = null;
                }
                if (ViewState["Excel Data 2"] != null)
                {
                    DataTable dt2 = (DataTable)ViewState["Excel Data 2"];// grdExcelData.DataSource;
                    SubmitExcelData(dt2);

                    ViewState["Excel Data 2"] = null;
                }
                if (ViewState["Excel Data 3"] != null)
                {
                    DataTable dt3 = (DataTable)ViewState["Excel Data 3"];// grdExcelData.DataSource;
                    SubmitExcelData(dt3);

                    ViewState["Excel Data 3"] = null;
                }

            }
            catch (Exception ex)
            {
                lblErrorMsg.Text = ex.Message;
            }
        }

      
        protected void dd_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSubmit.Visible = false;
            uploadRow.Visible = false;
            Button1.Visible = true;
            //txtDate.Enabled = true;

        }

        #endregion - Events -

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (fileuploader.HasFile)
            {
                string FileName = Path.GetFileName(fileuploader.PostedFile.FileName);
                string Extension = Path.GetExtension(fileuploader.PostedFile.FileName);

                if (Extension.ToLower() == ".xls" || Extension.ToLower() == ".xlsx")
                {
                    string FilePath = Server.MapPath("Files/" + FileName);
                    fileuploader.SaveAs(FilePath);

                    Import_To_Grid(FilePath, Extension);

                    try
                    {
                        if (System.IO.File.Exists(FilePath))
                            File.Delete(FilePath);
                    }
                    catch (Exception)
                    {

                    }
                }
                else
                    lblErrorMsg.Text = "Choose only *.xls or *.xlsx format files.";
            }
            else
                lblErrorMsg.Text = "Choose the file to upload.";
        }

        private void Import_To_Grid(string FilePath, string Extension)
        {
            try
            {
                //ePodDataLst = null;
                string conStr = "";
                switch (Extension)
                {
                    case ".xls": //Excel 97-03
                        conStr = xlsDataSource;
                        break;
                    case ".xlsx": //Excel 07
                        conStr = xlsxDataSurce;
                        break;
                }

                conStr = String.Format(conStr, FilePath);
                OleDbConnection connExcel = new OleDbConnection(conStr);
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                cmdExcel.Connection = connExcel;

                //Get the name of First Sheet
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dtExcelSchema.Rows.Count > 0)
                {
                    int i = 0;

                    // Bind the sheets to the Grids  
                    foreach (DataRow row in dtExcelSchema.Rows)
                    {
                        DataTable dt_sheet = null;
			//exclude sheets that contain the following
                        if (row["TABLE_NAME"].ToString().Contains("Instructions") || row["TABLE_NAME"].ToString().Contains("Sheet") ||
                            row["TABLE_NAME"].ToString().Contains("Summary") || row["TABLE_NAME"].ToString().Contains("FilterDatabase") ||
                             row["TABLE_NAME"].ToString().Contains("Country") || row["TABLE_NAME"].ToString().Contains("Customer") ||
                             row["TABLE_NAME"].ToString().Contains("List") || row["TABLE_NAME"].ToString().Contains("Period") ||
                             row["TABLE_NAME"].ToString().Contains("Site") || row["TABLE_NAME"].ToString().Contains("Print_Titles")
                            )
                        { }
                        else
                        {
                            dt_sheet = getSheetData(conStr, row["TABLE_NAME"].ToString());
                            Debug.WriteLine(i + " " + row["TABLE_NAME"].ToString());
                            switch (i)
                            {
                                case 0:
                                    grdExcelData.DataSource = dt_sheet;
                                    grdExcelData.DataBind();
                                    ViewState["Excel Data"] = this.getFilteredDT(dt_sheet);
                                    break;
                                case 1:
                                    grdExcelData2.DataSource = dt_sheet;
                                    grdExcelData2.DataBind();
                                    ViewState["Excel Data 2"] = this.getFilteredDT(dt_sheet);
                                    break;
                                case 2:
                                    grdExcelData3.DataSource = dt_sheet;
                                    grdExcelData3.DataBind();
                                    ViewState["Excel Data 3"] = this.getFilteredDT(dt_sheet);

                                    break;

                            }
                            i++;

                        }
                    }
                }


                btnSubmit.Visible = true;
                connExcel.Close();
            }
            catch (Exception ex)
            {
                lblErrorMsg.Text = ex.Message;
            }
        }


        private List<List<TDTO>> GetTromDT(DataTable dt)
        {
            List<List<TDTO>> dataList = new List<List<TDTO>>();

            foreach (DataRow row in dt.Rows)
            {
                               
		TDTO data = new TDTO();

                data.Period = (!string.IsNullOrEmpty(row[0].ToString())) ? new DateTime(Convert.ToDateTime(row[0]).Date.Year, Convert.ToDateTime(row[0]).Date.Month, 1) : (DateTime?)null;
                data.Division = (!string.IsNullOrEmpty(row[1].ToString())) ? row[1].ToString() : null;
                data.Country = (!string.IsNullOrEmpty(row[2].ToString())) ? row[2].ToString() : null;
                data.Region = (!string.IsNullOrEmpty(row[3].ToString())) ? row[3].ToString() : null;
                data.Site = (!string.IsNullOrEmpty(row[4].ToString())) ? row[4].ToString() : null;
                data.Numbers = Convert.ToDecimal(row[5]);
                List<TDTO> data2 = new List<TDTO>() { data };
                dataList.Add(data2);

            }

             return dataList;
        }

        private List<List<NTO>> GetromDT(DataTable dt)
        {
            List<List<NTO>> dataList = new List<List<NTO>>();

            foreach (DataRow row in dt.Rows)
            {
 
  		 NDTO data = new NDTO();

                data.Period = (!string.IsNullOrEmpty(row[0].ToString())) ? new DateTime(Convert.ToDateTime(row[0]).Date.Year, Convert.ToDateTime(row[0]).Date.Month, 1) : (DateTime?)null;
                data.Group = row[2].ToString();
                

                List<NDTO> data2 = new List<NDTO>() { data };
                dataList.Add(data2);

            }

        
            return dataList;
        }

        private List<List<HDTO>> GetHFromDT(DataTable dt)
        {
            List<List<HDTO>> dataList = new List<List<HDTO>>();

            foreach (DataRow row in dt.Rows)
            {
                 HDTO data = new HDTO();

                data.Period = (!string.IsNullOrEmpty(row[0].ToString())) ? new DateTime(Convert.ToDateTime(row[0]).Date.Year, Convert.ToDateTime(row[0]).Date.Month, 1) : (DateTime?)null;
                data.ANumber = row[1].ToString();
              
                List<HDTO> data2 = new List<HDTO>() { data };
                dataList.Add(data2);

            }

             return dataList;
        }


        private void SubmitExcelData(DataTable dt)
        {
            esProviderFactory.Factory = new EntitySpaces.LoaderMT.esDataProviderFactory();
            lblReport.Text = dt.TableName;
            switch (dt.TableName)
            {

                case "'T Data$'":
                    List<List<TDTO>> dataT = GetTFromDT(dt);

                    lblReport.Text = "There is some error !!";
                    foreach (List<TDTO> data in dataT)
                    {
                        lblReport.Text = "There is some error !!";
                        if (DAL.TRDA.GetInstance().InsUpdTdata))
                        {
                            lblReport.Text = "There is some error inserting !!";
                        }
                        else
                        {

                            lblReport.Text = "There is some error !!";
                        }
                    }
                    lblReport.Text = "All Records Submitted successfully !!";
                    break;

                case "'N Data$'":
                    List<List<NDTO>> dataN = GetNromDT(dt);

                    lblReport.Text = "There is some error 2!!";
                    foreach (List<NDTO> data in dataN)
                    {
                        lblReport.Text = "There is some error 3 !!";
                        if (DAL.MRDA.GetInstance().InsUpdNdata))
                        {
                            lblReport.Text = "There is some error 4 !!";
                        }
                        else
                        {

                            lblReport.Text = "There is some error !!";
                        }
                    }
                    lblReport.Text = "All Records Submitted successfully !!";
                    break;

                case "'H Data$'":
                    List<List<HDTO>> dataH = GetHromDT(dt);

                    lblReport.Text = "There is some error 2!!";
                    foreach (List<HDTO> data in dataH
                    {
                        lblReport.Text = "There is some error 3 !!";
                        if (DAL.HRDA.GetInstance().InsUpdH(data))
                        {
                            lblReport.Text = "There is some error 4 !!";
                        }
                        else
                        {

                            lblReport.Text = "There is some error !!";
                        }
                    }
                    lblReport.Text = "All Records Submitted successfully !!";
                    break;
              
            }



            grdExcelData.DataSource = null;
            grdExcelData.DataBind();

            grdExcelData2.DataSource = null;
            grdExcelData2.DataBind();

            grdExcelData3.DataSource = null;
            grdExcelData3.DataBind();

            uploadRow.Visible = true;
            btnSubmit.Visible = false;
            //txtDate.Enabled = true;

        }


        private void ShowExcelData(DataTable dt)
        {
            grdExcelData.DataSource = dt;
            grdExcelData.DataBind();
            ViewState["Excel Data"] = dt;
        }

        private DataTable getSheetData(string strConn, string sheet)
        {
            string query = "select * from [" + sheet + "]";
            OleDbConnection objConn;
            OleDbDataAdapter oleDA;
            DataTable dt = new DataTable();
            objConn = new OleDbConnection(strConn);
            objConn.Open();
            oleDA = new OleDbDataAdapter(query, objConn);
            oleDA.Fill(dt);
            objConn.Close();
            oleDA.Dispose();
            objConn.Dispose();
            dt.TableName = sheet;
            return dt;
        }

        private DataTable getFilteredDT(DataTable dt)
        {
            DataTable temoDT = new DataTable(dt.TableName);

            switch (dt.TableName.ToUpper().Trim())
            {
                default:
                    return dt;
                          
            }

        }


    }
}
