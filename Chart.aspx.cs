using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;
using System.Text;

public partial class Chart : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void ClearUpload_Click(object sender, EventArgs e)
    {
        this.Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Javascript", "<script type='text/javascript'>Redraw();</script>");
    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {
        if (FileUpload1.HasFile)
        {
            try
            {
                String FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                String FileExtension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                String FileLocation = Server.MapPath("~/Excel/" + FileName);

                if (FileExtension == ".xls" || FileExtension == ".xlsx")
                {
                    FileUpload1.SaveAs(FileLocation);

                    DataTable dt = ExtractDataFromExcel(FileLocation, FileExtension);

                    StringBuilder str = ManipulateDataTable(dt);

                    this.Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Javascript", "<script type='text/javascript'>testFunction(" + str + ");</script>");
                }
                else
                {
                    this.Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Javascript", "<script type='text/javascript'>alert('File format not correct. use .xls or .xlsx file only.');</script>");
                }
            }
            catch (Exception ex)
            {
                this.Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Javascript", "<script type='text/javascript'>alert('Exception Occur' "+ Convert.ToString(ex.InnerException) +");</script>");
            }
        }
    }

    protected void DownloadFile_Click(object sender, EventArgs e)
    {
        Response.ContentType = "application/vnd.ms-excel";
        Response.AppendHeader("Content-Disposition", "attachment; filename=SampleExcel.xlsx");
        Response.TransmitFile(Server.MapPath("~/SampleExcel.xlsx"));
        Response.End();
    }

    public static DataTable ExtractDataFromExcel(String Location, String Extension)
    {
        String ConnectionStr = String.Empty;
        DataTable dt = null;
        if (Extension == ".xls")
        {
            ConnectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Location + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
        }
        else if (Extension == ".xlsx")
        {
            ConnectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Location + ";Extended Properties=\"Excel 12.0;IMEX=0;HDR=YES;TypeGuessRows=0;ImportMixedTypes=Text\"";
        }
        using (OleDbConnection con = new OleDbConnection(ConnectionStr))
        {
            con.Open();
            using (OleDbDataAdapter olda = new OleDbDataAdapter())
            {
                DataTable dtExcelSheetName = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                String ExcelSheetName = dtExcelSheetName.Rows[0]["TABLE_NAME"].ToString();

                string Query = "Select * from [" + ExcelSheetName + "]";

                OleDbCommand cmd = new OleDbCommand(Query, con);
                dt = new DataTable();
                olda.SelectCommand = cmd;
                olda.Fill(dt);
            }
            con.Close();
        }
        return dt;
    }

    static StringBuilder ManipulateDataTable(DataTable dt)
    {
        StringBuilder List = new StringBuilder();
        List.Append("[");
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            if (i == 0)
            {
                List.Append("[");
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0)
                    {
                        List.Append("'" + dt.Rows[i][dt.Columns[j].ColumnName] + "'");
                    }
                    else
                    {
                        List.Append("," + dt.Rows[i][dt.Columns[j].ColumnName]);
                    }
                }
                List.Append("]");
            }
            else
            {
                List.Append(",[");
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0)
                    {
                        List.Append("'" + dt.Rows[i][dt.Columns[j].ColumnName] + "'");
                    }
                    else
                    {
                        List.Append("," + dt.Rows[i][dt.Columns[j].ColumnName]);
                    }
                }
                List.Append("]");
            }
        }
        List.Append("]");
        return List;
    }
}