using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OfficeOpenXml;
using System.IO;


public partial class PreNPI_Fun_NPIImport : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        if (IsPostBack)
        {            

        }
    }
    public void ImportToDataTable()
    {
        string fileName = "";
        lblMsg.Text = "";
        try
        {
            fileName = Path.GetFileName(FileUploadToServer.PostedFile.FileName);
            string filePath = Server.MapPath("~\\FileUpload\\") + Path.GetFileName(FileUploadToServer.PostedFile.FileName);
            FileUploadToServer.SaveAs(filePath);
            DataTable dt = new DataTable();           
            //---------------取出file type : DIF/D_R/DRC,取檔名的前三字元
            if (CheckExcelFile(fileName))         //判斷File是否為xlsx
            {
                string strType = fileName.Substring(0, 3);
                string strSQL = "";
                switch (strType)
                {
                    case "DIF":
                        strSQL = "select Mtype,Mname,Mpos,MRow,MCol from npimap where mtype = 'DIF' order by mid";
                        break;
                    case "Q_R":
                        strSQL = "select Mtype,Mname,Mpos,MRow,MCol from npimap where mtype = 'Q_R' order by mid";
                        break;
                    case "DRC":
                        strSQL = "select Mtype,Mname,Mpos,MRow,MCol from npimap where mtype = 'DRC' order by mid";
                        break;
                    default:
                        lblMsg.ForeColor = System.Drawing.Color.Red;
                        lblMsg.Text = "您選擇的檔案:["+ fileName +"]無法匯入,請重新選擇正確的Excel檔案!!!";
                        strSQL = "";
                        break;
                }            
            //--------------------------------------------------------------------------------
            var existingFile = new FileInfo(filePath);
            
                using (var package = new ExcelPackage(existingFile))
                {
                    ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null && strSQL != "")
                    {
                        if (workBook.Worksheets.Count > 0)
                        {
                            ExcelWorksheet worksheet = workBook.Worksheets.First();//Excel File的第一個Sheet1                           
                            DataTable dt_Pos = new DataTable();
                            clsMySQL db = new clsMySQL();
                            ArrayList arSQL = new ArrayList();
                            db.dbConn();

                            dt_Pos = db.dbQueryDT(strSQL);//db.QueryDataTable(strSQL);
                            db.dbClose();

                            dt.Columns.Add("檔案類型");
                            dt.Columns.Add("欄位名稱");
                            dt.Columns.Add("欄位位置");
                            //dt.Columns.Add("Rows");
                            //dt.Columns.Add("Cols");
                            dt.Columns.Add("欄位內容");
                            //Get Customer and Device value----------------------------------------------------------------
                            string tmpCustomer = "";
                            string tmpDevice = "";
                            for (int t = 0; t <= 2; t++)
                            {
                                if (dt_Pos.Rows[t][1].ToString() == "Customer")
                                {
                                    if (strType == "DRC")
                                    {
                                        string[] str2 = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[t][3].ToString()), Convert.ToInt32(dt_Pos.Rows[t][4].ToString())].Value.ToString().Split(':');
                                        tmpCustomer = str2[1];
                                    }
                                    else
                                    {
                                        tmpCustomer = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[t][3].ToString()), Convert.ToInt32(dt_Pos.Rows[t][4].ToString())].Value.ToString();
                                    }
                                }
                                else
                                {
                                    if (dt_Pos.Rows[t][1].ToString() == "Device")
                                    {
                                        if (strType == "DRC")
                                        {
                                            string[] str2 = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[t][3].ToString()), Convert.ToInt32(dt_Pos.Rows[t][4].ToString())].Value.ToString().Split(':');
                                            tmpDevice = str2[1];
                                        }
                                        else
                                        {
                                            tmpDevice = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[t][3].ToString()), Convert.ToInt32(dt_Pos.Rows[t][4].ToString())].Value.ToString();
                                        }
                                    }
                                }
                            }
                            //---------------------------------------------------------------------------------------------------------
                            for (int i = 0; i < dt_Pos.Rows.Count; i++)
                            {
                                DataRow dr = dt.NewRow();
                                string tmpValue = "";
                                int x = 0;
                                db.dbConn();
                                dr[x++] = dt_Pos.Rows[i][0].ToString();
                                dr[x++] = dt_Pos.Rows[i][1].ToString();
                                dr[x++] = dt_Pos.Rows[i][2].ToString();
                                //dr[x++] = dt_Pos.Rows[i][3].ToString();
                                //dr[x++] = dt_Pos.Rows[i][4].ToString();
                                if (strType == "DRC" && (dt_Pos.Rows[i][2].ToString() == "B46" || dt_Pos.Rows[i][2].ToString() == "B47" || dt_Pos.Rows[i][2].ToString() == "B48"))
                                {
                                    string[] str1 = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[i][3].ToString()), Convert.ToInt32(dt_Pos.Rows[i][4].ToString())].Value.ToString().Split(':');
                                    dr[x++] = str1[1];
                                    tmpValue = str1[1];
                                }
                                else
                                {
                                    dr[x++] = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[i][3].ToString()), Convert.ToInt32(dt_Pos.Rows[i][4].ToString())].Value;

                                    if (worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[i][3].ToString()), Convert.ToInt32(dt_Pos.Rows[i][4].ToString())].Value == null)
                                    {
                                        tmpValue = "";
                                    }
                                    else
                                    {
                                        tmpValue = worksheet.Cells[Convert.ToInt32(dt_Pos.Rows[i][3].ToString()), Convert.ToInt32(dt_Pos.Rows[i][4].ToString())].Value.ToString();
                                    }
                                }
                                dt.Rows.Add(dr);
                                gvRecord.DataSource = dt;
                                gvRecord.DataBind();

                                //-----------------------------------insert MySql Database: npiImportData                                   
                                string strSQL_Query = string.Format("select * from npiImportData where New_Customer = '{0}' and New_Device ='{1}' and Stype='{2}' and Im_Pos='{3}'", tmpCustomer.Trim(), tmpDevice.Trim(), strType.Trim(), dt_Pos.Rows[i][2].ToString().Trim());
                                string strSQL_Del = string.Format("Delete from npiImportData where New_Customer='{0}' and New_Device='{1}' and Stype='{2}' and Im_Pos='{3}'", tmpCustomer.Trim(), tmpDevice.Trim(), strType.Trim(), dt_Pos.Rows[i][2].ToString().Trim());
                                arSQL.Add(strSQL_Del);
                                string strSQL_Insert = string.Format("Insert into npiImportData (New_Customer,New_Device,Stype,UpdateTime,npiUser,Im_Name,Im_Pos,Im_Value) values (" +
                                                       "'{0}','{1}','{2}',NOW(),'','{3}','{4}','{5}')", tmpCustomer.Trim(), tmpDevice.Trim(), strType.Trim(), dt_Pos.Rows[i][1].ToString().Trim(), dt_Pos.Rows[i][2].ToString().Trim(), tmpValue.Trim());
                                arSQL.Add(strSQL_Insert);
                                lblMsg.ForeColor = System.Drawing.Color.Red;
                                //if (db.QueryDataReader(strSQL_Query).HasRows)
                                if (db.dbQueryDR(strSQL_Query).HasRows)
                                {
                                    if (!db.myBatchNonQuery(arSQL))
                                    {
                                        lblMsg.Text = "[Import Error Message] Delete/Insert Fail!!<br/>";
                                    }
                                }
                                else
                                {
                                    if (!db.QueryExecuteNonQuery(strSQL_Insert))
                                    {
                                        lblMsg.Text = "[Import Error Message] Insert Fail!! <br/>";
                                    }
                                }
                                arSQL.Clear();
                                lblMsg.Text = strSQL_Insert + "<br/>" + lblMsg.Text;
                                //---------------------------------------
                                db.dbClose();
                            }
                            lblMsg.ForeColor = System.Drawing.Color.Green;
                            lblMsg.Text = "[" + fileName + "],完成資料匯入!!";
                        }
                        else 
                        {
                            lblMsg.ForeColor = System.Drawing.Color.Red;
                            lblMsg.Text = "您選擇的[" + fileName + "]無法匯入,請重新選擇Excel檔案或檢查Excel檔案內容!!";
                        }
                    }
                    else
                    {
                        gvRecord.DataSource = dt;
                        gvRecord.DataBind();
                    }                    
                }
            }
            File.Delete(filePath); //將IIS上的excel file刪除,C:\inetpub\wwwroot\BU3Web\FileUpload           
        }
        catch (Exception exfile)
        {
            lblMsg.ForeColor = System.Drawing.Color.Red;
            if (fileName == "")
            {
                lblMsg.Text = "[Import Error Message] 請選擇要匯入的Excel檔案!!";
            }
            else {
                lblMsg.Text = "[Import Error Message]您選擇的[" +fileName +"]無法匯入,請重新選擇Excel檔案或檢查Excel檔案內容!!";
            }            
        }
    }

    public Boolean CheckExcelFile(string filename)
    {
        string[] allowdFile = { ".xlsx" };
        bool isValidFile = allowdFile.Contains(System.IO.Path.GetExtension(filename));        
        if (!isValidFile)
        {
            lblMsg.ForeColor = System.Drawing.Color.Red;
            lblMsg.Text = "[Import Error Message] 您選擇檔案:" + filename + ",並不是.xlsx的檔案類型!!!<br/>請重新選擇正確檔案類型.";
        }
        return isValidFile;

    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        ImportToDataTable();
    }
    protected void OpenQueryImport_Click(object sender, EventArgs e)
    {
        string strScript = string.Format("<script language='javascript'>OpenQuery();</script>");
        Page.ClientScript.RegisterStartupScript(this.GetType(), "onload", strScript);
    }
}