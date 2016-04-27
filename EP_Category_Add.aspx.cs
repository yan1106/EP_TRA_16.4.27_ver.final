using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using NPOI.XSSF.UserModel;



public partial class EP_Category_Add : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }


   

    protected void insert_cate(int item_num, List<string> Cate_Iiitems, List<string> Cate_SpeChar, List<string> Cate_Md, List<string> Cate_Cate, List<string> Cate_KP, string name,string status,string stage)
    {
        clsMySQL db = new clsMySQL();
        List<int> success_insert = new List<int>();
        List<int> fail_insert = new List<int>();
        int success_count = 0;
        int fail_count = 0;
        List<string> history_cate = new List<string>();

        for (int i = 0; i < item_num; i++)
        {

            String insert_cate = string.Format("insert into npi_cap_ep" +
                                           "(npi_EP_Cate_Username,npi_EP_Cate_UpdateTime,npi_EP_Cate_Status," +
                                         "EP_Cate_Stage,EP_Cate_Iiitems,EP_Cate_SpeChar," +
                                         "EP_Cate_Md,EP_Cate_Cate,EP_Cate_KP)values" +
                                         "('{0}',NOW(),'{1}'," +
                                          "'{2}','{3}','{4}','{5}','{6}','{7}')"
                                          , name, status, stage, Cate_Iiitems[i], Cate_SpeChar[i], Cate_Md[i], Cate_Cate[i], Cate_KP[i]);
            if (db.QueryExecuteNonQuery(insert_cate) == true)
            {
                success_count++;
            }
            else
            {
                fail_count++;
                history_cate.Add(Cate_Iiitems[i] + "|" + Cate_SpeChar[i] + "|" + Cate_Md[i] + "|" + Cate_Cate[i] + "|" + Cate_KP[i]);
            }
        }





        /*
        try
        {
            /*if (Text_Packge_insert.Text.Trim() == "")
            {
                string strScript = string.Format("<script language='javascript'>alert('您沒有輸入Packge_Name!');</script>");
                Page.ClientScript.RegisterStartupScript(this.GetType(), "onload", strScript);
            }
            if ()
            {
                string strScript = string.Format("<script language='javascript'>alert('Category新增成功');</script>");
                Page.ClientScript.RegisterStartupScript(this.GetType(), "onload", strScript);
                
            }
            else
            {
                string strScript = string.Format("<script language='javascript'>alert('Category新增成功');</script>");
                Page.ClientScript.RegisterStartupScript(this.GetType(), "onload", strScript);
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }

        */





    }


    protected void btnUpload_Click(object sender, EventArgs e)
    {
        String SavePath = "D:\\brunohuang\\FileUpload_Folder\\";
        string sheet_name = "";
        int sheet_num;
        int dLastNum;
        int cate_items = 0;
        String Cate_Username = "CIM";
        String Cate_Status = "Y";
        /*List<string> Cate_Iiitems = new List<string>();
        List<string> Cate_SpeChar = new List<string>();
        List<string> Cate_Md = new List<string>();
        List<string> Cate_Cate = new List<string>();
        List<string> Cate_KP = new List<string>();
        */
        List<int> success_insert = new List<int>();
        List<int> fail_insert = new List<int>();
        int success_count = 0;
        int fail_count = 0;
        List<string> history_cate = new List<string>();
        string fileName = "";
       clsMySQL db = new clsMySQL();

        try {

            

            fileName = Path.GetFileName(FileUploadToServer.PostedFile.FileName);
            string filePath = Server.MapPath("D:\\brunohuang\\FileUpload_Folder\\") + Path.GetFileName(FileUploadToServer.PostedFile.FileName);
            FileUploadToServer.SaveAs(filePath);


            lblMsg.Text = "上傳成功!!" + SavePath;


            if (!CheckExcelFile(fileName)) { 

            if (FileUploadToServer.HasFile)
            {
               
            }


            XSSFWorkbook wk = new XSSFWorkbook(FileUploadToServer.FileContent);
            XSSFSheet hst;
            XSSFRow hr;


            sheet_num = wk.NumberOfSheets;

            for (int k = 0; k < sheet_num; k++)
            {
                hst = (XSSFSheet)wk.GetSheetAt(k);
                cate_items = hst.LastRowNum; //每一張工作表有幾筆資料

                sheet_name = hst.SheetName;

                hr = (XSSFRow)hst.GetRow(0);
                dLastNum = hr.LastCellNum; //每一列的欄位數



                for (int j = 1; j <= cate_items; j++)
                {
                    hr = (XSSFRow)hst.GetRow(j);

                    /*for(int i=0;i<dLastNum;i++)
                    {
                        string strcell = hr.GetCell(i) == null ? "0" : hr.GetCell(i).ToString();

                    }*/

                    String insert_cate = string.Format("insert into npi_ep_category" +
                                               "(npi_EP_Cate_Username,npi_EP_Cate_UpdateTime,npi_EP_Cate_Status," +
                                             "EP_Cate_Stage,EP_Cate_Iiitems,EP_Cate_SpeChar," +
                                             "EP_Cate_Md,EP_Cate_Cate,EP_Cate_KP)values" +
                                             "('{0}',NOW(),'{1}'," +
                                              "'{2}','{3}','{4}','{5}','{6}','{7}')"
                                              , Cate_Username, Cate_Status, sheet_name, hr.GetCell(0), hr.GetCell(1), hr.GetCell(2), hr.GetCell(3), hr.GetCell(4));
                    if (db.QueryExecuteNonQuery(insert_cate) == true)
                    {
                        success_count++;
                    }
                    else
                    {
                        fail_count++;
                        history_cate.Add(hr.GetCell(0) + "|" + hr.GetCell(1) + "|" + hr.GetCell(2) + "|" + hr.GetCell(3) + "|" + hr.GetCell(4));
                    }


                }
            }
        }
        else{
                lblMsg.ForeColor = System.Drawing.Color.Red;
                lblMsg.Text = "您選擇的[" + fileName + "]無法匯入,請重新選擇Excel檔案或檢查Excel檔案內容!!";
            }

        }
        catch (Exception exfile)
        {
            lblMsg.ForeColor = System.Drawing.Color.Red;
            if (fileName == "")
            {
                lblMsg.Text = "[Import Error Message] 請選擇要匯入的Excel檔案!!";
            }
            else {
                lblMsg.Text = "[Import Error Message]您選擇的[" + fileName + "]無法匯入,請重新選擇Excel檔案或檢查Excel檔案內容!!";
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



}