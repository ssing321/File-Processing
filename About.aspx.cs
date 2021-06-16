using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.IO;
using System.Data.Common;

namespace ProjExcel
{
    public partial class About : Page
    {
        public DataSet ds_glb = new DataSet(); 
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            string path = Path.GetFileName(FileUpload1.FileName);
            path = path.Replace(" ", "");
            FileUpload1.SaveAs(Server.MapPath("~/Upld_files/") + path);
            String ExcelPath = Server.MapPath("~/Upld_files/") + path;
            OleDbConnection mycon = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties=Excel 8.0; Persist Security Info = False");
            mycon.Open();
            OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$]", mycon);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds);
            ds_glb = ds;
            GridView1.DataSource = ds.Tables[0];
            GridView1.DataBind();
            mycon.Close();
            Label3.Text = "Excel File Has Been Saved and Data Captured";

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = "Data Source=LAPTOP-4FBNESCC\\SQLEXPRESS;Initial Catalog=ExcelDatabase;Integrated Security=True";
            cn.Open();

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                int val1 = Convert.ToInt32(ds.Tables[0].Rows[i]["Rno"].ToString());
                String val2 = ds.Tables[0].Rows[i]["FirstName"].ToString();
                String val3 = ds.Tables[0].Rows[i]["LastName"].ToString();
                int val4 = Convert.ToInt32(ds.Tables[0].Rows[i]["Marks"].ToString());

                SqlCommand cmd2 = new SqlCommand("insert into StudentDetail(Rno,FirstName,LastName,Marks) values (" + val1 + ",'" + val2 + "','" + val3 + "'," + val4 + ")", cn);
                cmd2.ExecuteNonQuery();
            }
        }

        protected void Button2_Click(object sender, EventArgs e)
        {

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = "Data Source=LAPTOP-4FBNESCC\\SQLEXPRESS;Initial Catalog=ExcelDatabase;Integrated Security=True";
            cn.Open();

            for (int i = 0; i < ds_glb.Tables[0].Rows.Count; i++)
            {
                int val1 = Convert.ToInt32(ds_glb.Tables[0].Rows[i]["Rno"].ToString());
                String val2 = ds_glb.Tables[0].Rows[i]["FirstName"].ToString();
                String val3 = ds_glb.Tables[0].Rows[i]["LastName"].ToString();
                int val4 = Convert.ToInt32(ds_glb.Tables[0].Rows[i]["Marks"].ToString());

                SqlCommand cmd = new SqlCommand("insert into StudentDetail(Rno,FirstName,LastName,Marks) values (" + val1 + ",'" + val2 + "','" + val3 + "'+" + val3 + ")", cn);
                cmd.ExecuteNonQuery();
            }
        }
    }
}

