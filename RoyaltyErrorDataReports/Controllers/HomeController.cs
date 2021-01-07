using ClosedXML.Excel;
using RoyaltyErrorDataReports.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;

namespace RoyaltyErrorDataReports.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(string CompanyCode)
        {
            ViewBag.Company = CompanyCode;

            List<SqlParameter> lstParam3 = new List<SqlParameter>();
            lstParam3.Add(new SqlParameter("Company", CompanyCode));

            DataTable dtSubDetail = GNF.ExceuteStoredProcedure("SP_Validate_Roy_FireOffStart", lstParam3);
            lstParam3 = new List<SqlParameter>();
            lstParam3.Add(new SqlParameter("Company", CompanyCode));
            DataTable dtSubDetail2 = GNF.ExceuteStoredProcedure("SP_Validate_Roy_Result_Error", lstParam3);
            DataTable dtSubDetail3 = GNF.ExceuteStoredProcedure("[SP_Validate_Roy_Data]");
            //ViewBag.subDetail = dtSubDetail;// ConvertDataTableToHTML(dtSubDetail);
            using (XLWorkbook wb = new XLWorkbook())
            {
                dtSubDetail3.TableName = "Data1";
                wb.Worksheets.Add(dtSubDetail3);
                //wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                //wb.Style.Font.Bold = true;
                //wb.find.Style.Fill.BackgroundColor=XLColor.LightGreen;
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
                }
            }
            return View(dtSubDetail3);
        }

        [HttpPost]
        public ActionResult Export(string CompanyCode)
        {
            List<SqlParameter> lstParam3 = new List<SqlParameter>();
            lstParam3.Add(new SqlParameter("Company", CompanyCode));

            DataTable dtSubDetail = GNF.ExceuteStoredProcedure("SP_Validate_Roy_Data");
            //ViewBag.subDetail = dtSubDetail;// ConvertDataTableToHTML(dtSubDetail);
            using (XLWorkbook wb = new XLWorkbook())
            {
                dtSubDetail.TableName = "Data1";
                wb.Worksheets.Add(dtSubDetail);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Grid.xlsx");
                }
            }
            return RedirectToAction("Index");
        }

        public ActionResult MailOut()
        {

            return View();
        }

        [HttpPost]
        public ActionResult MailOut(string Company)
        {

            string FilePath = "~/temp/";
            DirectoryInfo di = new DirectoryInfo(Server.MapPath(FilePath));
            if (!di.Exists)
            {
                di.Create();
            }
            DataTable dtCompanies = GNF.ExceuteStoredProcedure("SP_RoyaltyErrorDataReports_Companies");
            if (dtCompanies != null && dtCompanies.Rows.Count > 0)
            {
                foreach (DataRow dr in dtCompanies.Rows)
                {
                    List<SqlParameter> lstParam3 = new List<SqlParameter>();
                    lstParam3.Add(new SqlParameter("Company", dr["Company_Code"]));

                    DataTable dtSubDetail = GNF.ExceuteStoredProcedure("SP_Validate_Roy_FireOffStart", lstParam3);
                    if (dtSubDetail != null && dtSubDetail.Rows.Count > 0)
                    {
                        lstParam3 = new List<SqlParameter>();
                        lstParam3.Add(new SqlParameter("Company", dr["Company_Code"]));
                        DataTable dtSubDetail2 = GNF.ExceuteStoredProcedure("SP_Validate_Roy_Result_Error", lstParam3);

                        DataTable dtSubDetail3 = GNF.ExceuteStoredProcedure("[SP_Validate_Roy_Data]");
                        //ViewBag.subDetail = dtSubDetail;// ConvertDataTableToHTML(dtSubDetail);
                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            dtSubDetail3.TableName = "Data1";
                            wb.Worksheets.Add(dtSubDetail3);
                            //wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            //wb.Style.Font.Bold = true;
                            //wb.find.Style.Fill.BackgroundColor=XLColor.LightGreen;
                            string FileName = "File-" + DateTime.Now.ToString("MMddyyyyHHmmsstt") + ".xlsx";
                            string path = FilePath + FileName;
                            wb.SaveAs(path);

                            lstParam3 = new List<SqlParameter>();
                            lstParam3.Add(new SqlParameter("CompanyCode", dr["Company_Code"]));

                            DataTable dtSubDetail4 = GNF.ExceuteStoredProcedure("SP_RoyaltyErrorDataReports_Email", lstParam3);
                            if (dtSubDetail4 != null && dtSubDetail4.Rows.Count > 0)
                            {
                                foreach (DataRow dr4 in dtSubDetail4.Rows)
                                {
                                    List<Attachment> lstAttachment = new List<Attachment>();
                                    lstAttachment.Add(new Attachment(path));
                                    string Subject = "Please be advised that the attached file contains errors on item setup. This is for Company '"+ dr["Company_Code"] .ToString()+ "'";
                                    string Body = "Please fix these errors within (3) business days upon receipt of this email. Please reach out to Eli Maiman with any questions or concerns";
                                    NotificationHelper.SendMail(dr4["Email"].ToString(), Subject, Body, true, lstAttachment);
                                }
                            }

                        }
                    }
                }
            }

            return View();
        }


        public static string ConvertDataTableToHTML(DataTable dt)
        {
            string html = "<table class='' border='1'>";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
            {

                html += "<th>" + dt.Columns[i].ColumnName + "</th>";
            }
            html += "</tr>";

            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString() == "FAIL")
                    {
                        html += "<td style=' padding: 4px; background: red'>" + dt.Rows[i][j].ToString() + "</td>";

                    }
                    else
                    {
                        html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                    }
                }
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }
        public static string ConvertDataRowToHTML(DataRow dr)
        {
            DataTable dt = dr.Table;
            string html = "<table class='' border='1'>";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
            {

                html += "<th>" + dt.Columns[i].ColumnName + "</th>";
            }
            html += "</tr>";

            //add rows

            html += "<tr>";
            for (int j = 0; j < dt.Columns.Count; j++)
            {

                html += "<th>" + dr[j].ToString() + "</th>";

            }
            html += "</tr>";

            html += "</table>";
            return html;
        }
    }
}