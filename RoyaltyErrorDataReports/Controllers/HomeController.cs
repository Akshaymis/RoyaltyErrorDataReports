﻿using ClosedXML.Excel;
using Newtonsoft.Json;
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
            try
            {
                ViewBag.Message = "";

                string FilePath = "E:\\IIS_Published_Websites\\RoyaltyErrorDataReports_Publish\\temp\\";
                DirectoryInfo di = new DirectoryInfo(FilePath);
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
                        lstParam3 = new List<SqlParameter>();
                        lstParam3.Add(new SqlParameter("Company", dr["Company_Code"]));
                        DataTable dtSubDetail2 = GNF.ExceuteStoredProcedure("SP_Validate_Roy_Result_Error", lstParam3);
                        if (dtSubDetail2 != null && dtSubDetail2.Rows.Count > 0)
                        {


                            DataTable dtSubDetail3 = GNF.ExceuteStoredProcedure("[SP_Validate_Roy_Data]");
                            if (dtSubDetail3 != null && dtSubDetail3.Rows.Count > 0)
                            {
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
                                            FileInfo fi = new FileInfo(path);
                                            string newFileName = (FilePath + (fi.Name.Replace(".xlsx", Guid.NewGuid().ToString() + ".xlsx")));
                                            fi.CopyTo(newFileName);
                                            lstAttachment.Add(new Attachment(newFileName));
                                            string Subject = "Please be advised that the attached file contains errors on item setup. This is for Company '" + dr["Company_Code"].ToString() + "'";
                                            string Body = "Please fix these errors within (3) business days upon receipt of this email. Please reach out to Eli Maiman with any questions or concerns";
                                            NotificationHelper.SendMail(dr4["Email"].ToString(), Subject, Body, true, lstAttachment);
                                        }
                                    }

                                }
                            }
                        }
                        else
                        {
                            ViewBag.Message = "No data to sent in subDetail";
                        }
                    }
                    ViewBag.Message = "MailOut completed successfully";
                }
                else
                {
                    ViewBag.Message = "No data to sent";
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                NotificationHelper.SendMail("amishra@bentex.com", "Mailout >> Exception occured", ex.Message + "<br>" + ex.StackTrace, true, null);
                NotificationHelper.SendMail("emaiman@bentex.com", "Mailout >> Exception occured", ex.Message + "<br>" + ex.StackTrace, true, null);
            }

            return View();
        }


        public ActionResult MailAutomation(string Code)
        {
            bool MailAutomationWorked = false;
            if (Code != "11")
            {
                return View();
            }
            try
            {
                ViewBag.Message = "";

                string FilePath = "temp";// "E:\\IIS_Published_Websites\\RoyaltyErrorDataReports_Publish\\temp\\";
                DirectoryInfo di = new DirectoryInfo(HttpContext.Server.MapPath(FilePath));
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
                        lstParam3 = new List<SqlParameter>();
                        lstParam3.Add(new SqlParameter("Company", dr["Company_Code"]));
                        DataTable dtSubDetail2 = GNF.ExceuteStoredProcedure("SP_Validate_Roy_Result_Error", lstParam3);
                        if (dtSubDetail2 != null && dtSubDetail2.Rows.Count > 0)
                        {


                            DataTable dtSubDetail3 = GNF.ExceuteStoredProcedure("[SP_Validate_Roy_Data]");
                            if (dtSubDetail3 != null && dtSubDetail3.Rows.Count > 0)
                            {
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
                                            FileInfo fi = new FileInfo(path);
                                            string newFileName = (FilePath + (fi.Name.Replace(".xlsx", Guid.NewGuid().ToString() + ".xlsx")));
                                            fi.CopyTo(newFileName);
                                            lstAttachment.Add(new Attachment(newFileName));
                                            string Subject = "Please be advised that the attached file contains errors on item setup. This is for Company '" + dr["Company_Code"].ToString() + "'";
                                            string Body = "Please fix these errors within (3) business days upon receipt of this email. Please reach out to Eli Maiman with any questions or concerns";
                                            NotificationHelper.SendMail(dr4["Email"].ToString(), Subject, Body, true, lstAttachment);
                                        }
                                    }

                                }
                            }
                        }
                        else
                        {
                            MailAutomationWorked = true;
                            ViewBag.Message = "No data to sent in subDetail";
                        }
                    }
                    MailAutomationWorked = true;
                    ViewBag.Message = "MailOut completed successfully";
                }
                else
                {
                    MailAutomationWorked = true;
                    NotificationHelper.SendMail("amishra@bentex.com", "Mailout >> No data to sent", "", true, null);
                    ViewBag.Message = "No data to sent";
                }
            }
            catch (Exception ex)
            {
                MailAutomationWorked = false;
                ViewBag.Message = ex.Message;
                NotificationHelper.SendMail("amishra@bentex.com", "Mailout >> Exception occured", ex.Message + "<br>" + ex.StackTrace, true, null);
                NotificationHelper.SendMail("emaiman@bentex.com", "Mailout >> Exception occured", ex.Message + "<br>" + ex.StackTrace, true, null);
            }
            if (MailAutomationWorked)
            {
                DBHelperCronJObDB objCronDb = new DBHelperCronJObDB();
                string Sql = "insert into Cron_Job_Processed (job_number,job_name) values (\'1004\',\'RoyaltyErrorReports')";
                objCronDb.ExecuteNonQuery(Sql);
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