using RoyaltyErrorDataReports.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
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
            List<SqlParameter> lstParam3 = new List<SqlParameter>();
            lstParam3.Add(new SqlParameter("Company", CompanyCode));

            DataTable dtSubDetail = GNF.ExceuteStoredProcedure("SP_Validate_Roy_FireOffStart", lstParam3);
            ViewBag.subDetail = ConvertDataTableToHTML(dtSubDetail);

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

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