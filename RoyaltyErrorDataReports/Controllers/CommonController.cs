using RoyaltyErrorDataReports.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RoyaltyErrorDataReports.Controllers
{
    public class CommonController : Controller
    {
        // GET: Common
        public ActionResult Index()
        {
            return View();
        }
        public JsonResult GetCompany()
        {
            try
            {
                DataTable dt = GNF.ExceuteStoredProcedure("SP_SelectCompany");
                return Json(GNF.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(ex.ToString(), JsonRequestBehavior.AllowGet);
            }
        }
    }
}

// Akshay are you on?
////

//Private Sub SubFireOffError(Company As String)


//Dim cmd As ADODB.Command
//Dim PRM As ADODB.Parameter
//Dim adoRecordset As New ADODB.Recordset
//Dim ctr As Integer
//Dim X As Variant

//    Set cmd = New ADODB.Command
//    With cmd
//            .CommandText = "SP_Validate_Roy_FireOffStart_ReportNew_1"
//            .CommandType = adCmdStoredProc
//            .ActiveConnection = MyConnObj
//            .CommandTimeout = 9000


//        End With
//     Set PRM = cmd.CreateParameter(_
//    "@Company", adVarChar, adParamInput, 3)
//    cmd.Parameters.Append PRM
//    PRM.Value = Txt box company code



//    Set adoRecordset = New ADODB.Recordset
//    Set adoRecordset = cmd.Execute



//Exit Sub
//errorHandle:
//MsgBox "SubFireOffError"
//End Sub