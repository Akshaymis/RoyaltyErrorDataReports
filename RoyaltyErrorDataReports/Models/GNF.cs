using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace RoyaltyErrorDataReports.Models
{
    public class GNF
    {
        public static DataTable ExceuteQuery(string sp)
        {
            using (var context = new ApplicationDbContext())
            {
                var dt = new DataTable();
                var conn = context.Database.Connection;
                var connectionState = conn.State;
                try
                {
                    if (connectionState != ConnectionState.Open) conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = sp;
                        cmd.CommandType = CommandType.Text;

                        using (var reader = cmd.ExecuteReader())
                        {
                            dt.Load(reader);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // error handling
                    throw;
                }
                finally
                {
                    if (connectionState != ConnectionState.Closed) conn.Close();
                }
                return dt;
            }
        }
        public static DataTable ExceuteStoredProcedure(string sp  , List<SqlParameter> param = null)
        {
            using (var context = new ApplicationDbContext())
            {
                var dt = new DataTable();
                var conn = context.Database.Connection;
                var connectionState = conn.State;
                try
                {
                   if (connectionState != ConnectionState.Open) conn.Open();
                 
                       
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = sp;
                        cmd.CommandType = CommandType.StoredProcedure;
                        if (param != null)
                        {
                            foreach (SqlParameter obj in param)
                            {
                                cmd.Parameters.Add(obj);
                            }
                        }
                        using (var reader = cmd.ExecuteReader())
                        {
                            dt.Load(reader);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // error handling
                    throw;
                }
                finally
                {
                    if (connectionState != ConnectionState.Closed) conn.Close();
                }
                return dt;
            }
        }
        public static string ConvertDataTabletoString(DataTable dt)
        {
            System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;
            foreach (DataRow dr in dt.Rows)
            {
                row = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    row.Add(col.ColumnName, dr[col]);
                }
                rows.Add(row);
            }
            return serializer.Serialize(rows);
        }

    }
}