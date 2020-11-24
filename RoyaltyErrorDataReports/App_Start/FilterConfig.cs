using System.Web;
using System.Web.Mvc;

namespace RoyaltyErrorDataReports
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
