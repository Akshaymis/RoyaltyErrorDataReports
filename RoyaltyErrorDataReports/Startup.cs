using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(RoyaltyErrorDataReports.Startup))]
namespace RoyaltyErrorDataReports
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
