using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(KLST.TataCommunications.Email.Startup))]
namespace KLST.TataCommunications.Email
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
