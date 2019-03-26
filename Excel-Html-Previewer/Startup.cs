using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Excel_Html_Previewer.Startup))]
namespace Excel_Html_Previewer
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {

        }
    }
}
