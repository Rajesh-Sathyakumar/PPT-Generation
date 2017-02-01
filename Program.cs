
using System.Configuration;

using Topshelf;

namespace AOAService
{
    class Program
    {
        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();
            HostFactory.Run(x =>
            {
                x.Service<DataGenerationScheduler>(s =>
                {
                    s.ConstructUsing(name => new DataGenerationScheduler());
                    s.WhenStarted(tc => tc.Start());
                    s.WhenStopped(tc => tc.Stop());
                });
                x.RunAsLocalSystem();

                x.SetDescription("Generates Excel and Powerpoint Report based on Template files that is required for DA's Opportunity Analysis");
                x.SetDisplayName("AOAv1.0");
                x.SetServiceName("AOAv1.0");
            });
        }
    }
}