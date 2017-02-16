using System;
using System.Diagnostics;
using System.ServiceProcess;
using ACadService.Service;
using System.ServiceModel;
using System.ServiceModel.Description;

namespace ACadService
{
    public partial class ACadService : ServiceBase
    {
        public ACadService()
        {
            InitializeComponent();
            ServiceName = "AirVentsCadCode";
        }

        public static void Main()
        {
            Run(new ACadService());
        }

        public ServiceHost ServiceHost = null;

        protected override void OnStart(string[] args)
        {
            AddLog("start");

            if (ServiceHost != null)
            {
                ServiceHost.Close();
            }


           // var baseAddress = new Uri("http://192.168.14.86/bomService");

            var baseAddress = new Uri("http://192.168.14.86/bomService");
                //http://192.168.14.11:8085/bomtable   http://192.168.14.86:8085/bomService   http://srvkb:8085/bomtable

            using (ServiceHost = new ServiceHost(typeof (BomServiceClass.BomTableService), baseAddress))
            {
                var serviceMetadataBehavior = new ServiceMetadataBehavior
                {
                    HttpGetEnabled = true,
                    MetadataExporter = {PolicyVersion = PolicyVersion.Policy15}
                };

                ServiceHost.Description.Behaviors.Add(serviceMetadataBehavior);

                ServiceHost.Open();

                AddLog(String.Format("The service is ready at {0}", baseAddress));
            }
        }

        protected override void OnStop()
        {
            if (ServiceHost != null)
            {
                ServiceHost.Close();
                ServiceHost = null;
            }
            AddLog("stop");
        }
        
        public void AddLog(string log)
            {
                try
                {
                    if (!EventLog.SourceExists("BomService"))
                    {
                        EventLog.CreateEventSource("BomService", "BomService");
                    }
                    eventLog1.Source = "BomService";
                    eventLog1.WriteEntry(log);
                }
                catch
                {
                    eventLog1.WriteEntry(log);
                }
            }
        
    }
}
