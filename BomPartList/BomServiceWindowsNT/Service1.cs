using System;
using System.Diagnostics;
using System.ServiceProcess;
using GetBomService.Service;
using System.ServiceModel;
using System.ServiceModel.Description;


namespace BomServiceWindowsNT
{
    public partial class Service1 : ServiceBase
    {
        // ServiceHost _service;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            AddLog("start");

            var baseAddress = new Uri("http://192.168.14.86/bomService");  //http://192.168.14.11:8085/bomtable   http://192.168.14.86:8085/bomService   http://srvkb:8085/bomtable

            using (var host = new ServiceHost(typeof(BomServiceClass.BomTableService), baseAddress))
            {
                //   Enable metadata publishing.
                var smb = new ServiceMetadataBehavior
                {
                    HttpGetEnabled = true,
                    MetadataExporter = { PolicyVersion = PolicyVersion.Policy15 },

                    //HttpGetBinding = new BasicHttpBinding
                    //{
                    //    MaxBufferSize = 2147483647,
                    //    MaxReceivedMessageSize = 2147483647
                    //}
                };
                
                host.Description.Behaviors.Add(smb);

                // Open the ServiceHost to start listening for messages. Since
                // no endpoints are explicitly configured, the runtime will create
                // one endpoint per base address for each service contract implemented
                // by the service.
                host.Open();

                AddLog(String.Format("The service is ready at {0}", baseAddress));

                // Close the ServiceHost.
                host.Close();
            }
        }

        protected override void OnStop()
        {
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
            catch { eventLog1.WriteEntry(log); }
        }
    }
}
