using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Web;
using System.Text;
using System.Threading.Tasks;
using OfficeToPdf.Web;

namespace OfficeToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Uri uri = new Uri("http://localhost:8000/converter");
            var serviceImpl = new OfficeToPdf.Web.WebContract();
            using (var host = new WebServiceHost(serviceImpl, uri))
            {
                host.Open();

                Console.WriteLine("Service up and running at:");
                foreach (var ea in host.Description.Endpoints)
                {
                    Console.WriteLine(ea.Address);
                }

                Console.ReadLine();
                host.Close();
            }
        }
    }
}
