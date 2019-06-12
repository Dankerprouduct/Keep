using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Keep
{
    class Program
    {
        static void Main(string[] args)
        {

            while (true)
            {
                ListPrinters();

                Console.WriteLine("\n");
                Thread.Sleep(1000);
            }
        }

        static void ListPrinters()
        {
            var printerQuery = new ManagementObjectSearcher("SELECT * from Win32_Printer");
            foreach (ManagementObject printer in printerQuery.Get())
            {
                var name = printer.GetPropertyValue("Name");
                var status = printer["WorkOffline"];
                var isDefault = printer.GetPropertyValue("Default");
                var isNetworkPrinter = printer.GetPropertyValue("Network");

                if (bool.Parse(status.ToString()) == true && printer.ToString().ToLower().Contains("dymo"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("printer disconnected. Removing");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    printer.Delete();

                    RestartHardwareService();
                    // rerfresh raptor hardware service
                }

                Console.WriteLine("{0} (Status: {1}, Default: {2}, Network: {3}",
                    name, status, isDefault, isNetworkPrinter);
            }
        }

        static void RestartHardwareService()
        {
            ServiceController service = new ServiceController("Raptor Hardware Service");
            try
            {
                Console.WriteLine("Stopping Hardware Service");
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped);
                Console.WriteLine("Starting Hardware Service");
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running);

                Console.WriteLine("Hardware Service Restarted!");
            }
            catch (Exception ex)
            {
                // ...

                Console.WriteLine(ex.ToString());
            }
        }
    }
}
