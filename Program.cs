using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Management;
using System.Configuration;


namespace CheckServiceStatus
{
    class Program
    {
        static void Main(string[] args)
        {
            ServerUtility objUtil = new ServerUtility();
            objUtil.GetServiceStatusFromServers();
            //Console.ReadLine();
        }
    }
}
