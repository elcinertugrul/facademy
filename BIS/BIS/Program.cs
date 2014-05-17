using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BIS
{
    class Program
    {
        static void Main(string[] args)
        {
            //http://a810-bisweb.nyc.gov/bisweb/FacadeStatusInformationServlet?allisn=0000024332&allbin=1016880&requestid=2

            string url = "http://a810-bisweb.nyc.gov/bisweb/FacadesByLocationServlet?requestid=1&allbin=1016880";
            //string url = "http://a810-bisweb.nyc.gov/bisweb/FacadesByLocationServlet?requestid=1&allbin=1009012";
            System.Net.WebRequest request = System.Net.WebRequest.Create(url);
            System.Net.WebResponse myresponse = request.GetResponse();
            // Open data stream:
            System.IO.Stream _WebStream = myresponse.GetResponseStream();
            //string path = @"C:\Temp\streetview";
            var result = new System.Net.WebClient().DownloadData(url);

            string astringHTML = new System.Net.WebClient().DownloadString(url);
            
        }
    }
}
