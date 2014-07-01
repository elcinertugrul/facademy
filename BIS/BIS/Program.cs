using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TToffice;
using System.Web.Script.Serialization;
using System.IO;

namespace BIS
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //1. read excel & get the list of all BINs
            List<string> bins = new List<string>();

            TT_Excel xl = new TT_Excel();
            //xl.openInvisible(@"C:\Users\EErtugrul\Documents\GitHub\facademy\ZipCode10280_BINs.xlsx");
            xl.openInvisible(@"C:\Users\ELCIN\Documents\GitHub\facademy\ZipCode10010_BINs.xlsx");
            string[] xldata;
            xl.readColumn(1, 1, out xldata);

            xl.close();
            xl.release();
            xl = null;

            bins = xldata.ToList();

            //2. Loop over BINs & create list of object
            List<facadedata> scrapeddata = new List<facadedata>();

            foreach (var bin in bins)
            {

                facadedata bldgdata = new facadedata();


                //http://a810-bisweb.nyc.gov/bisweb/FacadeStatusInformationServlet?allisn=0000024332&allbin=1016880&requestid=2
                //string url = "http://a810-bisweb.nyc.gov/bisweb/FacadesByLocationServlet?requestid=1&allbin=1016880";
                //string url = "http://a810-bisweb.nyc.gov/bisweb/FacadesByLocationServlet?requestid=1&allbin=1009012";

                string url = "http://a810-bisweb.nyc.gov/bisweb/FacadesByLocationServlet?requestid=1&allbin=" + bin;

                //System.Net.WebRequest request = System.Net.WebRequest.Create(url);
                //System.Net.WebResponse myresponse = request.GetResponse();
                // Open data stream:
                //System.IO.Stream _WebStream = myresponse.GetResponseStream();
                
                //string path = @"C:\Temp\streetview";
                //var result = new System.Net.WebClient().DownloadData(url);
                string astringHTML = string.Empty;
                try
                {
                   astringHTML = new System.Net.WebClient().DownloadString(url);
                }
                catch (Exception)
                {
                    
                   // throw; // do nothing
                }


                if (astringHTML != string.Empty)
                {
                    string[] splitone = { "ErrorMsg :: " };
                    string[] words = astringHTML.Split(splitone, StringSplitOptions.None);

                    if (words[1].StartsWith("\n")) // there is data available keep contunie
                    {
                        string[] splittwo = { "[0:FCycle]{" };
                        string[] wordscycles = astringHTML.Split(splittwo, StringSplitOptions.None);

                        for (int i = 1; i < wordscycles.ToList().Count; i++)
                        {
                            cycle acycle = new cycle();
                            string[] splitthree = { "}\n","[1:FControlNumber]{", "[5:FCurrentStatus]{", "[8:FInitFileDate]{", "[9:FaIsn]{" };
                            string[] splitfour = { "}\n", "[1:FControlNumber]{", "[2:FHouseNumber]{", "[3:FStreetName]{", "[5:FCurrentStatus]{", "[6:FBin]{", "[7:FNumStories]{","[8:FInitFileDate]{", "[9:FaIsn]{" };


                            if (i != 1)
                            {
                                string[] wordsinfo = wordscycles[i].Split(splitthree, StringSplitOptions.None);

                                acycle.FCycle = wordsinfo[0];
                                acycle.FCtrlNum = wordsinfo[2];
                                acycle.FCurStat = wordsinfo[7];
                                acycle.FInitDate = wordsinfo[11];
                                acycle.FaISN = wordsinfo[13];
                            }
                            else 
                            {
                                string[] wordsinfo = wordscycles[i].Split(splitfour, StringSplitOptions.None);

                                acycle.FCycle = wordsinfo[0];
                                acycle.FCtrlNum = wordsinfo[2];
                                bldgdata.Num = wordsinfo[4];
                                bldgdata.Strt = wordsinfo[6];
                                acycle.FCurStat = wordsinfo[9];
                                bldgdata.BIN = wordsinfo[11];
                                bldgdata.NumStory = wordsinfo[13];
                                acycle.FInitDate = wordsinfo[15];
                                acycle.FaISN = wordsinfo[17];
                            
                            }

                            bldgdata.Cycles.Add(acycle);
                        }

                        //populate zipcode and borough

                        string[] splitfive = {"\nNmBoro :: " , "\nVlBin :: " ,"\nVlNumZip :: ", "\nVlTaxBlock :: " };
                        string[] wordsthree = astringHTML.Split(splitfive, StringSplitOptions.None);
                        bldgdata.Boro = wordsthree[1];
                        bldgdata.Zip = wordsthree[3];

                        //string[] splittwo = { "\nVlNumHous :: ", "\nNmStrt :: " };
                        //string[] wordstwo = words[1].Split(splittwo, StringSplitOptions.None);

                        //bldgdata.NumHous = wordstwo[1];

                        //string[] splitthree = { "\nVlNmStrt :: ", "\nNmBoro :: " };
                        //string[] wordsthree = wordstwo[2].Split(splitthree, StringSplitOptions.None);

                        //bldgdata.StrtName = wordsthree[0];

                        //string[] splitfour = { "\nVlBin :: " };
                        //string[] wordsfour = wordsthree[1].Split(splitfour, StringSplitOptions.None);

                        //bldgdata.Boro = wordsfour[0];

                        //string[] splitfive = { "\nVlNumZip :: " };
                        //string[] wordsfive = wordsfour[1].Split(splitfive, StringSplitOptions.None);

                        //bldgdata.BIN = wordsfive[0];

                        //string[] splitsix = { "\nVlTaxBlock :: " };
                        //string[] wordssix = wordsfive[1].Split(splitsix, StringSplitOptions.None);

                        //bldgdata.Zip = wordssix[0];

                        //string[] splitseven = { "[0:FCycle]{" };
                        //string[] wordseven = wordssix[1].Split(splitseven, StringSplitOptions.None);

                        //string[] spliteight = { "}" };
                        //int a = wordseven.ToList().Count - 1;
                        //string[] wordnine = wordseven[a].Split(spliteight, StringSplitOptions.None);

                        //bldgdata.Cycle = wordnine[0];

                        //string[] splitnine = { "[5:FCurrentStatus]{" };
                        //string[] wordsnine = wordssix[1].Split(splitnine, StringSplitOptions.None);

                        //int b = wordsnine.ToList().Count - 1;
                        //string[] wordeleven = wordsnine[b].Split(spliteight, StringSplitOptions.None);

                        //bldgdata.CurrentStatus = wordeleven[0];

                        //string[] splitten = { "[7:FNumStories]{" };
                        //string[] wordtwelve = wordssix[1].Split(splitten, StringSplitOptions.None);

                        //int c = wordtwelve.ToList().Count - 1;
                        //string[] wordthirteen = wordtwelve[c].Split(spliteight, StringSplitOptions.None);

                        //bldgdata.NumStories = wordthirteen[0];

                        scrapeddata.Add(bldgdata);

                        //request.Abort();
                        //myresponse.Close();

                    } 
                }
            }
            //3. exchange List of facadedata to json
            //"C:\Users\EErtugrul\Documents\GitHub\facademy\JSONSrapperdata10010.txt"
            
            var Json = new JavaScriptSerializer().Serialize(scrapeddata.ToArray());

            //File.WriteAllText(@"C:\Users\EErtugrul\Documents\GitHub\facademy\JSONSrapperdata10280.txt", Json);
            //File.WriteAllText(@"C:\Users\EErtugrul\Documents\GitHub\facademy\JSONSrapperdata10280.json", Json);

            File.WriteAllText(@"C:\Users\ELCIN\Documents\GitHub\facademy\JSONSrapperdata10010TEST.txt", Json);
            File.WriteAllText(@"C:\Users\ELCIN\Documents\GitHub\facademy\JSONSrapperdata10010TEST.json", Json);
   
        }
    }
}
