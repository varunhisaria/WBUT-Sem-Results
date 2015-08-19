using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using HtmlAgilityPack;

namespace WBUTResultScript
{
    class Program
    {
        static HtmlDocument doc = new HtmlDocument();
        static List<Records> records = new List<Records>();
        static void Main(string[] args)
        {
            Asli();
            Console.ReadLine();
        }
        static void Asli()
        {
            int len = 0;
            long start, end, sem;
            string pathDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = "result.csv";
            string filePath = pathDesktop + "\\"+fileName;
            var file = File.CreateText(filePath);
            string header="";
            Console.Write("Enter starting roll number: ");
            start = Convert.ToInt64(Console.ReadLine());
            Console.Write("Enter ending roll number: ");
            end = Convert.ToInt64(Console.ReadLine());
            Console.Write("Enter semester(2,4,6): ");
            sem = Convert.ToInt64(Console.ReadLine());
            for (long i = start; i <=end; i++)
            {
                var request = (HttpWebRequest)WebRequest.Create("http://wbutech.net/show-result_even.php");
                var postData = "semno=" + sem.ToString()+ "&rectype=1"+ "&rollno="+i.ToString();
                var data = Encoding.ASCII.GetBytes(postData);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;
                request.Headers.Add("Origin", "http://wbutech.net");
                request.Referer = "http://wbutech.net/result_even.php";
                ((HttpWebRequest)request).UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.130 Safari/537.36";

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                    stream.Close();
                }

                var response = (HttpWebResponse)request.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                doc.LoadHtml(responseString);
                if (response.ResponseUri.Equals("http://wbutech.net/result_even.php"))
                {
                    response.Close();
                    continue;
                }
                response.Close();
                if (header.Equals(""))
                {
                    header = "Name,Roll No,Registration No,";
                    len = header.Length;
                }


                Records record = new Records();

                record.Name = doc.DocumentNode.SelectSingleNode("//table[1]//tbody[1]//tr[2]//th[1]").InnerText.Substring(8);
                record.RollNo = doc.DocumentNode.SelectSingleNode("//table[1]//tbody[1]//tr[2]//th[2]").InnerText.Substring(11, 12);
                record.RegNo = doc.DocumentNode.SelectSingleNode("//table[1]//tbody[1]//tr[3]//th[1]").InnerText.Substring(20, 12);

                HtmlNode table = doc.DocumentNode.SelectSingleNode("//table[2]");
                HtmlNode tbody = table.SelectSingleNode(".//tbody");
                int cnt,c;
                if (header.Length == len)
                {
                    cnt = 0;
                    foreach (HtmlNode row in tbody.SelectNodes(".//tr").Skip(1))
                    {
                        cnt++;
                    }
                    c = 0;
                    foreach (HtmlNode row in tbody.SelectNodes(".//tr").Skip(1))
                    {
                        
                        c++;
                        if (c < cnt)
                        {
                            header += (row.SelectSingleNode(".//td[2]").InnerHtml);
                            header = header + "("+ row.SelectSingleNode(".//td[1]").InnerHtml + "),";
                        }
                            
                    }
                    header = header + "Odd Sem SGPA,Even Sem SGPA,YGPA\n";
                    file.Write(header);
                }
                    
                cnt = 0;
                foreach (HtmlNode row in tbody.SelectNodes(".//tr").Skip(1))
                {
                    cnt++;
                }
                c=0;
                foreach (HtmlNode row in tbody.SelectNodes(".//tr").Skip(1))
                {
                    if (header.Length==len)
                    {
                        header = "Name,Roll No,Registration No,";
                    }
                    c++;
                    if (c < cnt)
                        record.Marks.Add(row.SelectSingleNode(".//td[3]").InnerHtml);
                }

                table = doc.DocumentNode.SelectSingleNode("//table[3]");
                record.Odd = table.SelectSingleNode(".//tbody[1]//tr[1]//td[1]").InnerText.Substring(47, 4);
                record.Even = table.SelectSingleNode(".//tbody[1]//tr[2]//td[1]").InnerText.Substring(27, 4);
                record.YGPA = table.SelectSingleNode(".//tbody[1]//tr[3]//td[1]").InnerText.Substring(9, 4);
                
                records.Add(record);
                record.print(file);
                Console.SetCursorPosition(0, 3);
                Console.WriteLine("Written {0} record(s)...", i - start + 1);
            }
            Console.SetCursorPosition(0, 3);
            Console.WriteLine("Finished Writing! File {0} saved on Desktop. Press enter key to exit...",fileName);
            file.Close();
        }
    }

    class Records
    {
        public string Name { get; set; }
        public string RollNo { get; set; }
        public string RegNo { get; set; }

        public List<string> Marks = new List<string>();
        public string Odd { get; set; }
        public string Even { get; set; }
        public string YGPA { get; set; }
        public void print(StreamWriter file)
        {
            file.Write(Name + ", ");
            file.Write(RollNo + ", ");
            file.Write(RegNo + ", ");
            foreach (var a in Marks)
            {
                file.Write(a + ", ");
            }
            file.Write(Odd + ", ");
            file.Write(Even + ", ");
            file.Write(YGPA + "");
            file.WriteLine();
        }
    }
}
