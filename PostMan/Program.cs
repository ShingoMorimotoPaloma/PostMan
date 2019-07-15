using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using RestSharp;

namespace PostMan
{
    class Program
    {
        static void Main(string[] args)
        {
            var arg = new ArgsSchema(args);
            var body = new BodySchema(arg.WorkBook);
            PostData(arg, body);
            Console.WriteLine("Complete");
        }
        static void PostData(ArgsSchema arg, BodySchema data)
        {
            var client = new RestClient();
            var request = new RestRequest();
            client.BaseUrl = arg.URI;
            request.Method = Method.POST;
            request.AddJsonBody(data);
            var response = client.Execute(request);
        }
        class ArgsSchema
        {
            public ArgsSchema(string[] args)
            {
                _WorkBookPath = args[0];
                _uri = args[1];
                WorkBookOpen();
            }
            string _WorkBookPath;
            string _uri;
            //public string WorkBookPath { get { return _WorkBookPath; } }
            public Uri URI
            {
                get
                {
                    return new Uri(_uri);
                }
            }
            public XLWorkbook WorkBook;
            void WorkBookOpen()
            {
                FileStream fs = new FileStream(_WorkBookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                WorkBook = new XLWorkbook(fs, XLEventTracking.Disabled);
            }

        }
        public class BodySchema
        {
            public List<OneData> data = new List<OneData>();
            public BodySchema(XLWorkbook wb)
            {
                IXLWorksheet ws = wb.Worksheet(1);
                int cnt = 11;
                while (ws.Cell(cnt,4).Value.ToString() != "")
                {
                    var one = new OneData()
                    {
                        Date = ws.Cell(3, 3).Value.ToString(),
                        Name = ws.Cell(cnt, 4).Value.ToString(),
                        ID = int.Parse(ws.Cell(cnt, 5).Value.ToString()),
                        TactTime = long.Parse(ws.Cell(cnt, 6).Value.ToString()),
                        SwUnit = int.Parse(ws.Cell(cnt, 8).Value.ToString()),
                        Lots = int.Parse(ws.Cell(cnt, 7).Value.ToString())
                    };
                    data.Add(one);
                    cnt++;
                }
            }

            public class OneData
            {
                DateTime _dateTime = new DateTime();
                public string Date
                {
                    get
                    {
                        return _dateTime.ToShortDateString();
                    }
                    set
                    {
                        _dateTime = DateTime.Parse(value);
                    }
                }
                public string Name { get; set; }
                public int ID { get; set; }
                public long TactTime { get; set; }
                /// <summary>
                /// スイッチ1回押下あたりの個数
                /// </summary>
                public int SwUnit { get; set; }
                /// <summary>
                /// 生産予定個数
                /// </summary>
                public long Lots { get; set; }
            }
        }
        
    }
}
