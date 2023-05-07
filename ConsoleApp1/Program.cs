using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApp1
{
    public class Program
    {
        public static void Main(string[] args)
        {

            ReadXMLDataBase();
            WriteAddWord(channelList);
        }

    public static Channels ReadXMLDataBase()
        ///Read data from a file using a data model.
        {
            Channels channels;
            string path = @"C:\Users\n.tolochka\Downloads\data.xml";

            XmlSerializer serializer = new XmlSerializer(typeof(Channels));

            StreamReader reader = new StreamReader(path);
            channels = (Channels)serializer.Deserialize(reader);
            reader.Close();

            Console.WriteLine("Данные отлично считаны!");

            var selectedchannels = channels.ChannelList.Where(p => p.category.Contains("Политика"));
                                   channels.ChannelList.OrderBy(p => p.pubDate);
            return channels;
              
        }

        [Serializable()]
        public class Channel
        {
            [System.Xml.Serialization.XmlElement("title")]
            public string title { get; set; }

            [System.Xml.Serialization.XmlElement("link")]
            public string link { get; set; }

            [System.Xml.Serialization.XmlElement("description")]
            public string description { get; set; }

            [System.Xml.Serialization.XmlElement("category")]
            public string category { get; set; }

            [System.Xml.Serialization.XmlElement("pubDate")]
            public string pubDate { get; set; }
        }


        [XmlRootAttribute("channel")]
        public class Channels
        {
            [XmlElement("item")]
            public Channel[] ChannelList { get; set; }
        }

        public static void WriteAddWord(Channel[] channelList)
        {
            ///Write data to word
            {
                try
                {
                    //Create an instance for word app  
                    Word.Application winword = new Word.Application();

                    //Set animation status for word application  
                    winword.ShowAnimation = false;

                    //Set status for word application is to be visible or not.  
                    winword.Visible = true;

                    //Create a missing variable for missing value  
                    object missing = System.Reflection.Missing.Value;

                    //Create a new document  
                    Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                    if (channelList == null)
                    {
                        Console.WriteLine("Данные из файла не были взяты!");
                        return;
                    }

                    Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);

                    foreach (Channel channel in channelList)
                    {
                        para1.Range.Text = $"\t{channel.title} " + Environment.NewLine;
                        para1.Range.Text = $"\t{channel.link}" + Environment.NewLine;
                        para1.Range.Text = $"\t{channel.description} " + Environment.NewLine;
                        para1.Range.Text = $"\t{channel.category} " + Environment.NewLine;
                        para1.Range.Text = $"\t{channel.pubDate}\n\n\n " + Environment.NewLine;
                    }
                    Console.WriteLine("Данные успешно были записаны в WordAdd.docx файл!");
                    channelList = null;


                    //Save the document  
                    object filename = "WordApp.docx";
                    document.SaveAs2(ref filename);
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                    winword.Quit(ref missing, ref missing, ref missing);
                    winword = null;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                }
            }
        }
    }
}
