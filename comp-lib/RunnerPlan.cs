using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace compare_lib
{
    public class RunnerPlan
    {
        public string[] FilesA { get; set; }
        public string[] FilesB { get; set; }
        public string CleanRegEx { get; set; }
        public string ReportFolder { get; set; }
        public bool CompValue { get; set; }
        public bool CompStyle { get; set; }
        public bool CompShape { get; set; }

        public void ParseToXml(string xmlPath)
        {
            XmlSerializer writer = new XmlSerializer(typeof(RunnerPlan));
            using (StreamWriter file = new StreamWriter(xmlPath, false, Encoding.UTF8))
            {
                writer.Serialize(file, this);
            }
        }

        public static RunnerPlan ParseFromXml(string xmlPath)
        {
            RunnerPlan scheduleplan;
            XmlSerializer writer = new XmlSerializer(typeof(RunnerPlan));
            using (FileStream ReadFileStream = new FileStream(xmlPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                scheduleplan = (RunnerPlan)writer.Deserialize(ReadFileStream);
                ReadFileStream.Close();
            }
            return scheduleplan;
        }



    }
}
