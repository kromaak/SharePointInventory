using System.IO;
using System.Text;
using MNIT.Utilities;

namespace MNIT.Inventory
{
    public class WriteReports
    {
        public static void WriteText(string[] args)
        {
            // Write data to CSV file
            StringBuilder builder = new StringBuilder();
            StreamWriter streamWriter= new StreamWriter(args[0], true, Encoding.UTF8);
            for (int j = 1; j < args.Length; j++)
            {
                builder.Append(Csv.Escape(args[j]));
                builder.Append(',');
            }
            streamWriter.WriteLine(builder);
            streamWriter.Close();
        }
    }
}
