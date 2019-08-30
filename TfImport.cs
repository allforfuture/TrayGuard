using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml.Serialization;
using System.IO;

namespace TrayGuard
{
    class TfImport
    {
        public string CartonNumber { get; set; }

        public static List<TfImport> loadCartonListFromDesktopCsv(string path)
        {
            var tf = new List<TfImport>();

            foreach (var line in File.ReadAllLines(path))
            {
                var columns = line.Split(',');

                string buff = columns[0].Trim();
                if (!string.IsNullOrEmpty(buff))
                {
                    tf.Add(new TfImport
                    {
                        CartonNumber = columns[0].Trim(),
                    });
                }
            }

            return tf;
        }
    }
}
