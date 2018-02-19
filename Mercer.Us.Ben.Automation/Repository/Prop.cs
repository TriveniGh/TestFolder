using Mercer.Us.Ben.Automation.Utility_Code;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercer.Us.Ben.Automation.Repository

{
  public  class Prop

    {

        public static  void  ReadDictionaryFile() 
        {

          string CurrentProjectPath = ResultSheet.CurrentProjectPath();
          string Repositorypath = CurrentProjectPath + "Mercer.US.Ben.Automation" + Path.DirectorySeparatorChar + "Repository" + Path.DirectorySeparatorChar+"OR.txt";

            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            foreach (string line in File.ReadAllLines(Repositorypath))
            {
                if ((!string.IsNullOrEmpty(line)) &&
                    (!line.StartsWith(";")) &&
                    (!line.StartsWith("#")) &&
                    (!line.StartsWith("'")) &&
                    (line.Contains('=')))
                {
                    int index = line.IndexOf('=');
                    string key = line.Substring(0, index).Trim();
                    string value = line.Substring(index + 1).Trim();

                    if ((value.StartsWith("\"") && value.EndsWith("\"")) ||
                        (value.StartsWith("'") && value.EndsWith("'")))
                    {
                        value = value.Substring(1, value.Length - 2);
                    }
                    dictionary.Add(key, value);
                }
            }

        
        }
    }
}
