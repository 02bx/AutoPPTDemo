using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CSharpeLibrary
{
    public class Ini
    {
        public void FileWrite(string Path, string content)
        {
            FileStream ofs = new FileStream(Path, FileMode.Create);
            StreamWriter sw = new StreamWriter(ofs);
            sw.WriteLine(content);
            sw.Close();
            ofs.Close();
        }

        public void FileWrite(string Path, string[] content)
        {
            FileStream ofs = new FileStream(Path, FileMode.Create);
            StreamWriter sw = new StreamWriter(ofs);
            for (int i = 0; i < content.Length; i++)
            {
                sw.WriteLine(content[i]);
            }
            sw.Close();
            ofs.Close();
        }

        public void FileWriteAppend(string Path, string content)
        {
            FileStream ofs = new FileStream(Path, FileMode.Append);
            StreamWriter sw = new StreamWriter(ofs);
            sw.WriteLine(content);
            sw.Close();
            ofs.Close();
        }

        public ArrayList FileRead(string Path)
        {
            FileStream ifs = new FileStream(Path, FileMode.Open);
            StreamReader sr = new StreamReader(ifs);
            ArrayList list = new ArrayList();
            while (!sr.EndOfStream)
            {
                list.Add(sr.ReadLine());
            }
            sr.Close();
            ifs.Close();
            return list;
        }

        public string[] cfgRead(string Path, string[] name)
        {
            string[] value = new string[name.Length];
            ArrayList list = FileRead(Path);
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < name.Length; j++)
                {
                    string cfgName = list[i].ToString().Split('=')[0];
                    string cfgValue = list[i].ToString().Split('=')[1];
                    if (cfgName == name[j])
                    {
                        value[j] = cfgValue;
                    }
                }
            }
            return value;
        }
    }
}
