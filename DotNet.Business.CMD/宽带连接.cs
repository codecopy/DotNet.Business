﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNet.Business.CMD
{
    class 宽带连接
    {
        public static string Connect(string UserS, string PwdS)
        {
            string arg = @"rasdial.exe 宽带连接" + " " + UserS + " " + PwdS;
            return InvokeCmd(arg);
        }

        public static string Disconnect()
        {
            string arg = string.Format("rasdial \"{0}\" /disconnect", "宽带连接");
            return InvokeCmd(arg);
        }

        private static string InvokeCmd(string cmdArgs)
        {
            string Tstr = "";
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();

            p.StandardInput.WriteLine(cmdArgs);
            p.StandardInput.WriteLine("exit");
            Tstr = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            p.Close();
            return Tstr;
        }
    }
}
