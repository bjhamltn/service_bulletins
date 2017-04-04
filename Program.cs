﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length>0)
            {
                if (args[0].Trim() != "")
                {
                    Application.Run(new Form1(args[0]));
                    return;
                }
                
            }
            Application.Run(new Form1());
        }
    }
}
