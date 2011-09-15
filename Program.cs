// <copyright file="Program.cs" company="Chris Pont">
// Copyright 2011, Chris Pont
// [Copyright]
// </copyright>

namespace FirewallService
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.ServiceProcess;
    using System.Text;

    /// <summary>
    /// Main Entry Point for application
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static void Main()
        {
            ServiceBase[] servicesToRun;
            servicesToRun = new ServiceBase[] { new Service1() };
            ServiceBase.Run(servicesToRun);
        }
    }
}
