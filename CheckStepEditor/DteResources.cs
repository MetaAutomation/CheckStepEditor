using System;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//  MetaAutomation (C) 2016 by Matt Griscom.
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System.Runtime.InteropServices;

namespace CheckStepEditor
{
    public class DteResources
    {
        private static DteResources m_Instance = null;

        private DteResources()
        {
        }

        public static DteResources Instance
        {
            get
            {
                if (DteResources.m_Instance == null)
                {
                    DteResources.m_Instance = new DteResources();
                }

                return DteResources.m_Instance;
            }
        }

        [DllImport("ole32.dll")]
        private static extern void CreateBindCtx(int reserved,
            out System.Runtime.InteropServices.ComTypes.IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern void GetRunningObjectTable(int reserved,
            out System.Runtime.InteropServices.ComTypes.IRunningObjectTable prot);

        public EnvDTE._DTE GetDteInstance()
        {
            //rot entry for visual studio running under current process.
            string rotEntry = String.Format("!VisualStudio.DTE.14.0:{0}", GetDteProcess().Id);
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            GetRunningObjectTable(0, out rot);
            System.Runtime.InteropServices.ComTypes.IEnumMoniker enumMoniker;
            rot.EnumRunning(out enumMoniker);
            enumMoniker.Reset();
            IntPtr fetched = IntPtr.Zero;
            System.Runtime.InteropServices.ComTypes.IMoniker[] moniker = new System.Runtime.InteropServices.ComTypes.IMoniker[1];
            while (enumMoniker.Next(1, moniker, fetched) == 0)
            {
                System.Runtime.InteropServices.ComTypes.IBindCtx bindCtx;
                CreateBindCtx(0, out bindCtx);
                string displayName;
                moniker[0].GetDisplayName(bindCtx, null, out displayName);
                if (displayName == rotEntry)
                {
                    object comObject;
                    rot.GetObject(moniker[0], out comObject);
                    return (EnvDTE._DTE)comObject;
                }
            }
            return null;
        }

        private System.Diagnostics.Process GetDteProcess()
        {
            // IF debugging with an experimental IDE instance
            //return GetProcessWhileDebuggingWithExperimentalInstance();

            // IF using the extension with your IDE
            return GetProcessForThisIdeInstance();
        }

        private System.Diagnostics.Process GetProcessWhileDebuggingWithExperimentalInstance()
        {
            System.Diagnostics.Process[] vSProcesses = System.Diagnostics.Process.GetProcessesByName("devenv");
            System.Diagnostics.Process ideProcess = null;

            foreach (System.Diagnostics.Process process in vSProcesses)
            {
                if (process.MainWindowTitle.ToLower().Contains("experimental"))
                {
                    ideProcess = process;
                    break;
                }
            }

            return ideProcess;
        }

        private System.Diagnostics.Process GetProcessForThisIdeInstance()
        {
            return System.Diagnostics.Process.GetCurrentProcess();
        }
    }
}
