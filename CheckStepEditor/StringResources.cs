////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//  MetaAutomation (C) 2016 by Matt Griscom.
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml.Linq;

namespace CheckStepEditor
{
    public class StringResources
    {
        private StringResources()
        {
        }

        private static StringResources m_Instance = null;

        public static StringResources Instance
        {
            get
            {
                if (StringResources.m_Instance == null)
                {
                    StringResources.m_Instance = new StringResources();
                }

                return StringResources.m_Instance;
            }
        }

        private string[] m_StringsBeforeSelectedCode;

        public string[] StringsBeforeSelectedCode
        {
            get
            {
                if (this.m_StringsBeforeSelectedCode == null)
                {
                    try
                    {
                        this.LoadStringsFromXmlFile();
                    }
                    catch (Exception)
                    {
                        this.LoadStringDefaults();
                    }
                }

                return this.m_StringsBeforeSelectedCode;
            }
            private set { }
        }

        private string[] m_StringsAfterSelectedCode;

        public string[] StringsAfterSelectedCode
        {
            get
            {
                if (this.m_StringsAfterSelectedCode == null)
                {
                    try
                    {
                        this.LoadStringsFromXmlFile();
                    }
                    catch (Exception)
                    {
                        this.LoadStringDefaults();
                    }
                }

                return this.m_StringsAfterSelectedCode;
            }
            private set { }
        }

        private void LoadStringsFromXmlFile()
        {
            // Get directory of image of executing DLL, which is where the XML file with the strings is
            string absoluteAssemblyPath = Assembly.GetExecutingAssembly().Location;
            string absoluteStringsFilePath = Path.Combine(Path.GetDirectoryName(absoluteAssemblyPath), "SubstitutionsStrings.xml");
            XDocument stringXDoc = XDocument.Load(absoluteStringsFilePath);
            List<string> beforeSelectedCodeStrings = new List<string>();
            List<string> afterSelectedCodeStrings = new List<string>();

            foreach (XElement stringElement in stringXDoc.Elements("Strings").Elements("BeforeSelectedCode").Elements("String"))
            {
                beforeSelectedCodeStrings.Add(stringElement.Value);
            }

            foreach (XElement stringElement in stringXDoc.Elements("Strings").Elements("AfterSelectedCode").Elements("String"))
            {
                afterSelectedCodeStrings.Add(stringElement.Value);
            }

            m_StringsBeforeSelectedCode = beforeSelectedCodeStrings.ToArray();
            m_StringsAfterSelectedCode = afterSelectedCodeStrings.ToArray();
        }

        private void LoadStringDefaults()
        {
            m_StringsBeforeSelectedCode = new string[] { "Check.Step(\"xxxx.\", delegate", "{" };
            m_StringsAfterSelectedCode = new string[] { "});" };
        }
    }
}
