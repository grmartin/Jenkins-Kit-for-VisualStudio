using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

using Microsoft.Win32;
using Microsoft.VisualStudio.Coverage.Analysis;

namespace JenkinsKit
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                printHelp(args);
                return;
            }

            switch (args[0].ToLowerInvariant())
            {
                case "codecoverage":
                    codeCoverage(args);
                    break;
                case "help":
                default:
                    printHelp(args);
                    break;
            }
        }

        static void codeCoverage(string[] args)
        {
            // TODO: This needs TONS of cleanup and to be made safe.
            // codecoverage <test dll> <output xml>
            if (args.Length != 3)
            {
                printHelp(args);
                return;
            }

            String pathToDll = args[1];

            String pathToVSTest = PathUtilities.pathToVSTest();

            // Call in Code Coverage
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = pathToVSTest;
            p.StartInfo.Arguments = pathToDll + " /EnableCodeCoverage /InIsolation /Logger:Console";
            p.Start();
            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            // Get file Path for Binary Output
            string resultString = null;
            try
            {
                resultString = Regex.Match(output, @"Attachments:(.*?)\.coverage", RegexOptions.Singleline).Groups[1].Value;
            }
            catch (ArgumentException ex)
            {
                throw new Exception("Cannot parse code coverage file path from output.", ex);
            }

            string binaryCoverageReport = resultString.Trim() + ".coverage";

            // Convert Binary Output to XML
            CoverageInfo cinfo = CoverageInfo.CreateFromFile(binaryCoverageReport);
            CoverageDS ds = cinfo.BuildDataSet(null);

            String xmlDS = Path.ChangeExtension(binaryCoverageReport, "xml");

            ds.WriteXml(xmlDS);

            // Transform XML to Emma
            XPathDocument myXPathDoc = new XPathDocument(xmlDS);
            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            
            string xslt = null;

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("JenkinsKit.MSCoverageToEmma.xslt"))
            using (StreamReader reader = new StreamReader(stream))
            { 
                xslt = reader.ReadToEnd();
            }

            myXslTrans.Load(XmlReader.Create(new StringReader(xslt)));

            // We are going to generate ugly XML, but this is likely due to an untraced issue with the XSLT forcing us to handle the XML as a Fragment not a Document.
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = false;
            settings.IndentChars = "\t";
            settings.WriteEndDocumentOnClose = true;
            settings.Encoding = Encoding.UTF8;
            settings.ConformanceLevel = ConformanceLevel.Auto;

            XmlWriter myWriter = XmlWriter.Create(args[2], settings);

            myXslTrans.Transform(myXPathDoc, null, myWriter);
        }

        static void printHelp(string[] args)
        {
            Console.WriteLine("JenkinsKit");
            Console.WriteLine("");
            Console.WriteLine("codecoverage\t\t<test dll> <emma output xml>");
            Console.WriteLine("");
        }

        protected static class PathUtilities {
            public static string pathToVisualStudio() {
                string[] versionsByPreference = new string[] {"12.0", "11.0", "10.0"};

                string path = null;

                string keyFormat = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\VisualStudio\\{0}";
          
                int i = 0;
                while (path == null && versionsByPreference.Length>i) {
                    object keyVal = Registry.GetValue(String.Format(keyFormat, versionsByPreference[i]), "ShellFolder", null);
                    
                    if (keyVal != null) path = Convert.ToString(keyVal);

                    i++;
                }

                return path;
            }

            public static string pathToVSTest() {
                string vsPath = pathToVisualStudio();

                if (vsPath == null) throw new Exception("Cannot find Visual Studio installation.");

                return String.Format("{0}\\Common7\\IDE\\CommonExtensions\\Microsoft\\TestWindow\\vstest.console.exe", vsPath);
            }
        }
    }
}
