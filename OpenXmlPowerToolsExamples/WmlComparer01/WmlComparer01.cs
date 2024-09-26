// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace OpenXmlPowerTools
{
    class WmlComparer01
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("TestExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlComparerSettings settings = new WmlComparerSettings() { TrackFormatting = true, TrackMoves = true};

            //WmlDocument result = WmlComparer.Compare(
            //    new WmlDocument(@"C:\Users\sezgi\source\repos\Open-Xml-PowerTools\OpenXmlPowerToolsExamples\WmlComparer01\Source1.docx"),
            //    new WmlDocument(@"C:\Users\sezgi\source\repos\Open-Xml-PowerTools\OpenXmlPowerToolsExamples\WmlComparer01\Source2.docx"),
            //    settings);

            WmlDocument result = WmlComparer.CompareDocuments(
                @"C:\Users\sezgi\source\repos\Open-Xml-PowerTools\OpenXmlPowerToolsExamples\WmlComparer01\Source1.docx",
               @"C:\Users\sezgi\source\repos\Open-Xml-PowerTools\OpenXmlPowerToolsExamples\WmlComparer01\Source2.docx",
                settings);

            result.SaveAs(Path.Combine(tempDi.FullName, "Compared.docx"));

            var revisions = WmlComparer.GetRevisions(result, settings);
            foreach (var rev in revisions)
            {
                Console.WriteLine("Author: " + rev.Author);
                Console.WriteLine("Revision type: " + rev.RevisionType);
                Console.WriteLine("Revision text: " + rev.Text);
                Console.WriteLine();
            }
        }
    }
}
