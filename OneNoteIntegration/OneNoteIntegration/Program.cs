using Microsoft.Office.Interop.OneNote;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Xml.Linq;


const string InputDirectory = "C:\\src\\data\\pdf-to-onenote\\2nd Round V2 Revised Project Narrative 052224 With Tables";
const string OutputDirectory = "C:\\src\\data\\pdf-to-onenote\\OneNoteNotebooks";
if (!Directory.Exists(OutputDirectory))
{
    Directory.CreateDirectory(OutputDirectory);
}

var inputDirectoryInfo = new DirectoryInfo(InputDirectory);
string inputDirectoryName = inputDirectoryInfo.Name.Replace(" ", "-").Replace(".", "");

string outputNotebook = Path.Combine(OutputDirectory, inputDirectoryName);


Application? onenoteApp = null;

try
{
    if (Directory.Exists(outputNotebook))
    {
        Directory.Delete(outputNotebook, true);
    }

    onenoteApp = new Application();
    onenoteApp.OpenHierarchy(outputNotebook, null, out string notebookId, CreateFileType.cftNotebook);

    XNamespace oneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    var directories = inputDirectoryInfo.GetDirectories();

    string newSectionsXml = string.Join("", directories.Select(d => $"<one:Section name='{d.Name}' />"));

    onenoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out string _);
    string createdSectionsXml = @$"<one:Notebooks xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote'>
  <one:Notebook ID='{notebookId}'>
    {newSectionsXml}
  </one:Notebook>
</one:Notebooks>";

    onenoteApp.UpdateHierarchy(createdSectionsXml);

    onenoteApp.GetHierarchy(null, HierarchyScope.hsSections, out string sectionXml);

    var sectionDocument = XDocument.Parse(sectionXml);

    foreach (var directory in directories)
    {
        string targetPath = Path.Combine(outputNotebook, directory.Name);
        var section = sectionDocument.Descendants(oneNs + "Section").FirstOrDefault(s => s.Attribute("path")!.Value!.StartsWith(targetPath));
        if (section == null)
        {
            Console.WriteLine($"Skipping {directory.Name}. It was not found in the xml");
            continue;
        }

        foreach (var file in directory.GetFiles())
        {
            onenoteApp.CreateNewPage(section.Attribute("ID")!.Value, out string pageId);

            string fileName = file.Name.Replace(file.Extension, "");
            string pageContentText = File.ReadAllText(file.FullName);

            int startOfAsciiTable = pageContentText.IndexOf("+--");
            int endOfAsaciiTable = pageContentText.LastIndexOf("--+");
            if (startOfAsciiTable != -1 && endOfAsaciiTable != -1)
            {
                string contentBeforeTable = pageContentText.Substring(0, startOfAsciiTable);
                string contentAfterTable = pageContentText.Substring(endOfAsaciiTable + 3);

                string asciiTable = pageContentText.Substring(startOfAsciiTable, endOfAsaciiTable - startOfAsciiTable + 3);
                string[] tableLines = asciiTable.Split(Environment.NewLine);
                var tableBuilder = new StringBuilder();
                var columnAccumulator = new List<List<string>>();
                tableBuilder.AppendLine("<one:Table bordersVisible=\"true\" hasHeaderRow=\"true\">");
                bool addedColumns = false;
                for (int i = 0; i < tableLines.Length; i++)
                {
                    string line = tableLines[i];
                    if (columnAccumulator.Any() && !addedColumns)
                    {
                        addedColumns = true;
                        tableBuilder.AppendLine("<one:Columns>");
                        for (int l = 0; l < columnAccumulator.Count; l++)
                        {
                            tableBuilder.AppendLine($"<one:Column index=\"{l}\" width=\"300\"/>");
                        }
                        tableBuilder.AppendLine("</one:Columns>");
                    }

                    if (line.Contains("+---") || line.Contains("+==="))
                    {
                        if (columnAccumulator.Any())
                        {
                            tableBuilder.AppendLine("<one:Row>");
                            foreach (var accumulator in columnAccumulator)
                            {
                                tableBuilder.AppendLine($"<one:Cell><one:OEChildren><one:OE alignment=\"left\" quickStyleIndex=\"1\"><one:T><![CDATA[{string.Join(" ", accumulator)}]]></one:T></one:OE></one:OEChildren></one:Cell>");
                                accumulator.Clear();
                            }
                            tableBuilder.AppendLine("</one:Row>");
                        }
                        continue;
                    }

                    string[] rawColumns = tableLines[i].Split('|');
                    // skip this first then take 2 minus the total to ignore the end and account for taking one away
                    string[] lineColumns = rawColumns.Skip(1).Take(rawColumns.Length - 2).Select(c => c.Trim()).ToArray();
                    if (!columnAccumulator.Any())
                    {
                        for (int k = 0; k < lineColumns.Length; k++)
                        {
                            columnAccumulator.Add(new List<string>());
                        }
                    }

                    for (int j = 0; j < lineColumns.Length; j++)
                    {
                        columnAccumulator[j].Add(lineColumns[j]);
                    }
                }

                tableBuilder.AppendLine("</one:Table>");
                pageContentText = $"<one:OE><one:T><![CDATA[{contentBeforeTable}]]></one:T></one:OE><one:OE>{tableBuilder}</one:OE><one:OE><one:T><![CDATA[{contentAfterTable}]]></one:T></one:OE>";

                string pageContent = @$"<one:Page xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote' ID='{pageId}'>
                  <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:T><![CDATA[Page {fileName}]]></one:T>
                      </one:OE>
                      {pageContentText}
                    </one:OEChildren>
                  </one:Outline>
                </one:Page>";

                onenoteApp.UpdatePageContent(pageContent);
            }
            else
            {
                string pageContent = @$"<one:Page xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote' ID='{pageId}'>
                  <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:T><![CDATA[Page {fileName}]]></one:T>
                      </one:OE>
                      <one:OE>
                        <one:T><![CDATA[{pageContentText}]]></one:T>
                      </one:OE>
                    </one:OEChildren>
                  </one:Outline>
                </one:Page>";
                onenoteApp.UpdatePageContent(pageContent);
            }
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}
finally
{
    if (onenoteApp != null)
    {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(onenoteApp);
        onenoteApp = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}

//foreach (var inputSubdirectoryInfo in inputDirectoryInfo.GetDirectories())
//{
//    onenoteApp.UpdateHierarchy(notebookXml);
//}
