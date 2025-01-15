using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;


const string InputDirectory = "C:\\src\\data\\pdf-to-onenote\\AHEAD NOFO Final 11.15.2023 508";
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

    foreach(var directory in directories)
    {
        var section = sectionDocument.Descendants(oneNs + "Section").FirstOrDefault(s => s.Attribute("name")?.Value == directory.Name);
        if (section == null)
        {
            Console.WriteLine($"Skipping {directory.Name}. It was not found in the xml");
            continue;
        }

        foreach(var file in directory.GetFiles())
        {
            onenoteApp.CreateNewPage(section.Attribute("ID")!.Value, out string pageId);
            string pageContent = @$"<one:Page xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote' ID='{pageId}'>
                  <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:T><![CDATA[Page {file.Name.Replace(file.Extension, "")}]]></one:T>
                      </one:OE>
                      <one:OE>
                        <one:T><![CDATA[{File.ReadAllText(file.FullName)}]]></one:T>
                      </one:OE>
                    </one:OEChildren>
                  </one:Outline>
                </one:Page>";
            onenoteApp.UpdatePageContent(pageContent);
        }        
    }
}
catch(Exception ex)
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
