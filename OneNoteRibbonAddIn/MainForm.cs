using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNoteRibbonAddIn
{
    [ComVisible(false)]
    public partial class MainForm : Form
    {
        private readonly OneNote.Application _oneNoteApp;

        public MainForm(OneNote.Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
            InitializeComponent();
        }

        private void btnEnumerateNotebooks_Click(object sender, EventArgs e)
        {
            MessageBox.Show(EnumNotebooks());
        }

        private void btnEnumerateSections_Click(object sender, EventArgs e)
        {
            MessageBox.Show(EnumSections());
        }

        private void btnGetPageTitle_Click(object sender, EventArgs e)
        {
            MessageBox.Show(GetPageTitle());
        }

        private void btnGetPageContent_Click(object sender, EventArgs e)
        {
            MessageBox.Show(GetPageContent());
        }

        private void btnUpdatePageContent_Click(object sender, EventArgs e)
        {
            MessageBox.Show(UpdatePageContent());
        }

        private string EnumNotebooks()
        {
            string notebookXml;
            _oneNoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsNotebooks, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;

            StringBuilder sb = new StringBuilder();
            foreach (var notebookNode in from node in doc.Descendants(ns + "Notebook") select node)
            {
                sb.AppendLine(notebookNode.Attribute("name").Value);
            }
            return sb.ToString();
        }

        private string EnumSections()
        {
            string notebookXml;
            _oneNoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsSections, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            StringBuilder sb = new StringBuilder();
            foreach (var notebookNode in from node in doc.Descendants(ns + "Notebook") select node)
            {
                sb.AppendLine(notebookNode.Attribute("name").Value);
                foreach (var sectionNode in from node in notebookNode.Descendants(ns + "Section") select node)
                {
                    sb.AppendLine("  " + sectionNode.Attribute("name").Value);
                }
            }
            return sb.ToString();
        }

        private string GetPageTitle()
        {
            string pageXmlOut = GetActivePageContent();
            var doc = XDocument.Parse(pageXmlOut);

            return doc.Descendants().FirstOrDefault().Attribute("ID").NextAttribute.Value;
        }

        private string GetPageContent()
        {
            string notebookXml;
            _oneNoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            var pageNode = doc.Descendants(ns + "Page").FirstOrDefault(n => n.Attribute("name").Value == GetPageTitle());
            StringBuilder sb = new StringBuilder();
            if (pageNode != null)
            {
                string pageXml;
                _oneNoteApp.GetPageContent(pageNode.Attribute("ID").Value, out pageXml);
                sb.AppendLine(XDocument.Parse(pageXml).ToString());
            }
            return sb.ToString();
        }

        private string UpdatePageContent()
        {
            string notebookXml;
            _oneNoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            var pageNode = doc.Descendants(ns + "Page").FirstOrDefault(n => n.Attribute("name").Value == GetPageTitle());
            var existingPageId = pageNode.Attribute("ID").Value;

            var page = new XDocument(new XElement(ns + "Page",
                new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                            new XElement(ns + "T",
                                new XCData("Current date/time: " +
                                           DateTime.Now)))))));

            page.Root.SetAttributeValue("ID", existingPageId);
            _oneNoteApp.UpdatePageContent(page.ToString(), DateTime.MinValue);

            return "Update Complete";
        }

        /// <summary>
        /// Get active page content and output the xml string
        /// </summary>
        /// <returns>string</returns>
        private string GetActivePageContent()
        {
            string activeObjectId = GetActiveObjectId(ObjectType.Page);
            string pageXmlOut;
            _oneNoteApp.GetPageContent(activeObjectId, out pageXmlOut);

            return pageXmlOut;
        }

        /// <summary>
        /// Get ID of current page 
        /// </summary>
        /// <param name="obj">_Object Type</param>
        /// <returns>current page Id</returns>
        private string GetActiveObjectId(ObjectType obj)
        {
            string currentPageId = "";
            uint count = _oneNoteApp.Windows.Count;
            foreach (OneNote.Window window in _oneNoteApp.Windows)
            {
                if (window.Active)
                {
                    switch (obj)
                    {
                        case ObjectType.Notebook:
                            currentPageId = window.CurrentNotebookId;
                            break;
                        case ObjectType.Section:
                            currentPageId = window.CurrentSectionId;
                            break;
                        case ObjectType.SectionGroup:
                            currentPageId = window.CurrentSectionGroupId;
                            break;
                    }

                    currentPageId = window.CurrentPageId;
                }
            }

            return currentPageId;
        }

        private enum ObjectType
        {
            Notebook,
            Section,
            SectionGroup,
            Page,
            SelectedPages,
            PageObject
        }
    }
}
