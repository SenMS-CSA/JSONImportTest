using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using Newtonsoft.Json;

namespace JSONImportTest
{
    public partial class ThisAddIn
    {
        private JsonImportService _jsonImportService;
        private Office.CommandBarButton _jsonImportButton;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _jsonImportService = new JsonImportService(this.Application);

            // Add custom ribbon button or menu item for JSON import
           // AddJsonImportButton();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up - no need to remove Click event handler when using OnAction
            if (_jsonImportButton != null)
            {
             //   _jsonImportButton.Delete();
                _jsonImportButton = null;
            }
        }

        private void AddJsonImportButton()
        {
            try
            {
                // Add a custom menu item to the Tools menu
                var menuBar = this.Application.CommandBars["Menu Bar"];
                var toolsMenu = menuBar.Controls["Tools"] as Office.CommandBarPopup;

                if (toolsMenu != null)
                {
                    // Check if button already exists to avoid duplicates
                    foreach (Office.CommandBarControl control in toolsMenu.Controls)
                    {
                        if (control.Caption == "Import from JSON...")
                        {
                            control.Delete();
                            break;
                        }
                    }

                    // Add the button
                    _jsonImportButton = toolsMenu.Controls.Add(
                        Office.MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        toolsMenu.Controls.Count + 1,
                        true) as Office.CommandBarButton;

                    if (_jsonImportButton != null)
                    {
                        _jsonImportButton.Caption = "Import from JSON...";
                        _jsonImportButton.TooltipText = "Import Visio diagram from JSON file";

                        // Use OnAction instead of Click event - this is the key fix
                        //_jsonImportButton.OnAction += "JsonImportButtonAction";

                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding menu item: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"Error adding menu item: {ex.Message}", "Menu Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        // This method will be called when the button is clicked
        public void JsonImportButtonAction()
        {
            try
            {
                ImportJsonFile();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in JsonImportButtonAction: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"Error importing JSON: {ex.Message}", "Import Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void ImportJsonFile()
        {
            _jsonImportService?.ImportFromFile();
        }

        // Public method to import from JSON string (useful for automation)
        public void ImportFromJsonString(string jsonContent)
        {
            _jsonImportService?.ImportFromJson(jsonContent);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    public class VisioShape
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("text")]
        public string Text { get; set; }

        [JsonProperty("x")]
        public double X { get; set; }

        [JsonProperty("y")]
        public double Y { get; set; }

        [JsonProperty("width")]
        public double Width { get; set; } = 1.0;

        [JsonProperty("height")]
        public double Height { get; set; } = 0.5;

        [JsonProperty("stencil")]
        public string Stencil { get; set; } = "Basic Shapes";

        [JsonProperty("master")]
        public string Master { get; set; } = "Rectangle";

        [JsonProperty("properties")]
        public Dictionary<string, object> Properties { get; set; } = new Dictionary<string, object>();
    }

    public class VisioConnector
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("fromShape")]
        public string FromShape { get; set; }

        [JsonProperty("toShape")]
        public string ToShape { get; set; }

        [JsonProperty("text")]
        public string Text { get; set; }

        [JsonProperty("connectorType")]
        public string ConnectorType { get; set; } = "Dynamic connector";
    }

    public class VisioDocument
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("template")]
        public string Template { get; set; } = "";

        [JsonProperty("shapes")]
        public List<VisioShape> Shapes { get; set; } = new List<VisioShape>();

        [JsonProperty("connectors")]
        public List<VisioConnector> Connectors { get; set; } = new List<VisioConnector>();
    }

}

