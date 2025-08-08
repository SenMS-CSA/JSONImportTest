using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using Visio = Microsoft.Office.Interop.Visio;

namespace JSONImportTest
{
    public class JsonImportService
    {
        private readonly Visio.Application _visioApp;
        private Dictionary<string, Visio.Shape> _createdShapes;

        public JsonImportService(Visio.Application visioApplication)
        {
            _visioApp = visioApplication ?? throw new ArgumentNullException(nameof(visioApplication));
            _createdShapes = new Dictionary<string, Visio.Shape>();
        }

        public void ImportFromFile()
        {
            try
            {
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                    openFileDialog.Title = "Select JSON file to import";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var jsonContent = File.ReadAllText(openFileDialog.FileName);
                        ImportFromJson(jsonContent);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing JSON file: {ex.Message}", "Import Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ImportFromJson(string jsonContent)
        {
            try
            {
                var document = JsonConvert.DeserializeObject<VisioDocument>(jsonContent);
                CreateVisioDocument(document);
            }
            catch (JsonException ex)
            {
                MessageBox.Show($"Invalid JSON format: {ex.Message}", "JSON Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing JSON: {ex.Message}", "Import Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateVisioDocument(VisioDocument document)
        {
            Visio.Document visioDoc;
            Visio.Page page;

            // Create new document or use active one
            if (string.IsNullOrEmpty(document.Template))
            {
                visioDoc = _visioApp.Documents.Add("");
            }
            else
            {
                visioDoc = _visioApp.Documents.Add(document.Template);
            }

            page = visioDoc.Pages[1];
            _createdShapes.Clear();

            // Set document name if provided
            if (!string.IsNullOrEmpty(document.Name))
            {
                visioDoc.Title = document.Name;
            }

            // Create shapes first
            foreach (var shape in document.Shapes)
            {
                CreateShape(page, shape);
            }

            // Create connectors after all shapes are created
            foreach (var connector in document.Connectors)
            {
                CreateConnector(page, connector);
            }

            // Fit page to contents
            page.ResizeToFitContents();
        }

        private void CreateShape(Visio.Page page, VisioShape shapeData)
        {
            try
            {
                Visio.Shape shape;

                // Try to drop from stencil
                if (!string.IsNullOrEmpty(shapeData.Stencil) && !string.IsNullOrEmpty(shapeData.Master))
                {
                    shape = DropShapeFromStencil(page, shapeData.Stencil, shapeData.Master, 
                        shapeData.X, shapeData.Y);
                }
                else
                {
                    // Create basic rectangle if no stencil specified
                    shape = page.DrawRectangle(shapeData.X, shapeData.Y, 
                        shapeData.X + shapeData.Width, shapeData.Y + shapeData.Height);
                }

                if (shape != null)
                {
                    // Set shape properties
                    if (!string.IsNullOrEmpty(shapeData.Name))
                    {
                        shape.Name = shapeData.Name;
                    }

                    if (!string.IsNullOrEmpty(shapeData.Text))
                    {
                        shape.Text = shapeData.Text;
                    }

                    // Set custom properties
                    foreach (var property in shapeData.Properties)
                    {
                        SetShapeProperty(shape, property.Key, property.Value);
                    }

                    // Store shape for connector creation
                    if (!string.IsNullOrEmpty(shapeData.Id))
                    {
                        _createdShapes[shapeData.Id] = shape;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating shape '{shapeData.Name}': {ex.Message}", 
                    "Shape Creation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private Visio.Shape DropShapeFromStencil(Visio.Page page, string stencilName, 
            string masterName, double x, double y)
        {
            try
            {
                var documents = _visioApp.Documents;
                Visio.Document stencil = null;

                // Try to find already opened stencil
                foreach (Visio.Document doc in documents)
                {
                    if (doc.Name.Contains(stencilName) || doc.Title.Contains(stencilName))
                    {
                        stencil = doc;
                        break;
                    }
                }

                // Open stencil if not found
                if (stencil == null)
                {
                    try
                    {
                        stencil = documents.OpenEx(stencilName, 
                            (short)Visio.VisOpenSaveArgs.visOpenDocked);
                    }
                    catch
                    {
                        // Try with .vss extension
                        stencil = documents.OpenEx(stencilName + ".vss", 
                            (short)Visio.VisOpenSaveArgs.visOpenDocked);
                    }
                }

                if (stencil != null)
                {
                    var master = stencil.Masters[masterName];
                    return page.Drop(master, x, y);
                }
            }
            catch (Exception ex)
            {
                // Fallback to basic shape if stencil/master not found
                System.Diagnostics.Debug.WriteLine($"Could not load stencil/master: {ex.Message}");
            }

            return null;
        }

        private void CreateConnector(Visio.Page page, VisioConnector connectorData)
        {
            try
            {
                if (!_createdShapes.ContainsKey(connectorData.FromShape) ||
                    !_createdShapes.ContainsKey(connectorData.ToShape))
                {
                    return;
                }

                var fromShape = _createdShapes[connectorData.FromShape];
                var toShape = _createdShapes[connectorData.ToShape];

                // Create connector
                var connector = page.DrawLine(
                    fromShape.CellsU["PinX"].ResultIU,
                    fromShape.CellsU["PinY"].ResultIU,
                    toShape.CellsU["PinX"].ResultIU,
                    toShape.CellsU["PinY"].ResultIU);

                // Make it a dynamic connector
                connector.CellsU["ObjType"].FormulaU = "3";

                // Connect to shapes
                connector.CellsU["BeginX"].GlueTo(fromShape.CellsU["PinX"]);
                connector.CellsU["EndX"].GlueTo(toShape.CellsU["PinX"]);

                // Set text if provided
                if (!string.IsNullOrEmpty(connectorData.Text))
                {
                    connector.Text = connectorData.Text;
                }

                // Store connector if it has an ID
                if (!string.IsNullOrEmpty(connectorData.Id))
                {
                    _createdShapes[connectorData.Id] = connector;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating connector: {ex.Message}", 
                    "Connector Creation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SetShapeProperty(Visio.Shape shape, string propertyName, object value)
        {
            try
            {
                // Check if the property already exists
                Visio.Cell propCell = null;
                try
                {
                    propCell = shape.CellsU[$"Prop.{propertyName}"];
                }
                catch
                {
                    // Property does not exist, so add it
                    shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionProp, propertyName, (short)Visio.VisRowTags.visTagDefault);
                    propCell = shape.CellsU[$"Prop.{propertyName}"];
                }

                // Set property value and label
                propCell.FormulaU = $"\"{value}\"";
                shape.CellsU[$"Prop.{propertyName}.Label"].FormulaU = $"\"{propertyName}\"";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not set property {propertyName}: {ex.Message}");
            }
        }
    }
}