using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Microsoft;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADExcel.MyCommands))]

namespace AutoCADExcel
{

    public class MyCommands
    {

        [CommandMethod("makeRibbon")]
        public static void MakeRibbon()
        {   
            //init RibbonConrol
            RibbonControl ribbon = ComponentManager.Ribbon;
            if(ribbon == null) { return; }

            //make Tab
            RibbonTab ribbonTab = ribbon.FindTab("Plugin");
            if(ribbonTab != null) { ribbon.Tabs.Remove(ribbonTab); }
            ribbonTab = new RibbonTab();
            ribbonTab.Title = "Plugin";
            ribbonTab.Id = "Plugin";

            //add it to the other Tabs
            ribbon.Tabs.Add(ribbonTab);

            //make Panel
            RibbonPanelSource PanelSource = new RibbonPanelSource();
            PanelSource.Title = "Export to Excel";
            RibbonPanel Panel = new RibbonPanel();
            Panel.Source = PanelSource;

            //Make Button
            RibbonButton RibbonButton;
            RibbonButton = new RibbonButton();
            RibbonButton.Name = "Export Button";
            RibbonButton.ShowImage = true;
            RibbonButton.CommandHandler = new CommandHandler();
            RibbonButton.Size = RibbonItemSize.Large;

            //make BitmapImage for the Button
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            image.UriSource = new Uri("pack://application:,,,/AutoCADExcel;component/excelLarge.png");
            image.EndInit();
            RibbonButton.LargeImage = image;

            //add Button to Panel
            PanelSource.Items.Add(RibbonButton);

            //add Panel to Tab
            ribbonTab.Panels.Add(Panel);
        }

        [CommandMethod("MyGroup", "ExportToExcel", "MyPickFirstLocal", CommandFlags.UsePickSet)]
        public static void ExportToExcel()
        {
            //count to manage the rows
            int count = 1;
            //init excel
            Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
            if(excelApplication == null)
            {
                Globals.doc.Editor.WriteMessage("Error: Excel is not installed.");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApplication.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            worksheet.Name = "ObjectProperties";
            worksheet.Cells[1, 1] = "Properties";
            worksheet.Cells[1, 2] = "Value";
            
            //filter for GetProperties
            var flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
            
            //get Selection
            PromptSelectionResult result = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetSelection();

            if (result.Status == PromptStatus.OK)
            {
                //data string for textblock
                string data = "";

                //AcTransaction class wrapper
                Transaction acTransaction = Globals.doc.TransactionManager.StartTransaction();

                //array of selected Objects
                ObjectId[] selectedObjectArray = result.Value.GetObjectIds();

                foreach (ObjectId selectedObject in selectedObjectArray)
                {
                    //grab every selected object
                    Entity entity = (Entity)acTransaction.GetObject(selectedObject, OpenMode.ForRead);

                    //get entity type
                    Type typeOfEntity = entity.GetType();

                    //advance line in data and add type to data
                    data += typeOfEntity.Name + Environment.NewLine;

                    //advance row a excel and add type to excel
                    count++;
                    worksheet.Cells[count, 1] = typeOfEntity.Name;

                    //get properties of entity type
                    System.Reflection.PropertyInfo[] propertiesArray = typeOfEntity.GetProperties(flags);
                    foreach (PropertyInfo property in propertiesArray)
                    {
                        try
                        {
                            //advance line in data and add property of type and its value for the given entity to data
                            data += property.Name + "    " + property.GetValue(entity).ToString() + Environment.NewLine;

                            //advance row in excel and add property of type and its value for the given entity to excel
                            count++;
                            worksheet.Cells[count, 1] = property.Name;
                            worksheet.Cells[count, 2] = property.GetValue(entity).ToString();

                        }
                        catch (System.Exception e)
                        {
                            //print error if something fails
                            Globals.doc.Editor.WriteMessage("Error: \n {0}", e);
                        }
                      
                    }
                    //advance line in data for next object
                    data += Environment.NewLine + Environment.NewLine + Environment.NewLine;

                    //advance row in excel for next object
                    count++;

                    //delete unneeded entity
                    entity.Dispose();
                }
                //commit changes (?)
                acTransaction.Commit();

                //open save Dialog
                Application.ShowModalWindow(new SaveAsPrompt(data, workbook));

                //close unneeded excel inits

                excelApplication.Quit();
            }
            else
            {
                //return if no selects found
                return;
            }
        }

    }
    public class CommandHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            if (parameter is RibbonButton)
            {
                RibbonButton button = parameter as RibbonButton;

                MyCommands.ExportToExcel();
            }
        }
    }

    public class Globals
    {
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
    }

}
