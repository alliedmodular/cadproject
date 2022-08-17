using System;
using System.IO;
using System.Reflection;
using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using System.Windows.Media.Imaging;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace AlliedModularBuilding
{
    public class Ribbon : Autodesk.AutoCAD.Runtime.IExtensionApplication
    {
        public RibbonCombo pan1ribcombo1 = new RibbonCombo();
        public RibbonCombo pan3ribcombo = new RibbonCombo();

        public Ribbon()
        {
            pan3ribcombo.CurrentChanged += new EventHandler<RibbonPropertyChangedEventArgs>(pan3ribcombo_CurrentChanged);
        }

        private void pan3ribcombo_CurrentChanged(object sender, RibbonPropertyChangedEventArgs e)
        {
            RibbonButton but = pan3ribcombo.Current as RibbonButton;
            acadApp.ShowAlertDialog(but.Text);
        }
        public void Initialize()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //ed.WriteMessage("Initializing - do something useful.");
            MyRibbon();
        }

        public void Terminate()
        {
            Console.WriteLine("Cleaning up...");
        }
        // [CommandMethod("MyRibbon")]
        public void MyRibbon()
        {
            try
            {
                Autodesk.Windows.RibbonControl ribbonControl = Autodesk.Windows.ComponentManager.Ribbon;
                RibbonTab Tab = new RibbonTab();
                Tab.Title = Utility.Constraint.TabName;
                Tab.Id = Utility.Constraint.TabID;

                ribbonControl.Tabs.Add(Tab);

                #region Added Tab button in the Auto-Cad menu 
                // create Ribbon panels
                Autodesk.Windows.RibbonPanelSource panel1Panel = new RibbonPanelSource();
                panel1Panel.Title = Utility.Constraint.RibbonPanelName;
                RibbonPanel Panel1 = new RibbonPanel();
                Panel1.Source = panel1Panel;
                Tab.Panels.Add(Panel1);

                #endregion Added Tab button in the Auto-Cad menu End


                #region button 1 code strat

                RibbonButton button = new RibbonButton();
                button.Text = Utility.Constraint.ButtonBOMName;
                button.Id = Utility.Constraint.ButtonBOMID;
                button.ShowText = true;
                button.ShowImage = true;
                button.Image = GetEmbeddedImage(Utility.Constraint.ButtonBOMImage);
                button.LargeImage = GetEmbeddedImage(Utility.Constraint.ButtonBOMLargeImage);
                button.Orientation = System.Windows.Controls.Orientation.Vertical;
                button.Size = RibbonItemSize.Large;
                button.CommandHandler = new RibbonCommandHandler();
                #endregion button 1 code End

                #region button 2 code strat

                //RibbonButton button1 = new RibbonButton();
                //button1.Text = "Button1";
                //button1.ShowText = true;
                //button1.ShowImage = true;
                //button1.Image = GetEmbeddedImage(Utility.Constraint.ButtonBOMImage);
                //button1.LargeImage = GetEmbeddedImage(Utility.Constraint.ButtonBOMLargeImage);
                //button1.Orientation = System.Windows.Controls.Orientation.Vertical;
                //button1.Size = RibbonItemSize.Large;
                //button1.CommandHandler = new RibbonCommandHandler();

                #endregion button 2 code End

                #region button 3 code strat

                //RibbonButton button2 = new RibbonButton();
                //button2.Text = "Button2";
                //button2.ShowText = true;
                //button2.ShowImage = true;
                //button2.Image = GetEmbeddedImage(Utility.Constraint.ButtonBOMImage);
                //button2.LargeImage = GetEmbeddedImage(Utility.Constraint.ButtonBOMLargeImage);
                //button2.Orientation = System.Windows.Controls.Orientation.Vertical;
                //button2.Size = RibbonItemSize.Large;
                //button2.CommandHandler = new RibbonCommandHandler();
                #endregion button 3 code End


                #region button 4 code strat

                //RibbonButton button3 = new RibbonButton();
                //button3.Text = "Button3";
                //button3.ShowText = true;
                //button3.ShowImage = true;
                //button3.Image = GetEmbeddedImage(Utility.Constraint.ButtonBOMImage);
                //button3.LargeImage = GetEmbeddedImage(Utility.Constraint.ButtonBOMLargeImage);
                //button3.Orientation = System.Windows.Controls.Orientation.Vertical;
                //button3.Size = RibbonItemSize.Large;
                //button3.CommandHandler = new RibbonCommandHandler();

                #endregion button 4 code End

                #region button 5 code strat

                //RibbonButton button4 = new RibbonButton();
                //button4.Text = "Button4";
                //button4.ShowText = true;
                //button4.ShowImage = true;
                //button4.Image = GetEmbeddedImage(Utility.Constraint.ButtonBOMImage);
                //button4.LargeImage = GetEmbeddedImage(Utility.Constraint.ButtonBOMLargeImage);
                //button4.Orientation = System.Windows.Controls.Orientation.Vertical;
                //button4.Size = RibbonItemSize.Large;
                //button4.CommandHandler = new RibbonCommandHandler();

                #endregion button 5 code End

                #region Added Seperator between each button
                panel1Panel.Items.Add(button);
                panel1Panel.Items.Add(new RibbonSeparator());
                //panel1Panel.Items.Add(button1);
                //panel1Panel.Items.Add(new RibbonSeparator());
                //panel1Panel.Items.Add(button2);
                //panel1Panel.Items.Add(new RibbonSeparator());
                //panel1Panel.Items.Add(button3);
                //panel1Panel.Items.Add(new RibbonSeparator());
                //panel1Panel.Items.Add(button4);
               // panel1Panel.Items.Add(new RibbonSeparator());

                #endregion Added Seperator between each button End

                Tab.IsActive = true;
            }
            catch (System.Exception ex)
            {
                throw;
            }
        }
        public static BitmapSource GetEmbeddedImage(string name)
        {
            try
            {
                Assembly a = Assembly.GetExecutingAssembly();
                Stream s = a.GetManifestResourceStream(name);
                return BitmapFrame.Create(s);
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    
    }
}
