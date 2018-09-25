using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace EditTools.Scripts
{
    /// <summary> 
    /// Class for the ribbon procedures
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        /// <summary>
        /// Used to reference the ribbon object
        /// </summary>
        public static Ribbon ribbonref;

        #region | Task Panes |

        /// <summary>
        /// Settings TaskPane
        /// </summary>
        public TaskPane.Settings mySettings;

        /// <summary>
        /// Settings Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;

        #endregion

        #region | Ribbon Events |

        /// <summary> 
        /// The ribbon
        /// </summary>
        public Ribbon()
        {
        }

        /// <summary> 
        /// Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="ribbonID">Represents the XML customization file </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        /// <remarks></remarks>
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("EditingTools.Ribbon_new.xml");
        }

        /// <summary>
        /// Called by the GetCustomUI method to obtain the contents of the Ribbon XML file.
        /// </summary>
        /// <param name="resourceName">name of  the XML file</param>
        /// <returns>the contents of the XML file</returns>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        /// <summary> 
        /// loads the ribbon UI and creates a log record
        /// </summary>
        /// <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code. </param>
        /// <remarks></remarks>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.ribbon = ribbonUI;
                ribbonref = this;
                ThisAddIn.e_ribbon = ribbonUI;
                //AssemblyInfo.SetAddRemoveProgramsIcon("WordAddin.ico");
                AssemblyInfo.SetAssemblyFolderVersion();
                ErrorHandler.SetLogPath();
                ErrorHandler.CreateLogRecord();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Assigns an image to a button on the ribbon in the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        public System.Drawing.Bitmap GetButtonImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnSnippingTool":
                        //return Properties.Resources.snipping_tool;
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;
            }
        }

        /// <summary> 
        /// Assigns the enabled to controls
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public bool GetEnabled(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnAddScriptColumn":
                        //return ErrorHandler.IsEnabled(false);
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;
            }
        }

        /// <summary> 
        /// Assigns text to a label on the ribbon from the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for a label. </returns> 
        public string GetLabelText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tabEditingTools":
                        if (Application.ProductVersion.Substring(0, 2) == "15") //for Excel 2013
                        {
                            return AssemblyInfo.Title.ToUpper();
                        }
                        else
                        {
                            return AssemblyInfo.Title;
                        }
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        DateTime dteCreateDate = Properties.Settings.Default.App_ReleaseDate;
                        return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt");
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary> 
        /// Assigns the number of items for a combobox or dropdown
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns an integer of total count of items used for a combobox or dropdown </returns> 
        public int GetItemCount(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboTableAlias":
                        //return Data.TableAliasTable.Rows.Count;
                    default:
                        return 0;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return 0;
            }
        }

        /// <summary> 
        /// Assigns the values to a combobox or dropdown based on an index
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <param name="index">Represents the index of the combobox or dropdown value </param>
        /// <returns>A method that returns a string per index of a combobox or dropdown </returns> 
        public string GetItemLabel(Office.IRibbonControl control, int index)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboTableAlias":
                        //return UpdateTableAliasComboBoxSource(index);
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary> 
        /// Assigns default values to comboboxes
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for the default value of a combobox </returns> 
        public string GetText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboTableAlias":
                        //return Properties.Settings.Default.Table_ColumnTableAlias;
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public bool GetPressed(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "chkBackstageMarkup":
                        //return Properties.Settings.Default.Visible_mnuScriptType_Markup;
                    default:
                        return true;
                }

            }
            catch (Exception)
            {
                return true;
                //ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary> 
        /// Assigns the visiblity to controls
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is visible </returns> 
        public bool GetVisible(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnScriptTypeXmlValues":
                        //return Properties.Settings.Default.Visible_mnuScriptType_Markup;
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;
            }
        }

        /// <summary>
        /// Assigns the value to an application setting
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                //Ribbon.AppVariables.ControlLabel = GetLabelText(control);
                switch (control.Id)
                {
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="control"></param>
        /// <param name="pressed"></param>
        public void OnAction_Checkbox(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                switch (control.Id)
                {
                    case "chkBackstageMarkup":
                        //Properties.Settings.Default.Visible_mnuScriptType_Markup = pressed;
                        break;
                }

                ribbon.Invalidate();

            }
            catch (Exception)
            {
                //ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary> 
        /// Return the updated value from the comboxbox
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <param name="text">Represents the text from the combobox value </param>
        public void OnChange(Office.IRibbonControl control, string text)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboTableAlias":
                        //Properties.Settings.Default.Table_ColumnTableAlias = text;
                        //Data.InsertRecord(Data.TableAliasTable, text);
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                Properties.Settings.Default.Save();
                ribbon.InvalidateControl(control.Id);
            }
        }

        #endregion

        #region | Ribbon Buttons |

        /// <summary> 
        /// Opens the settings taskpane
        /// </summary>
        /// <remarks></remarks>
        public void OpenSettings()
        {
            try
            {
                if (myTaskPaneSettings != null)
                {
                    if (myTaskPaneSettings.Visible == true)
                    {
                        myTaskPaneSettings.Visible = false;
                    }
                    else
                    {
                        myTaskPaneSettings.Visible = true;
                    }
                }
                else
                {
                    mySettings = new TaskPane.Settings();
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + Scripts.AssemblyInfo.Title);
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneSettings.Width = 675;
                    myTaskPaneSettings.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenReadMe()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenNewIssue()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathNewIssue);

        }

        #endregion

        #region | Subroutines |

        /// <summary> 
        /// Return the count of items in a delimited list
        /// </summary>
        /// <param name="valueList">Represents the list of values in a string </param>
        /// <param name="delimiter">Represents the list delimiter </param>
        /// <returns>the number of values in a delimited string</returns>
        public int GetListItemCount(string valueList, string delimiter)
        {
            try
            {
                string[] comboList = valueList.Split((delimiter).ToCharArray());
                return comboList.GetUpperBound(0) + 1;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return 0;
            }
        }

        /// <summary>
        /// Used to update/reset the ribbon values
        /// </summary>
        public void InvalidateRibbon()
        {
            ribbon.Invalidate();
        }
		
        #endregion

    }
}