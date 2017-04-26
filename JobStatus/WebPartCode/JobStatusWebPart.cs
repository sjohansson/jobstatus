/*
 * SQL Server Agent Job Status display Web part
 * 
 * 2010-07-01 - Stefan Johansson (mailto:stefan@stefanjohansson.org)
 * http://www.stefanjohansson.org
 * 
 * A Visual Web Part for SharePoint 2007 that displays the status of a SQL Server Agent job
 * 
 * created with the spvisualdev tool (http://spvisualdev.codeplex.com/)
 * 
 * install solution, activate feature, add web part to web part page
 * configure web part to connect to SQL Server and an Agent job.
 * 
 * Project home:
 * http://jobstatus.codeplex.com
 * 
 * Icon images courtesy of http://dashboardspy.com/
 * 
 */


using System;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace JobStatus
{
    [Guid("1d4e3971-293d-4711-b555-b188318eaa88")]
    public class JobStatusWebPart : Microsoft.SharePoint.WebPartPages.WebPart
    {
        #region enums
        //enumeration of the different icons available for the status display, corresponds to image names as specified in the AssemblyInfo.cs file and files/embedded resources in the Resources folder
        public enum IconListEnum
        {
            book,
            bug,
            circle,
            cone,
            flag,
            flask,
            pill,
            pin,
            puzzle,
            shield,
            star,
            torso,
            triangle
        }
        #endregion

        #region private members

        //holds result of status check for current job status and last run status
        private string _currentStatus = "";
        private string _lastRunOutcome = "";

        //configurable web part properties
        string _SQLServerName = "Server Name";
        string _SQLServerLoginName = "Login Name";
        string _SQLServerAgentJobName = "Job Name";
        string _SQLServerLoginPassword = "password";

        //statusholders
        private bool _error = false;
        private bool _foundJobName = false;

        //name of the SQL Server stored procedure to use
        private string SQLQuery = "sp_help_job";

        //SQL server connection object to use
        SqlConnection mySqlConn = null;

        //result table components
        private Table tblMain;
        private TableRow mainTableRow;
        private TableCell imageCell;
        private TableCell textCell;

        #endregion

        #region Configurable Web Part Properties
        /*
         * Below are the configurable properties that are available 
         * in the "configure shared web part" dialog in SharePoint
         */

        //Holds the SQL server name
        [
            Personalizable(PersonalizationScope.Shared),
            WebBrowsable(true),
            WebDisplayName("SQL Server Name"),
            WebDescription("SQL server name/address to attach monitor to"),
            System.ComponentModel.Category("Configuration")
        ]
        public string SQLServerName
        {
            get { return _SQLServerName; }
            set { _SQLServerName = value; }
        }

        //Holds the SQL server login name
        [
            Personalizable(PersonalizationScope.Shared),
            WebBrowsable(true),
            WebDisplayName("SQL Server Login Name"),
            WebDescription("SQL server login name to use"),
            System.ComponentModel.Category("Configuration")
        ]
        public string SQLServerLoginName
        {
            get { return _SQLServerLoginName; }
            set { _SQLServerLoginName = value; }
        }

        //Holds the SQL server login password
        [
            Personalizable(PersonalizationScope.Shared),
            WebBrowsable(true),
            WebDisplayName("SQL Server Login Password"),
            WebDescription("SQL server Login password to use"),
            System.ComponentModel.Category("Configuration")
        ]
        public string SQLServerLoginPassword
        {
            get { return _SQLServerLoginPassword; }
            set { _SQLServerLoginPassword = value; }
        }

        //Holds the SQL server Agent job name to monitor
        [
            Personalizable(PersonalizationScope.Shared),
            WebBrowsable(true),
            WebDisplayName("SQL Server Agent Job Name"),
            WebDescription("SQL server agent job name to monitor"),
            System.ComponentModel.Category("Configuration")
        ]
        public string SQLServerAgentJobName
        {
            get { return _SQLServerAgentJobName; }
            set { _SQLServerAgentJobName = value; }
        }

        //Holds the icon/image choice
        private IconListEnum iconList = IconListEnum.star;
        [
            Personalizable(PersonalizationScope.Shared),
            WebBrowsable(true),
            WebDisplayName("Status Icon"),
            WebDescription("The icon to use to diplay the job status"),
            System.ComponentModel.Category("Configuration")
        ]
        public IconListEnum IconList
        {
            get { return iconList; }
            set { iconList = value; }
        }
        #endregion

        #region internal methods

        /// <summary>
        /// Translate the outcome status code to text.
        /// </summary>
        private string getOutcomeText(string status)
        {
            switch (status)
            {
                case "0": return "Failed";
                case "1": return "Succeeded";
                case "3": return "Canceled";
                case "5": return "Unknown";
                default: return "Unknown";
            }
        }
        /// <summary>
        /// Translate the job status code to text.
        /// </summary>
        private string getStatusText(string status)
        {
            switch (status)
            {
                case "1": return "Executing";
                case "2": return "Waiting For Thread";
                case "3": return "Between Retries";
                case "4": return "Idle";
                case "5": return "Suspended";
                case "6": return "Unknown";
                case "7": return "Performing Completion";
                default: return "Unknown";
            }
        }

        /// <summary>
        /// build the resource address to the image icon to use.
        /// </summary>
        private string getIconResourceString()
        {
            string fileNameToImage = "";
            string fileNamePrefix = "JobStatus.Resources.";
            string fileNameSuffix = ".png";

            //If status 1-4, job done successfully, show green icon
            if ((_currentStatus == "4") && (_lastRunOutcome == "1"))
            {
                fileNameToImage = IconList + "_green" ;
            }

            //If status 0-4, job done but it failed, show red icon
            else if ((_currentStatus == "4") && (_lastRunOutcome == "0"))
            {
                fileNameToImage = IconList + "_red";
            }

            //If other status, i.e. ongoing etc. show yellow icon
            else
            {
                fileNameToImage = IconList + "_yellow";
            }

            return fileNamePrefix + fileNameToImage + fileNameSuffix;
        }

        #endregion

        #region Constructor
        /// <summary>
        /// Constructor
        /// </summary>
        public JobStatusWebPart()
        {
            this.ExportMode = WebPartExportMode.All;
        }
        #endregion

        #region override methods
        /// <summary>
        /// Creates controls for rendering.
        /// </summary>
        protected override void CreateChildControls()
        {
            if (!_error)
            {
                try
                {
                    base.CreateChildControls();

                    if (SQLServerName != "") //If SQL Server name has been configured (i.e. normal condition)
                    {
                        //Create connection string from settings
                        string ConnStr = "Data Source=" + SQLServerName + ";Initial Catalog=msdb;User Id=" + SQLServerLoginName + ";Password=" + SQLServerLoginPassword + ";";
                        SqlConnection mySqlConn = null;

                        //create and open connection
                        mySqlConn = new SqlConnection(ConnStr);
                        mySqlConn.Open();

                        //create sql command for the specified stored procedure
                        SqlCommand mySqlCmd = new SqlCommand(SQLQuery, mySqlConn);
                        mySqlCmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //read through the resulting rows and extract status if available
                        using (SqlDataReader reader = mySqlCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (string.Format("{0}", reader[2]) == SQLServerAgentJobName)
                                {
                                    //found the job name in the job list, extract proper columns from the result
                                    _currentStatus = string.Format("{0}", reader[25]);
                                    _lastRunOutcome = string.Format("{0}", reader[21]);
                                    _foundJobName = true;
                                }
                            }
                        }

                        //close the connection
                        mySqlConn.Close();

                        //Render result to the end user if the job was found
                        if (_foundJobName)
                        {
                            //information is presented in the table created and formatted below.
                            tblMain = new Table();
                            mainTableRow = new TableRow();

                            //the cell that contains the status image
                            imageCell = new TableCell();

                            //Style the imageCell
                            TableItemStyle imageCellStyle = new TableItemStyle();
                            imageCellStyle.HorizontalAlign = HorizontalAlign.Center;
                            imageCellStyle.VerticalAlign = VerticalAlign.Middle;
                            imageCellStyle.Width = Unit.Pixel(40);

                            //apply style
                            imageCell.ApplyStyle(imageCellStyle);

                            //construct cell contents for the status icon cell
                            Image statusImage = new Image();
                            statusImage.ImageAlign = ImageAlign.Middle;
                            statusImage.ImageUrl = Page.ClientScript.GetWebResourceUrl(typeof(JobStatusWebPart), getIconResourceString());
                            imageCell.Controls.Add(statusImage);

                            //the cell that contains the status text
                            textCell = new TableCell();

                            //Style the textCell
                            TableItemStyle textCellStyle = new TableItemStyle();
                            textCellStyle.HorizontalAlign = HorizontalAlign.Left;
                            textCellStyle.VerticalAlign = VerticalAlign.Middle;
                            //textCellStyle.Width = Unit.Pixel(40);

                            //apply style
                            textCell.ApplyStyle(textCellStyle);

                            //construct cell contents for the status text cell
                            textCell.Controls.Add(new LiteralControl("Execution Status: " + getStatusText(_currentStatus) + "<br/>"));
                            textCell.Controls.Add(new LiteralControl("Last Run Outcome: " + getOutcomeText(_lastRunOutcome)));

                            //piece together table
                            mainTableRow.Cells.Add(imageCell);
                            mainTableRow.Cells.Add(textCell);
                            tblMain.Rows.Add(mainTableRow);

                            //add the table to the page
                            this.Controls.Add(tblMain);

                        }
                        else //job name not found, display error message
                        {
                            this.Controls.Add(new LiteralControl("Error: agent job name not found, please check configuration."));
                        }
                    }
                    else //if no server is specified
                    {
                        // Add message to specify server name to enable web part functionality
                        this.Controls.Add(new LiteralControl("No SQL Server name specified, please configure web part to enable its functionality..."));
                    }
                }
                catch (Exception ex)
                {
                    mySqlConn.Close();
                    HandleException(ex);
                }
            }
        }


        /// <summary>
        /// Ensures that the CreateChildControls() is called before events.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            if (!_error)
            {
                try
                {
                    base.OnLoad(e);
                    this.EnsureChildControls();
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }
            }
        }

        #endregion

        #region Exception Management
        /// <summary>
        /// Clear all child controls and add an error message for display.
        /// </summary>
        /// <param name="ex"></param>
        private void HandleException(Exception ex)
        {
            this._error = true;
            this.Controls.Clear();
            this.Controls.Add(new LiteralControl(ex.Message));
        }
        #endregion
    }
}
