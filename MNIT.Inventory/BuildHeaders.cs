using System;

namespace MNIT.Inventory
{
    public class BuildHeaders
    {
        public static void WriteReportHeaders(string[] args)
        {
            // Get the value of the initial report file path
            // should be something like C:\temp\20160531-1240DetailedWorkflowReport.csv
            string outputFilePath = args[1];
            // Create Stream Writer object
            //StreamWriter streamWriter = new StreamWriter(outputFilePath);
            // If 1st arg is one action, write report headers for that action, else write headers for all reports
            string action = args[0];
            switch (action)
            {
                case "all":
                    WriteGroupReportHeaders(outputFilePath);
                    WriteInfoPathReportHeaders(outputFilePath);
		            WriteListReportHeaders(outputFilePath);
                    WriteStandardWfReportHeaders(outputFilePath);
                    WriteUserReportHeaders(outputFilePath);
                    WriteWebsReportHeaders(outputFilePath);
                    break;
                case "groups":
                    WriteGroupReportHeaders(outputFilePath);
                    break;
                case "infopath":
                    WriteInfoPathReportHeaders(outputFilePath);
                    break;
                case "lists":
	            case "versions":
		            WriteListReportHeaders(outputFilePath);
                    break;
                case "standard":
                    WriteStandardWfReportHeaders(outputFilePath);
                    break;
                case "users":
                    WriteUserReportHeaders(outputFilePath);
                    break;
                case "webs":
                    WriteWebsReportHeaders(outputFilePath);
                    break;
                default:
                    WriteInfoPathReportHeaders(outputFilePath);
                    WriteListReportHeaders(outputFilePath);
                    WriteStandardWfReportHeaders(outputFilePath);
                    WriteWebsReportHeaders(outputFilePath);
                    break;
            }
            
        }

        public static void WriteDetailedWfReportHeaders(string outputFilePath)
        {
            // Write the standard info CSV Header
            string detailedWfReportPath = outputFilePath;
            // Write the Headers for the Detailed Workflow Instances report
            string[] passingDetailedWfHeaderObject = new string[11];
            passingDetailedWfHeaderObject[0] = detailedWfReportPath;
            passingDetailedWfHeaderObject[1] = "Site ID";
            passingDetailedWfHeaderObject[2] = "Web ID";
            passingDetailedWfHeaderObject[3] = "Site Name";
            passingDetailedWfHeaderObject[4] = "Site URL";
            passingDetailedWfHeaderObject[5] = "Site Owner";
            passingDetailedWfHeaderObject[6] = "List Name";
            passingDetailedWfHeaderObject[7] = "List URL";
            passingDetailedWfHeaderObject[8] = "Workflow Type";
            passingDetailedWfHeaderObject[9] = "Workflow Name";
            passingDetailedWfHeaderObject[10] = "Running Instances";
            WriteReports.WriteText(passingDetailedWfHeaderObject);

            // Write the CSV header for the rolled up Inventory function
            string rollupDetailedWfReportPath = outputFilePath.Replace("DetailedWorkflow", "RollupDetailedWorkflow");
            // Write the Headers for the Rollup Workflow Instances report
            string[] passingDetailedWfRollupHeaderObject = new string[3];
            passingDetailedWfRollupHeaderObject[0] = rollupDetailedWfReportPath;
            passingDetailedWfRollupHeaderObject[1] = "SiteURL";
            passingDetailedWfRollupHeaderObject[2] = "RunningInstances";
            WriteReports.WriteText(passingDetailedWfRollupHeaderObject);
        }

        public static void WriteGroupReportHeaders(string outputFilePath)
        {
            // Write the header data to CSV file
            string detailedGroupReportPath = outputFilePath.Replace("DetailedWorkflow", "ADGroups");
            // Create a header for the AD Group report
            string[] passingGroupHeaderObject = new string[6];
            passingGroupHeaderObject[0] = detailedGroupReportPath;
            passingGroupHeaderObject[1] = "Web Application";
            passingGroupHeaderObject[2] = "Site ID";
            passingGroupHeaderObject[3] = "Site URL";
            passingGroupHeaderObject[4] = "AD Groups";
            passingGroupHeaderObject[5] = "Permission Levels";
            WriteReports.WriteText(passingGroupHeaderObject);
        }

        public static void WriteInfoPathReportHeaders(string outputFilePath)
        {
            // Write the standard info CSV Header
            string detailedInfoPathReportPath = outputFilePath.Replace("DetailedWorkflow", "InfoPath");
            // Create a header for the InfoPath report
            string[] passingInfoPathHeaderObject = new string[11];
            passingInfoPathHeaderObject[0] = detailedInfoPathReportPath;
            passingInfoPathHeaderObject[1] = "Web Application";
            passingInfoPathHeaderObject[2] = "Site ID";
            passingInfoPathHeaderObject[3] = "Web ID";
            passingInfoPathHeaderObject[4] = "Site Name";
            passingInfoPathHeaderObject[5] = "Site URL";
            passingInfoPathHeaderObject[6] = "Site Owner";
            passingInfoPathHeaderObject[7] = "List Name";
            passingInfoPathHeaderObject[8] = "List URL";
            passingInfoPathHeaderObject[9] = "List Template Type";
            passingInfoPathHeaderObject[10] = "External Connections";
            WriteReports.WriteText(passingInfoPathHeaderObject);
            // Write the rollup info CSV Header
            string rollupInfoPathReportPath = outputFilePath.Replace("DetailedWorkflow", "RollupInfoPath");
            // Create a header for the Rollup InfoPath report
            string[] passingIpRollupHeaderObject = new string[4];
            passingIpRollupHeaderObject[0] = rollupInfoPathReportPath;
            passingIpRollupHeaderObject[1] = "Site URL";
            passingIpRollupHeaderObject[2] = "InfoPath Lists";
            passingIpRollupHeaderObject[3] = "External Connections";
            WriteReports.WriteText(passingIpRollupHeaderObject);
        }

        public static void WriteListReportHeaders(string outputFilePath)
        {
            // Write the CSV Header
            string detailedListReportPath = outputFilePath.Replace("DetailedWorkflow", "ListItemVersions");
            // Create a header for the list report
            string[] passingLvHeaderObject = new string[17];
            passingLvHeaderObject[0] = detailedListReportPath;
            passingLvHeaderObject[1] = "Web Application";
            passingLvHeaderObject[2] = "Site ID";
            passingLvHeaderObject[3] = "Web ID";
            passingLvHeaderObject[4] = "Site Name";
            passingLvHeaderObject[5] = "Site URL";
            passingLvHeaderObject[6] = "Site Owner";
            passingLvHeaderObject[7] = "List Name";
            passingLvHeaderObject[8] = "List URL";
            passingLvHeaderObject[9] = "Number of Versions";
            passingLvHeaderObject[10] = "Total List Item Count";
            passingLvHeaderObject[11] = "Folder Count";
            passingLvHeaderObject[12] = "File Count";
            passingLvHeaderObject[13] = "Checked In";
            passingLvHeaderObject[14] = "Checked Out";
            passingLvHeaderObject[15] = "Never Been Checked In";
            passingLvHeaderObject[16] = "Manage Checked Out Docs";
            WriteReports.WriteText(passingLvHeaderObject);
            // Write the rollup CSV Header
            string rollupListReportPath = outputFilePath.Replace("DetailedWorkflow", "RollupVersions");
            // Create a header for the Rollup List report
            string[] passingLvRollupHeaderObject = new string[5];
            passingLvRollupHeaderObject[0] = rollupListReportPath;
            passingLvRollupHeaderObject[1] = "Site URL";
            passingLvRollupHeaderObject[2] = "Large Lists";
            passingLvRollupHeaderObject[3] = "Unlimited Versions";
            passingLvRollupHeaderObject[4] = "Checked out docs";
            WriteReports.WriteText(passingLvRollupHeaderObject);
        }

        public static void WriteStandardWfReportHeaders(string outputFilePath)
        {
            // Write the standard report CSV Header
            string detailedWorkflowReportPath = outputFilePath.Replace("DetailedWorkflow", "StandardWorkflow");
            // Create a header for the Standard Workfloow report
            string[] passingStandardHeaderObject = new string[12];
            passingStandardHeaderObject[0] = detailedWorkflowReportPath;
            passingStandardHeaderObject[1] = "Web Application";
            passingStandardHeaderObject[2] = "Site ID";
            passingStandardHeaderObject[3] = "Web ID";
            passingStandardHeaderObject[4] = "Site Name";
            passingStandardHeaderObject[5] = "Site URL";
            passingStandardHeaderObject[6] = "Site Owner";
            passingStandardHeaderObject[7] = "List Name";
            passingStandardHeaderObject[8] = "List URL";
            passingStandardHeaderObject[9] = "Workflow Type";
            passingStandardHeaderObject[10] = "Workflow Name";
            passingStandardHeaderObject[11] = "Workflow ID";
            WriteReports.WriteText(passingStandardHeaderObject);
            // Write the rollup CSV Header
            string rollupWorkflowReportPath = outputFilePath.Replace("DetailedWorkflow", "RollupStandardWorkflow");
            // Create a header for the Rollup Workflow report
            string[] passingStandardRollupHeaderObject = new string[5];
            passingStandardRollupHeaderObject[0] = rollupWorkflowReportPath;
            passingStandardRollupHeaderObject[1] = "Site URL";
            passingStandardRollupHeaderObject[2] = "Nintex";
            passingStandardRollupHeaderObject[3] = "SPD2010WFs";
            passingStandardRollupHeaderObject[4] = "SPD2013WFs";
            WriteReports.WriteText(passingStandardRollupHeaderObject);
        }

        public static void WriteUserReportHeaders(string outputFilePath)
        {
            // Write the CSV Header
            string detailedUserReportPath = outputFilePath.Replace("DetailedWorkflow", "User");
            // Create a header for the user report
            string[] passingUserHeaderObject = new string[6];
            passingUserHeaderObject[0] = detailedUserReportPath;
            passingUserHeaderObject[1] = "Web Application";
            passingUserHeaderObject[2] = "Site ID";
            passingUserHeaderObject[3] = "Site URL";
            passingUserHeaderObject[4] = "Users";
            passingUserHeaderObject[5] = "Permission Levels";
            WriteReports.WriteText(passingUserHeaderObject);
        }

        public static void WriteWebsReportHeaders(string outputFilePath)
        {
            // Write header data to CSV file
            string detailedWebsReportPath = outputFilePath.Replace("DetailedWorkflow", "Webs");
            // Create a header for the Webs report
            string[] passingWebHeaderObject = new string[19];
            passingWebHeaderObject[0] = detailedWebsReportPath;
            passingWebHeaderObject[1] = "Web Application";
            passingWebHeaderObject[2] = "Site ID";
            passingWebHeaderObject[3] = "Web ID";
            passingWebHeaderObject[4] = "Site Name";
            passingWebHeaderObject[5] = "Site URL";
            passingWebHeaderObject[6] = "Site Owner";
            passingWebHeaderObject[7] = "Template";
            passingWebHeaderObject[8] = "Sandbox URL";
            passingWebHeaderObject[9] = "Sandbox Solutions";
            passingWebHeaderObject[10] = "SharePoint Version";
            passingWebHeaderObject[11] = "Host Type";
            passingWebHeaderObject[12] = "Site Master Page";
            passingWebHeaderObject[13] = "System Master Page";
            passingWebHeaderObject[14] = "Drop Off Library";
            passingWebHeaderObject[15] = "Custom Page Layouts";
            passingWebHeaderObject[16] = "List Template Library";
            passingWebHeaderObject[17] = "Storage Size";
            passingWebHeaderObject[18] = "Sub Sites";
            //passingWebHeaderObject[19] = "Site Logo URL";
            //passingWebHeaderObject[20] = "Alternate CSS URL";
            WriteReports.WriteText(passingWebHeaderObject);
            // Write the rollup CSV Header
            string rollupWebsReportPath = outputFilePath.Replace("DetailedWorkflow", "RollupWebs");
            // Create a header for the Rollup Webs report
            string[] passingWebRollupHeaderObject = new string[11];
            passingWebRollupHeaderObject[0] = rollupWebsReportPath;
            passingWebRollupHeaderObject[1] = "Site URL";
            passingWebRollupHeaderObject[2] = "Workspace Templates";
            passingWebRollupHeaderObject[3] = "Custom Solutions";
            passingWebRollupHeaderObject[4] = "Custom Master Pages";
            passingWebRollupHeaderObject[5] = "Custom Page Layouts";
            passingWebRollupHeaderObject[6] = "Custom Pages";
            passingWebRollupHeaderObject[7] = "Deployed Apps";
            passingWebRollupHeaderObject[8] = "DropOff Libraries";
            passingWebRollupHeaderObject[9] = "List Templates";
            passingWebRollupHeaderObject[10] = "Sub Sites";
            //passingWebRollupHeaderObject[11] = "Exported Web Parts";
            WriteReports.WriteText(passingWebRollupHeaderObject);
            // Create a new file path for adding custom page layouts to a new custom report to act against
            string customPagesPath = outputFilePath.Replace("DetailedWorkflow", "Pages");
            // Create a header for the Custom Pages and Page Layouts report
            string[] passingDetailedPagesHeaderObject = new string[7];
            passingDetailedPagesHeaderObject[0] = customPagesPath;
            passingDetailedPagesHeaderObject[1] = "Web Application";
            passingDetailedPagesHeaderObject[2] = "Site ID";
            passingDetailedPagesHeaderObject[3] = "Web ID";
            passingDetailedPagesHeaderObject[4] = "Page Url";
            passingDetailedPagesHeaderObject[5] = "Page Layout Description";
            passingDetailedPagesHeaderObject[6] = "Page Layout URL";
            //passingDetailedPagesHeaderObject[4] = "Exported Web Part";
            //passingDetailedPagesHeaderObject[5] = "Page Modifiers";
            WriteReports.WriteText(passingDetailedPagesHeaderObject);
        }
    }
}
