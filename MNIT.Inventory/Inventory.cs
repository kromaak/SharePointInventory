using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security;
using System.Text;

using Utils = MNIT.Utilities;

namespace MNIT.Inventory
{
    class Inventory
    {
        private static void Main(string[] args)
        {
            #region console user input
            // use local credentials to work against O365 site
            ConsoleColor defaultForeground = Console.ForegroundColor;

            // User Enters the path to store the inventory csv file
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(
                "Enter the Path for the inventory file to be created (formatted like 'C:\\Temp\\')");
            Console.WriteLine("Do not enter file name.  That will be created for you with a date-time stamp.");

            Console.ForegroundColor = defaultForeground;
            string enteredFilePath = Console.ReadLine();
            string defaultFilePath = "C:\\Temp\\" + DateTime.Today.ToString("yyyyMMdd") + "-" +
                                     DateTime.Now.ToString("HHmm") + "DetailedWorkflowReport.csv";
            //string defaultListFilePath = "C:\\Temp\\SiteList.csv";
            string defaultListFilePath = "C:\\Temp\\";
            string filePath;
            string siteListFilePath;
            StreamWriter streamWriter;
            if (!string.IsNullOrEmpty(enteredFilePath))
            {
                if (enteredFilePath.Trim().EndsWith(@"\"))
                {
                    filePath = enteredFilePath + DateTime.Today.ToString("yyyyMMdd") + "-" +
                               DateTime.Now.ToString("HHmm") + "DetailedWorkflowReport.csv";
                    // Set up the Site Collection file to read from
                    siteListFilePath = enteredFilePath;
                }
                else
                {
                    filePath = enteredFilePath + "\\" + DateTime.Today.ToString("yyyyMMdd") + "-" +
                      DateTime.Now.ToString("HHmm") + "DetailedWorkflowReport.csv";
                    // Set up the Site Collection file to read from
                    siteListFilePath = enteredFilePath + "\\";
                }
                streamWriter = new StreamWriter(filePath, true, Encoding.UTF8);
            }
            else
            {
                filePath = defaultFilePath;
                streamWriter = new StreamWriter(filePath, true, Encoding.UTF8);
                // Set up the Site Collection file to read from
                siteListFilePath = defaultListFilePath;
            }

            // User Enters the operation to complete
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter the type of report you want to receive");
            //Console.WriteLine("    'groups' = a list of all AD Groups and custom permission levels");
            Console.WriteLine("    'webs' = a list of all Sites");
            Console.WriteLine("    'detailed' = a count of running Workflow instances");
            Console.WriteLine("    'standard' = a list of all Workflows");
            Console.WriteLine("    'infopath' = a count of InfoPath form libraries and external connections");
            Console.WriteLine("    'versions' = a count of Large Lists and lists with unlimited versioning configured");

            Console.ForegroundColor = defaultForeground;
            string defaultAction = "standard";
            string enteredAction = Console.ReadLine();
            string action = "";
            action = !string.IsNullOrEmpty(enteredAction) ? enteredAction : defaultAction;

            // User Enters login name
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your login name");

            Console.ForegroundColor = defaultForeground;
            string userLoginName;
            userLoginName = Console.ReadLine();

            // User Enters password
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password");

            Console.ForegroundColor = defaultForeground;
            //string userPassword;
            SecureString userPassword = Utils.Pass.GetPasswordFromConsoleInput();

            // User Enters domain
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your domain");

            Console.ForegroundColor = defaultForeground;
            string userDomain;
            userDomain = Console.ReadLine();

            //// If User is a global admin they can enter a number of input files
            //string numberOfInputFiles = "";
            //if (!string.IsNullOrEmpty(userDomain) && userDomain == "EAD")
            //{
            //    Console.ForegroundColor = ConsoleColor.Green;
            //    Console.WriteLine("How many SiteList Files do you have?");
            //    Console.WriteLine("Just hit Enter if only 1");
            //    Console.WriteLine("If more than 1, each file must have a number in the file name, like SiteList1.csv, SiteList2.csv, etc.");

            //    Console.ForegroundColor = defaultForeground;
            //    numberOfInputFiles = Console.ReadLine();
            //}
            #endregion
            // Call the ConsoleSpinner class to let people know that something is happening
            Console.Write("Working...");
            Utils.SpinAnimation.Start();
            // Build the user object
            Utils.ActingUser actingUser = new Utils.ActingUser(userLoginName, userPassword, userDomain);
            // Run the method to inventory RUNNING WORKFLOW INSTANCES
            if (action.ToLower() == "detailed")
            {
                // Write the CSV Header
                string[] passingDetailedWfHeaderObject = new string[11];
                passingDetailedWfHeaderObject[0] = filePath;
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
                string rollupFilePath = filePath.Replace("DetailedWorkflow", "RollupDetailedWorkflow");
                string[] passingDetailedWfRollupHeaderObject = new string[3];
                passingDetailedWfRollupHeaderObject[0] = rollupFilePath;
                passingDetailedWfRollupHeaderObject[1] = "SiteURL";
                passingDetailedWfRollupHeaderObject[2] = "RunningInstances";
                WriteReports.WriteText(passingDetailedWfRollupHeaderObject);
                string[] readUrls = System.IO.File.ReadAllLines(siteListFilePath + "SiteList.csv");
                int everyTen = 0;
                foreach (string readCurrentLine in readUrls)
                {
                    if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                    {
                        string currentLine = readCurrentLine.Trim();
                        everyTen++;
                        // Declare variables
                        int runningInstancesCounter = 0;
                        // Perform the action on the current web
                        try
                        {
                            // Write the site URL every ten lines from CSV, to let the user know progress is being made
                            if (everyTen % 10 == 0 && everyTen != 0)
                            {
                                Utils.SpinAnimation.Stop();
                                Console.WriteLine();
                                Console.WriteLine(
                                    @"Getting Workflow Running Instance info for the address provided: {0}", currentLine);
                                Utils.SpinAnimation.Start();
                            }
                            GetDetailedWorkflows.InventoryWorkflowsDetailed(currentLine, actingUser,
                                ref runningInstancesCounter, filePath);
                            // Run the SP Inventory List Update function
                            if (args.Length > 0 && args[0] == "i")
                            {
                                if (args[1].Length > 0 && args[1].Contains("http"))
                                {
                                    WriteDetailedWorkflowObjectToInvList(currentLine, ref runningInstancesCounter,
                                        args[1]);
                                }
                            }
                            // Run the rolled up Inventory function
                            string[] passingDetailedWfObject = new string[3];
                            passingDetailedWfObject[0] = filePath;
                            passingDetailedWfObject[1] = currentLine;
                            passingDetailedWfObject[2] = runningInstancesCounter.ToString();
                            WriteReports.WriteText(passingDetailedWfObject);
                        }
                        catch (WebException webException)
                        {
                            HttpWebResponse errorResponse = webException.Response as HttpWebResponse;
                            // If the error code from the attempt is a 404 or similar, inform the user that the site doesn't exist or is unreachable
                            if (!string.IsNullOrEmpty(errorResponse.ToString()) &&
                                errorResponse.StatusCode == HttpStatusCode.NotFound)
                            {
                                Utils.SpinAnimation.Stop();
                                Console.WriteLine();
                                Console.WriteLine(@"Could not find the site at the address provided: {0}", currentLine);
                                Utils.SpinAnimation.Start();
                            }
                            // If the error code from the attempt is a 401 or similar, inform the user that they are not authorized
                            if (!string.IsNullOrEmpty(errorResponse.ToString()) &&
                                errorResponse.StatusCode == HttpStatusCode.Unauthorized)
                            {
                                Utils.SpinAnimation.Stop();
                                Console.WriteLine();
                                Console.WriteLine(
                                    @"You do not have permissions with your current credentials for the site at this address: {0}",
                                    currentLine);
                                Utils.SpinAnimation.Start();
                            }
                        }
                    }
                }
                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", filePath);
            }
            // Run the method to inventory AD GROUPS
            if (action.ToLower() == "groups")
            {
                // Write the CSV Header
                string detailedGroupReportPath = filePath.Replace("DetailedWorkflow", "ADGroups");
                streamWriter.Close();
                System.IO.File.Move(filePath, detailedGroupReportPath);
                // Create a header for the user report
                //string[] passingGroupHeaderObject = new string[5];
                //passingGroupHeaderObject[0] = detailedGroupReportPath;
                //passingGroupHeaderObject[1] = "Site ID";
                //passingGroupHeaderObject[2] = "Site URL";
                //passingGroupHeaderObject[3] = "AD Groups";
                //passingGroupHeaderObject[4] = "Permission Levels";
                //WriteReports.WriteText(passingGroupHeaderObject);
                // Create a string of arguments to run the inventory function
                string[] passingGroupsHeaderArgs = new string[2];
                passingGroupsHeaderArgs[0] = action;
                passingGroupsHeaderArgs[1] = filePath;
                BuildHeaders.WriteReportHeaders(passingGroupsHeaderArgs);

                //// Check for multiple input files
                //if (numberOfInputFiles.Length > 0)
                //{
                //    // if there is an argument that is an integer
                //    int n;
                //    bool isNumeric = int.TryParse(numberOfInputFiles, out n);
                //    // and the argument specifies n for a number of input files
                //    if (isNumeric)
                //    {

                //        // run the inventory function n number of times, once for each input file
                //        for (int r = 0; r < n; r++)
                //        {
                //            // Need to start with file n + 1
                //            int fileNumber = r + 1;
                //            // get the inputFile with the number of iteration for each file that exists
                //            string inputFile = siteListFilePath + "SiteList" + fileNumber + ".csv";
                //            // Create a string of arguments to run the inventory function
                //            string[] passingUserArgs = new string[3];
                //            passingUserArgs[0] = inputFile;
                //            passingUserArgs[1] = action;
                //            passingUserArgs[2] = detailedGroupReportPath;
                //            ChooseReport.RunUserInventory(passingUserArgs, actingUser);
                //        }
                //    }
                //}
                //else
                //{
                    // run the inventory function one time
                    string inputFile = siteListFilePath + "SiteList.csv";
                    // Create a string of arguments to run the inventory function
                    string[] passingUserArgs = new string[3];
                    passingUserArgs[0] = inputFile;
                    passingUserArgs[1] = action;
                    passingUserArgs[2] = detailedGroupReportPath;
                    ChooseReport.RunUserInventory(passingUserArgs, actingUser);
                //}

                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", detailedGroupReportPath);
            }
            // Run the method to inventory INFOPATH COUNTS
            if (action.ToLower() == "infopath")
            {
                // Write the standard info CSV Header
                string detailedInfoPathReportPath = filePath.Replace("DetailedWorkflow", "InfoPath");
                streamWriter.Close();
                System.IO.File.Move(filePath, detailedInfoPathReportPath);
                //// Write the rollup info CSV Header
                string rollupInfoPathReportPath = filePath.Replace("DetailedWorkflow", "RollupInfoPath");
                // Create a string of arguments to run the inventory function
                string[] passingInfoPathHeaderArgs = new string[2];
                passingInfoPathHeaderArgs[0] = action;
                passingInfoPathHeaderArgs[1] = filePath;
                BuildHeaders.WriteReportHeaders(passingInfoPathHeaderArgs);

                // run the inventory function one time
                string inputFile = siteListFilePath + "SiteList.csv";
                // Create a string of arguments to run the inventory function
                string[] passingIpArgs = new string[3];
                passingIpArgs[0] = inputFile;
                passingIpArgs[1] = detailedInfoPathReportPath;
                passingIpArgs[2] = rollupInfoPathReportPath;
                ChooseReport.RunInfoPathInventory(passingIpArgs, actingUser);

                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", detailedInfoPathReportPath);
            }
            // This is incomplete, but will eventually call all the appropriate sub routines to get all info about lists
            // instead of just infopath or just versions
            // Run the method to inventory UNLIMITED VERSION SETTINGS and INFOPATH FORM DATA for lists
            if (action.ToLower() == "lists")
            {
                // Write the CSV Header
                string sitesFilePath = filePath.Replace("DetailedWorkflow", "List");
                streamWriter.Close();
                System.IO.File.Move(filePath, sitesFilePath);
                // Create a header for the user report
                string[] passingListHeaderObject = new string[11];
                passingListHeaderObject[0] = sitesFilePath;
                passingListHeaderObject[1] = "Site ID";
                passingListHeaderObject[2] = "Web ID";
                passingListHeaderObject[3] = "Site Name";
                passingListHeaderObject[4] = "Site URL";
                passingListHeaderObject[5] = "Site Owner";
                passingListHeaderObject[6] = "List Name";
                passingListHeaderObject[7] = "List URL";
                passingListHeaderObject[8] = "Unlimited Versions";
                passingListHeaderObject[9] = "List Template Type";
                passingListHeaderObject[10] = "External Connections";
                WriteReports.WriteText(passingListHeaderObject);
                // Write the rollup CSV Header
                string rollupFilePath = filePath.Replace("DetailedWorkflow", "RollupVersions");
                // Create a header for the user report
                string[] passingListRollupHeaderObject = new string[6];
                passingListRollupHeaderObject[0] = rollupFilePath;
                passingListRollupHeaderObject[1] = "SiteURL";
                passingListRollupHeaderObject[2] = "LargeLists";
                passingListRollupHeaderObject[3] = "UnlimitedVersions";
                passingListRollupHeaderObject[4] = "InfoPathForms";
                passingListRollupHeaderObject[5] = "ExternalConnections";
                WriteReports.WriteText(passingListRollupHeaderObject);
                // Read through the list of sites
                string[] readUrls = System.IO.File.ReadAllLines(siteListFilePath + "SiteList.csv");
                int everyTen = 0;
                // For each site address in the CSV file
                foreach (string readCurrentLine in readUrls)
                {
                    if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                    {
                        string currentLine = readCurrentLine.Trim();
                        everyTen++;
                        // Counter Variables
                        int largeListCounter = 0;
                        int unlimitedVerCounter = 0;
                        int infoPathFormCounter = 0;
                        int infoPathExternalConnCounter = 0;
                        // Run the inventory function for List Information
                        try
                        {
                            // Write the site URL every ten lines from CSV, to let the user know progress is being made
                            if (everyTen % 10 == 0 && everyTen != 0)
                            {
                                Utils.SpinAnimation.Stop();
                                Console.WriteLine();
                                Console.WriteLine(@"Getting List info for the address provided: {0}", currentLine);
                                Utils.SpinAnimation.Start();
                            }
                            GetLists.InventoryLists(currentLine, actingUser, ref largeListCounter,
                                ref unlimitedVerCounter, ref infoPathFormCounter, ref infoPathExternalConnCounter, sitesFilePath);
                            // Run the SP Inventory List Update function
                            if (args.Length > 0 && args[0] == "i")
                            {
                                if (args[1].Length > 0 && args[1].Contains("http"))
                                {
                                    WriteListVersionObjectToInvList(currentLine, ref largeListCounter,
                                        ref unlimitedVerCounter, args[1]);
                                }
                            }
                            // Run the rolled up Inventory function
                            string[] passingRollupListObject = new string[6];
                            passingRollupListObject[0] = rollupFilePath;
                            passingRollupListObject[1] = currentLine;
                            passingRollupListObject[2] = largeListCounter.ToString();
                            passingRollupListObject[3] = unlimitedVerCounter.ToString();
                            passingRollupListObject[4] = infoPathFormCounter.ToString();
                            passingRollupListObject[5] = infoPathExternalConnCounter.ToString();
                            WriteReports.WriteText(passingRollupListObject);
                        }
                        catch (WebException webException)
                        {
                            HttpWebResponse errorResponse = webException.Response as HttpWebResponse;
                            // If the error code from the attempt is a 404 or similar, inform the user that the site doesn't exist or is unreachable
                            if (!string.IsNullOrEmpty(errorResponse.ToString()) &&
                                errorResponse.StatusCode == HttpStatusCode.NotFound)
                            {
                                Console.WriteLine(@"Could not find the site at the address provided: {0}", currentLine);
                            }
                            // If the error code from the attempt is a 401 or similar, inform the user that they are not authorized
                            if (!string.IsNullOrEmpty(errorResponse.ToString()) &&
                                errorResponse.StatusCode == HttpStatusCode.Unauthorized)
                            {
                                Console.WriteLine(
                                    @"You do not have permissions with your current credentials for the site at this address: {0}",
                                    currentLine);
                            }
                        }
                    }
                }
                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", sitesFilePath);
            }
            // Run the method to inventory WORKFLOW DEFINITION COUNTS
            if (action.ToLower() == "standard")
            {
                // Write the standard report CSV Header
                string detailedWorkflowReportPath = filePath.Replace("DetailedWorkflow", "StandardWorkflow");
                streamWriter.Close();
                System.IO.File.Move(filePath, detailedWorkflowReportPath);
                // Write the rollup CSV Header
                string rollupWorkflowReportPath = filePath.Replace("DetailedWorkflow", "RollupStandardWorkflow");
                // Create a string of arguments to run the inventory function
                string[] passingStandardWfHeaderArgs = new string[2];
                passingStandardWfHeaderArgs[0] = action;
                passingStandardWfHeaderArgs[1] = filePath;
                BuildHeaders.WriteReportHeaders(passingStandardWfHeaderArgs);

                // run the inventory function one time
                string inputFile = siteListFilePath + "SiteList.csv";
                // Create a string of arguments to run the inventory function
                string[] passingUserArgs = new string[3];
                passingUserArgs[0] = inputFile;
                passingUserArgs[1] = detailedWorkflowReportPath;
                passingUserArgs[2] = rollupWorkflowReportPath;
                ChooseReport.RunWorkflowInventory(passingUserArgs, actingUser);

                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", detailedWorkflowReportPath);
            }
            // Run the method to inventory AD USERS
            if (action.ToLower() == "users")
            {
                // Write the CSV Header
                string detailedUserReportPath = filePath.Replace("DetailedWorkflow", "User");
                streamWriter.Close();
                System.IO.File.Move(filePath, detailedUserReportPath);
                // Create a string of arguments to run the inventory function
                string[] passingUserHeaderArgs = new string[2];
                passingUserHeaderArgs[0] = action;
                passingUserHeaderArgs[1] = filePath;
                BuildHeaders.WriteReportHeaders(passingUserHeaderArgs);

                // run the inventory function one time
                string inputFile = siteListFilePath + "SiteList.csv";
                // Create a string of arguments to run the inventory function
                string[] passingUserArgs = new string[3];
                passingUserArgs[0] = inputFile;
                passingUserArgs[1] = action;
                passingUserArgs[2] = detailedUserReportPath;
                ChooseReport.RunUserInventory(passingUserArgs, actingUser);

                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", detailedUserReportPath);
            }
            // Run the method to inventory UNLIMITED VERSION SETTINGS for lists
            if (action.ToLower() == "versions")
            {
                // Write the CSV Header
                string detailedListReportPath = filePath.Replace("DetailedWorkflow", "ListItemVersions");
                streamWriter.Close();
                System.IO.File.Move(filePath, detailedListReportPath);
                // Write the rollup CSV Header
                string rollupListReportPath = filePath.Replace("DetailedWorkflow", "RollupVersions");
                // Create a string of arguments to run the inventory function
                string[] passingVersionHeaderArgs = new string[2];
                passingVersionHeaderArgs[0] = action;
                passingVersionHeaderArgs[1] = filePath;
                BuildHeaders.WriteReportHeaders(passingVersionHeaderArgs);

                // run the inventory function one time
                string inputFile = siteListFilePath + "SiteList.csv";
                // Create a string of arguments to run the inventory function
                string[] passingListVersionArgs = new string[3];
                passingListVersionArgs[0] = inputFile;
                passingListVersionArgs[1] = detailedListReportPath;
                passingListVersionArgs[2] = rollupListReportPath;
                ChooseReport.RunListInventory(passingListVersionArgs, actingUser);

                // tell the user the file has been created
                Utilities.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", detailedListReportPath);
            }
            // Run the method to inventory webs and sandbox solutions instead of workflows and instances
            if (action.ToLower() == "webs")
            {
                // Write the CSV Headers
                // Create the detailed CSV file
                string detailedWebsReportPath = filePath.Replace("DetailedWorkflow", "Webs");
                streamWriter.Close();
                System.IO.File.Move(filePath, detailedWebsReportPath);
                // Create the rollup CSV File
                string rollupWebsReportPath = filePath.Replace("DetailedWorkflow", "RollupWebs");
                // Create a string of arguments to run the inventory function
                string[] passingWebsHeaderArgs = new string[2];
                passingWebsHeaderArgs[0] = action;
                passingWebsHeaderArgs[1] = filePath;
                BuildHeaders.WriteReportHeaders(passingWebsHeaderArgs);

                // Run the inventory Function
                string inputFile = siteListFilePath + "SiteList.csv";
                // Create a string of arguments to run the inventory function
                string[] passingWebArgs = new string[3];
                passingWebArgs[0] = inputFile;
                passingWebArgs[1] = detailedWebsReportPath;
                passingWebArgs[2] = rollupWebsReportPath;
                ChooseReport.RunWebInventory(passingWebArgs, actingUser);

                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Report Generated at {0}.", detailedWebsReportPath);
            }
            if (action.ToLower() != "webs" && action.ToLower() != "detailed" && action.ToLower() != "standard" && action.ToLower() != "versions" && action.ToLower() != "groups" && action.ToLower() != "infopath" && action.ToLower() != "users")
            {
                // tell the user the file has been created
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"There was something wrong with the action you chose, and a report could not be completely generated at {0}.", filePath);
            }
        }

        // For each URL Write Counters to update the SharePoint GCCMigrationSites Inventory List
        static void WriteSiteObjectToInvList(string siteUrl, ref int templateCounter, ref int solutionCounter, ref int masterPageCounter, ref int pageLayoutCounter, ref int appCounter, ref int dropoffCounter, ref int listTemplateCounter, string inventoryWebAddress)
        {
            string inventoryListTitle = "MNIT Site Collections";
            using (var invContext = new ClientContext(inventoryWebAddress))
            {
                invContext.Credentials = CredentialCache.DefaultCredentials;
                // Get web and web information from sharepoint site
                invContext.Load(invContext.Web);
                invContext.ExecuteQuery();
                // Get the list
                List currentList = invContext.Web.Lists.GetByTitle(inventoryListTitle);
                CamlQuery csvItemQuery = new CamlQuery();
                csvItemQuery.ViewXml = @"<View><Query><Where>" +
                "<Eq><FieldRef Name='FullURL' /><Value Type='Text'>" + siteUrl + "</Value></Eq>" +
                "</Where></Query></View>";
                ListItemCollection items = currentList.GetItems(csvItemQuery);
                // Load list items
                invContext.Load(items);
                // Execute Query against list items
                invContext.ExecuteQuery();

                foreach (ListItem listItem in items)
                {
                    invContext.Load(listItem, i => i.Id, i => i["FullURL"], i => i["TemplateCount"]);
                    invContext.ExecuteQuery();
                    listItem["TemplateCount"] = templateCounter;
                    listItem["SandboxSolutions"] = solutionCounter;
                    listItem["CustomMasterPageCount"] = masterPageCounter;
                    listItem["AppCount"] = appCounter;
                    listItem["DropOffLibraryCount"] = dropoffCounter;
                    listItem.Update();
                    invContext.ExecuteQuery();
                }
            }
        }

        // For each URL Write Counters to update the SharePoint GCCMigrationSites Inventory List
        static void WriteDetailedWorkflowObjectToInvList(string siteUrl, ref int runningInstancesCounter, string inventoryWebAddress)
        {
            string inventoryListTitle = "MNIT Site Collections";
            using (var invContext = new ClientContext(inventoryWebAddress))
            {
                invContext.Credentials = CredentialCache.DefaultCredentials;
                // Get web and web information from sharepoint site
                invContext.Load(invContext.Web);
                invContext.ExecuteQuery();
                // Get the list
                List currentList = invContext.Web.Lists.GetByTitle(inventoryListTitle);
                CamlQuery csvItemQuery = new CamlQuery();
                csvItemQuery.ViewXml = @"<View><Query><Where>" +
                "<Eq><FieldRef Name='FullURL' /><Value Type='Text'>" + siteUrl + "</Value></Eq>" +
                "</Where></Query></View>";
                ListItemCollection items = currentList.GetItems(csvItemQuery);
                // Load list items
                invContext.Load(items);
                // Execute Query against list items
                invContext.ExecuteQuery();
                foreach (ListItem listItem in items)
                {
                    invContext.Load(listItem, i => i.Id, i => i["FullURL"], i => i["RunningInstancesCount"]);
                    invContext.ExecuteQuery();
                    listItem["RunningInstancesCount"] = runningInstancesCounter;
                    listItem.Update();
                    invContext.ExecuteQuery();
                }
            }
        }

        // For each URL Write Counters to update the SharePoint GCCMigrationSites Inventory List
        static void WriteStandardWorkflowObjectToInvList(string siteUrl, ref int nintexCounter, ref int spd2010Counter, ref int spd2013Counter, string inventoryWebAddress)
        {
            string inventoryListTitle = "MNIT Site Collections";
            using (var invContext = new ClientContext(inventoryWebAddress))
            {
                invContext.Credentials = CredentialCache.DefaultCredentials;
                // Get web and web information from sharepoint site
                invContext.Load(invContext.Web);
                invContext.ExecuteQuery();
                // Get the list
                List currentList = invContext.Web.Lists.GetByTitle(inventoryListTitle);
                CamlQuery csvItemQuery = new CamlQuery();
                csvItemQuery.ViewXml = @"<View><Query><Where>" +
                "<Eq><FieldRef Name=' FullURL' /><Value Type='Text'>" + siteUrl + "</Value></Eq>" +
                "</Where></Query></View>";
                ListItemCollection items = currentList.GetItems(csvItemQuery);
                // Load list items
                invContext.Load(items);
                // Execute Query against list items
                invContext.ExecuteQuery();
                foreach (ListItem listItem in items)
                {
                    invContext.Load(listItem, i => i.Id, i => i[" FullURL"]);
                    invContext.ExecuteQuery();
                    listItem["NintexCount"] = nintexCounter;
                    listItem["SPD2010Count"] = spd2010Counter;
                    listItem["SPD2013Count"] = spd2013Counter;
                    listItem.Update();
                    invContext.ExecuteQuery();
                }
            }
        }

        // Write to SP List
        static void WriteListVersionObjectToInvList(string siteUrl, ref int largeListCounter, ref int unlimitedVerCounter, string inventoryWebAddress)
        {
            string inventoryListTitle = "MNIT Site Collections";
            using (var invContext = new ClientContext(inventoryWebAddress))
            {
                invContext.Credentials = CredentialCache.DefaultCredentials;
                // Get web and web information from sharepoint site
                invContext.Load(invContext.Web);
                invContext.ExecuteQuery();
                // Get the list
                List currentList = invContext.Web.Lists.GetByTitle(inventoryListTitle);
                CamlQuery csvItemQuery = new CamlQuery();
                csvItemQuery.ViewXml = @"<View><Query><Where>" +
                "<Eq><FieldRef Name='FullURL' /><Value Type='Text'>" + siteUrl + "</Value></Eq>" +
                "</Where></Query></View>";
                ListItemCollection items = currentList.GetItems(csvItemQuery);
                // Load list items
                invContext.Load(items);
                // Execute Query against list items
                invContext.ExecuteQuery();
                foreach (ListItem listItem in items)
                {
                    invContext.Load(listItem, i => i.Id, i => i["FullURL"]);
                    invContext.ExecuteQuery();
                    listItem["Large List Count"] = largeListCounter;
                    listItem["UnlimitedVersionsCount"] = unlimitedVerCounter;
                    listItem.Update();
                    invContext.ExecuteQuery();
                }
            }
        }

        // For each URL Write Counters to update the SharePoint GCCMigrationSites Inventory List
        static void WriteInfoPathObjectToInvList(string siteUrl, ref int infoPathFormCounter, ref int infoPathExternalConnCounter, string inventoryWebAddress)
        {
            string inventoryListTitle = "MNIT Site Collections";
            using (var invContext = new ClientContext(inventoryWebAddress))
            {
                invContext.Credentials = CredentialCache.DefaultCredentials;
                // Get web and web information from sharepoint site
                invContext.Load(invContext.Web);
                invContext.ExecuteQuery();
                // Get the list
                List currentList = invContext.Web.Lists.GetByTitle(inventoryListTitle);
                CamlQuery csvItemQuery = new CamlQuery();
                csvItemQuery.ViewXml = @"<View><Query><Where>" +
                "<Eq><FieldRef Name='FullURL' /><Value Type='Text'>" + siteUrl + "</Value></Eq>" +
                "</Where></Query></View>";
                ListItemCollection items = currentList.GetItems(csvItemQuery);
                // Load list items
                invContext.Load(items);
                // Execute Query against list items
                invContext.ExecuteQuery();
                foreach (ListItem listItem in items)
                {
                    invContext.Load(listItem, i => i.Id, i => i["FullURL"]);
                    invContext.ExecuteQuery();
                    listItem["IPFormLibraryCount"] = infoPathFormCounter;
                    listItem["IPWithExtConnCount"] = infoPathExternalConnCounter;
                    listItem.Update();
                    invContext.ExecuteQuery();
                }
            }
        }
        // For each URL write Counters to a CSV to rollup the information per site collection
    }
}
