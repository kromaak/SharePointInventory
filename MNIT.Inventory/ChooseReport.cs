﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

using Utils = MNIT.Utilities;

namespace MNIT.Inventory
{
    class ChooseReport
    {
        // Runs Group Inventory
        public static void RunGroupInventory(string[] args, Utils.ActingUser actingUser)
        {
            // 0 = input file
            string inputFile = args[0];
            // 1 = users or groups
            string action = args[1];
            // 3 = detailed report file
            string detailedGroupReportPath = args[2];
            // Read through the list of sites
            //string[] readUrls = null;

            //
            // Read in a file line-by-line, and store it all in a List.
            //
            List<string> list = new List<string>();
            try
            {
                using (StreamReader reader = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        list.Add(line); // Add to list.
                        //Console.WriteLine(line); // Write to console.
                    }
                }
            }
            //try
            //{
            //    readUrls = System.IO.File.ReadAllLines(inputFile);
            //}
            catch (Exception ex31Exception)
            {
                Console.WriteLine(ex31Exception.Message);
            }
            //string[] readUrls = System.IO.File.ReadAllLines(siteListFilePath + "SiteList.csv");
            int everyTen = 0;
            // For each site address in the CSV file
            //foreach (string readCurrentLine in readUrls)
            foreach (string readCurrentLine in list)
            {
                if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                {
                    string currentLine = readCurrentLine.Trim();
                    everyTen++;
                    try
                    {
                        //string adGroups = "\"";
                        string adGroups = "";
                        string permissionLevels = "";
                        // Write the site URL every ten lines from CSV, to let the user know progress is being made
                        if (everyTen % 10 == 0 && everyTen != 0)
                        {
                            Utils.SpinAnimation.Stop();
                            Console.WriteLine();
                            Console.WriteLine(@"Getting AD groups for the address provided: {0}", currentLine);
                            Utils.SpinAnimation.Start();
                        }
                        // Run the SP Inventory AD Group function 
                        GetUsersAndGroups.InventoryAdGroups(currentLine, actingUser, ref adGroups, ref permissionLevels, action, detailedGroupReportPath);
                        // Write the Inventory information to the stream
                        string[] passingGroupObject = new string[4];
                        passingGroupObject[0] = detailedGroupReportPath;
                        passingGroupObject[1] = currentLine;
                        passingGroupObject[2] = adGroups;
                        passingGroupObject[3] = permissionLevels;
                        WriteReports.WriteText(passingGroupObject);
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
        }

        // Run InfoPath Inventory
        public static void RunInfoPathInventory(string[] args, Utils.ActingUser actingUser)
        {
            // 0 = input file
            string inputFile = args[0];
            // 1 = detailed report file
            string detailedInfoPathReportPath = args[1];
            // 2 = rollup report file
            string rollupInfoPathReportPath = args[2];
            // Read through the list of sites
            //string[] readUrls = null;

            //
            // Read in a file line-by-line, and store it all in a List.
            //
            List<string> list = new List<string>();
            try
            {
                using (StreamReader reader = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        list.Add(line); // Add to list.
                        //Console.WriteLine(line); // Write to console.
                    }
                }
            }
            //try
            //{
            //    readUrls = System.IO.File.ReadAllLines(inputFile);
            //}
            catch (Exception ex31Exception)
            {
                Console.WriteLine(ex31Exception.Message);
            }
            int everyTen = 0;
            //foreach (string readCurrentLine in readUrls)
            foreach (string readCurrentLine in list)
            {
                if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                {
                    string currentLine = readCurrentLine.Trim();
                    everyTen++;
                    // Counter Variables
                    int infoPathFormCounter = 0;
                    int infoPathExternalConnCounter = 0;
                    // Run the inventory function for Basic Workflow Information
                    try
                    {
                        // Write the site URL every ten lines from CSV, to let the user know progress is being made
                        if (everyTen % 10 == 0 && everyTen != 0)
                        {
                            Utils.SpinAnimation.Stop();
                            Console.WriteLine();
                            Console.WriteLine(@"Getting InfoPath Form Library info for the address provided: {0}",
                                currentLine);
                            Utils.SpinAnimation.Start();
                        }
                        GetInfoPath.InventoryInfoPath(currentLine, actingUser, ref infoPathFormCounter, ref infoPathExternalConnCounter, detailedInfoPathReportPath);
                        //// Run the SP Inventory List Update function
                        //if (args.Length > 0 && args[0] == "i")
                        //{
                        //    if (args[1].Length > 0 && args[1].Contains("http"))
                        //    {
                        //        WriteInfoPathObjectToInvList(currentLine, ref infoPathFormCounter,
                        //            ref infoPathExternalConnCounter, args[1]);
                        //    }
                        //}
                        // Run the rolled up Inventory function
                        string[] passingRollupIpObject = new string[4];
                        passingRollupIpObject[0] = rollupInfoPathReportPath;
                        passingRollupIpObject[1] = currentLine;
                        passingRollupIpObject[2] = infoPathFormCounter.ToString();
                        passingRollupIpObject[3] = infoPathExternalConnCounter.ToString();
                        WriteReports.WriteText(passingRollupIpObject);
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
        }

        // Run Lists Inventory
        public static void RunListInventory(string[] args, Utils.ActingUser actingUser)
        {
            // 0 = input file
            string inputFile = args[0];
            // 1 = detailed report file
            string detailedListReportPath = args[1];
            // 2 = rollup report file
            string rollupListReportPath = args[2];
            // Read through the list of sites
            //string[] readUrls = null;

            //
            // Read in a file line-by-line, and store it all in a List.
            //
            List<string> list = new List<string>();
            try
            {
                using (StreamReader reader = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        list.Add(line);
                    }
                }
            }
            //try
            //{
            //    readUrls = System.IO.File.ReadAllLines(inputFile);
            //}
            catch (Exception ex31Exception)
            {
                Console.WriteLine(ex31Exception.Message);
            }
            int everyTen = 0;
            // For each site address in the CSV file
            //if (readUrls != null)
            //{
            //foreach (string readCurrentLine in readUrls)
            foreach (string readCurrentLine in list)
                {
                    if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                    {
                        string currentLine = readCurrentLine.Trim();
                        everyTen++;
                        // Counter Variables
                        int largeListCounter = 0;
                        int unlimitedVerCounter = 0;
                        int siteCollCheckedOut = 0;
                        // Run the inventory function for List Information
                        try
                        {
                            // Write the site URL every 5 lines from CSV to let the user know progress is being made
                            if (everyTen % 10 == 0 && everyTen != 0)
                            {
                                Utilities.SpinAnimation.Stop();
                                Console.WriteLine();
                                Console.WriteLine(
                                    @"Getting List size and versioning info for the address provided: {0}", currentLine);
                                Utilities.SpinAnimation.Start();
                            }
                            GetVersions.InventoryVersions(currentLine, actingUser, ref largeListCounter,ref unlimitedVerCounter, ref siteCollCheckedOut, detailedListReportPath);

                            //// Run the SP Inventory List Update function
                            //if (args.Length > 0 && args[0] == "i")
                            //{
                            //    if (args[1].Length > 0 && args[1].Contains("http"))
                            //    {
                            //        WriteListVersionObjectToInvList(currentLine, ref largeListCounter,
                            //            ref unlimitedVerCounter, args[1]);
                            //    }
                            //}
                            // Run the rolled up Inventory function
                            string[] passingLvRollupObject = new string[5];
                            passingLvRollupObject[0] = rollupListReportPath;
                            passingLvRollupObject[1] = currentLine;
                            passingLvRollupObject[2] = largeListCounter.ToString();
                            passingLvRollupObject[3] = unlimitedVerCounter.ToString();
                            passingLvRollupObject[4] = siteCollCheckedOut.ToString();
                            WriteReports.WriteText(passingLvRollupObject);
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
            //}
        }

        // Runs User Inventory
        public static void RunUserInventory(string[] args, Utils.ActingUser actingUser)
        {
            // 0 = input file
            string inputFile = args[0];
            // 1 = users or groups
            string action = args[1];
            // 3 = detailed report file
            string detailedUserReportPath = args[2];
            // Read through the list of sites
            //string[] readUrls = null;

            //
            // Read in a file line-by-line, and store it all in a List.
            //
            List<string> list = new List<string>();
            try
            {
                using (StreamReader reader = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        list.Add(line);
                    }
                }
            }
            //try
            //{
            //    readUrls = System.IO.File.ReadAllLines(inputFile);
            //}
            catch (Exception ex31Exception)
            {
                Console.WriteLine(ex31Exception.Message);
            }
            //string[] readUrls = System.IO.File.ReadAllLines(siteListFilePath + "SiteList.csv");
            int everyTen = 0;
            // For each site address in the CSV file
            //foreach (string readCurrentLine in readUrls)
            foreach (string readCurrentLine in list)
            {
                if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                {
                    string siteCollId = "";
                    string currentLine = readCurrentLine.Trim();
                    everyTen++;
                    try
                    {
                        // 0 out the reference variables
                        string users = "";
                        string permissionLevels = "";
                        // Write the site URL every ten lines from CSV, to let the user know progress is being made
                        if (everyTen % 10 == 0 && everyTen != 0)
                        {
                            Utils.SpinAnimation.Stop();
                            Console.WriteLine();
                            Console.WriteLine(@"Getting Users for the address provided: {0}", currentLine);
                            Utils.SpinAnimation.Start();
                        }
                        // Run the SP Inventory AD Group function 
                        GetUsersAndGroups.InventoryAdGroups(currentLine, actingUser, ref users, ref permissionLevels, action, detailedUserReportPath);
                        // Write the Inventory information to the stream
                        string[] passingUserObjects = new string[5];
                        passingUserObjects[0] = detailedUserReportPath;
                        passingUserObjects[1] = siteCollId;
                        passingUserObjects[2] = currentLine;
                        passingUserObjects[3] = users;
                        passingUserObjects[4] = permissionLevels;
                        WriteReports.WriteText(passingUserObjects);
                    }
                    catch (WebException webException)
                    {
                        if (webException.Status == WebExceptionStatus.ProtocolError && webException.Response != null)
                        {
                            HttpWebResponse errorResponse = webException.Response as HttpWebResponse;
                            // If the error code from the attempt is a 404 or similar, inform the user that the site doesn't exist or is unreachable
                            if (!string.IsNullOrEmpty(errorResponse.ToString()) &&
                                errorResponse.StatusCode == HttpStatusCode.NotFound)
                            {
                                Console.WriteLine(@"Could not find the site at the address provided: {0}",
                                    currentLine);
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
                        else
                        {
                            Console.WriteLine("No valid response from the site!");
                        }
                    }
                }
            }
        }

        // Run Webs Inventory
        public static void RunWebInventory(string[] args, Utils.ActingUser actingUser)
        {
            // 0 = input file
            string inputFile = args[0];
            // 1 = detailed report file
            string detailedWebsReportPath = args[1];
            // 2 = rollup report file
            string rollupWebsReportPath = args[2];
            List<string> list = new List<string>();
            try
            {
                using (StreamReader reader = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        list.Add(line);
                    }
                }
            }
            catch (Exception ex41Exception)
            {
                Console.WriteLine(ex41Exception.Message);
            }

            // For each site address in the CSV file
            int everyTen = 0;
            foreach (string readCurrentLine in list)
            {
                if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                {
                    string currentLine = readCurrentLine.Trim();
                    everyTen++;
                    int siteTemplateCounter = 0;
                    int solutionCounter = 0;
                    int masterPageCounter = 0;
                    int pageLayoutCounter = 0;
                    int customPageCounter = 0;
                    int appCounter = 0;
                    int dropoffCounter = 0;
                    int listTemplateCounter = 0;
                    int exportedWpCounter = 0;
                    // Run the inventory function for WEB Information
                    try
                    {
                        // Write the site URL every ten lines from CSV, to let the user know progress is being made
                        if (everyTen % 10 == 0 && everyTen != 0)
                        {
                            Utils.SpinAnimation.Stop();
                            Console.WriteLine();
                            Console.WriteLine(@"Getting Site info for the address provided: {0}", currentLine);
                            Utils.SpinAnimation.Start();
                        }
                        //// Check for arguments
                        //if (args.Length > 0)
                        //{
                        //    for (int a = 0; a < args.Length; a++)
                        //    {
                        //        // if there is an argument, and it specifies e for exported web parts, and there is a string following it,
                        //        // run the InventoryWebs function, passing the exp WP string as an arg
                        //        if (args[a] == "e" && args[a + 1].Length > 0)
                        //        {
                        //            GetWebs.InventoryWebs(currentLine, args[a + 1], actingUser,
                        //                ref siteTemplateCounter,
                        //                ref solutionCounter,
                        //                ref masterPageCounter, ref pageLayoutCounter, ref customPageCounter, ref appCounter,
                        //                ref dropoffCounter, ref listTemplateCounter, ref exportedWpCounter, detailedWebsReportPath);
                        //        }
                        //        //// if there is an argument, and it specifies i for inventory in a SharePoint list,
                        //        //// run the Write function to update counts for complex objects in the SharePoint list specified
                        //        //if (args[a] == "i")
                        //        //{
                        //        //    if (args[a + 1].Length > 0 && args[a + 1].Contains("http"))
                        //        //    {
                        //        //        WriteSiteObjectToInvList(currentLine, ref siteTemplateCounter, ref solutionCounter,
                        //        //            ref masterPageCounter, ref pageLayoutCounter, ref appCounter, ref dropoffCounter, ref listTemplateCounter, args[1]);
                        //        //    }
                        //        //}
                        //    }
                        //}
                        //else
                        //{
                            GetWebs.InventoryWebs(currentLine, null, actingUser, ref siteTemplateCounter,
                                ref solutionCounter,
                                ref masterPageCounter, ref pageLayoutCounter, ref customPageCounter, ref appCounter, ref dropoffCounter, ref listTemplateCounter, ref exportedWpCounter, detailedWebsReportPath);
                        //}
                        // Write the counters to the site collection rollup report
                        string[] passingWebRollupObject = new string[11];
                        passingWebRollupObject[0] = rollupWebsReportPath;
                        passingWebRollupObject[1] = currentLine;
                        passingWebRollupObject[2] = siteTemplateCounter.ToString();
                        passingWebRollupObject[3] = solutionCounter.ToString();
                        passingWebRollupObject[4] = masterPageCounter.ToString();
                        passingWebRollupObject[5] = pageLayoutCounter.ToString();
                        passingWebRollupObject[6] = customPageCounter.ToString();
                        passingWebRollupObject[7] = appCounter.ToString();
                        passingWebRollupObject[8] = dropoffCounter.ToString();
                        passingWebRollupObject[9] = listTemplateCounter.ToString();
                        passingWebRollupObject[10] = exportedWpCounter > 0 ? exportedWpCounter.ToString() : "";
                        //passingWebRollupObject[10] = "";
                        WriteReports.WriteText(passingWebRollupObject);
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
        }

        // Run Workflow Inventory
        public static void RunWorkflowInventory(string[] args, Utils.ActingUser actingUser)
        {
            // 0 = input file
            string inputFile = args[0];
            // 1 = detailed report file
            string detailedWorkflowReportPath = args[1];
            // 2 = rollup report file
            string rollupWorkflowReportPath = args[2];
            List<string> list = new List<string>();
            try
            {
                using (StreamReader reader = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        list.Add(line);
                    }
                }
            }
            catch (Exception ex31Exception)
            {
                Console.WriteLine(ex31Exception.Message);
            }
            int everyTen = 0;
            foreach (string readCurrentLine in list)
            {
                everyTen++;
                // Counter Variables
                int nintexCounter = 0;
                int spd2010Counter = 0;
                int spd2013Counter = 0;
                if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                {
                    string currentLine = readCurrentLine.Trim();
                    // Run the inventory function for Basic Workflow Information
                    try
                    {
                        // Write the site URL every ten lines from CSV, to let the user know progress is being made
                        if (everyTen % 10 == 0 && everyTen != 0)
                        {
                            Utils.SpinAnimation.Stop();
                            Console.WriteLine();
                            Console.WriteLine(@"Getting Workflow info for the address provided: {0}", currentLine);
                            Utils.SpinAnimation.Start();
                        }
                        GetStandardWorkflows.InventoryWorkflowsStandard(currentLine, actingUser, ref nintexCounter,
                            ref spd2010Counter, ref spd2013Counter, detailedWorkflowReportPath);
                        //// Run the SP Inventory List Update function
                        //if (args.Length > 0 && args[0] == "i")
                        //{
                        //    if (args[1].Length > 0 && args[1].Contains("http"))
                        //    {
                        //        WriteStandardWorkflowObjectToInvList(currentLine, ref nintexCounter,
                        //            ref spd2010Counter,
                        //            ref spd2013Counter, args[1]);
                        //    }
                        //}
                        // Run the rolled up Inventory function
                        string[] passingWfRollupObject = new string[5];
                        passingWfRollupObject[0] = rollupWorkflowReportPath;
                        passingWfRollupObject[1] = currentLine;
                        passingWfRollupObject[2] = nintexCounter.ToString();
                        passingWfRollupObject[3] = spd2010Counter.ToString();
                        passingWfRollupObject[4] = spd2013Counter.ToString();
                        WriteReports.WriteText(passingWfRollupObject);
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

        }
    }
}
