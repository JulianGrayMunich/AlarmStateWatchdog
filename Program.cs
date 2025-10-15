using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Reflection.Metadata.Ecma335;
using System.Security.Claims;
using System.Text.RegularExpressions;

using databaseAPI;

using EASendMail;

using GNA_CommercialLicenseValidator;

using gnaDataClasses;

using GNAgeneraltools;

using GNAspreadsheettools;

using Microsoft.Extensions.Configuration;

using OfficeOpenXml;

using Twilio;
using Twilio.Rest.Api.V2010.Account;

using static System.Runtime.InteropServices.JavaScript.JSType;


#pragma warning disable CS0162
#pragma warning disable CS0164
#pragma warning disable CS0219
#pragma warning disable CS8600
#pragma warning disable CS8602
#pragma warning disable CS8604


namespace AlarmStateWatchdog
{
    class Program
    {
        static void Main()
        {
            try
            {
            gnaTools gnaT = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();
            gnaDataClass gnaDC = new();

            #region Validate Config File
            Console.WriteLine("Parsing local config file");   
            gnaT.VerifyLocalConfig();
            #endregion

            #region DB Connection
            string strDBconnection = System.Configuration.ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;
            #endregion


            #region Console
            string strTab1 = "     ";
            string strTab2 = "        ";
            string strTab3 = "           ";
            string strTab4 = "              ";
            #endregion

            #region Config Variables
            var config = System.Configuration.ConfigurationManager.AppSettings;
            string licenseCode = config["LicenseCode"];
            #endregion


            #region Yes/No settings
            string strFreezeScreen = config["freezeScreen"];
            string strComputedRdT = config["computedRdT"];
            string strSendEmails = config["SendEmails"];
            #endregion


            #region Operational Settings
            string strFirstDataRow = config["FirstDataRow"];
            double dblAlarmWindowHrs = Convert.ToDouble(config["AlarmWindowHrs"]);
            int iNoOfSuccessfulReadings = Convert.ToInt16(config["NoOfSuccessfulReadings"]);
            string strDailyStatusReportTime = config["dailyStatusReportTime"];
            #endregion


            #region Project & Contract Info
            string strProjectTitle = config["ProjectTitle"];
            string strContractTitle = config["ContractTitle"];
            string strSMSTitle = config["SMSTitle"];
            #endregion


            #region Excel Configuration
            string strExcelPath = config["ExcelPath"];
            string strExcelFile = config["ExcelFile"];
            string strReferenceWorksheet = config["ReferenceWorksheet"];
            string strSurveyWorksheet = config["SurveyWorksheet"];
            #endregion

            #region System Folders
            string strStatusFolder = config["SystemStatusFolder"];
            string strAlarmsFolder = config["SystemAlarmsFolder"];
            #endregion

            #region SMS Recipients
            List<string> smsMobile = new();
            List<string> invalidMobiles = new();

            Regex e164 = new(@"^\+\d{8,15}$");

            for (int i = 1; i <= 9; i++)
            {
                string? raw = config[$"RecipientPhone{i}"];
                if (string.IsNullOrWhiteSpace(raw)) continue;

                string phone = raw.Trim();

                if (e164.IsMatch(phone))
                    smsMobile.Add(phone);
                else
                    invalidMobiles.Add(phone);
            }

            // Remove duplicates while preserving order
            smsMobile = smsMobile
                .Distinct(StringComparer.Ordinal)
                .ToList();

            // Comma-delimited string of valid, unique numbers
            string strMobileNumbers = string.Join(",", smsMobile);
            #endregion

            #region Email Settings
            string strEmailLogin = config["EmailLogin"];
            string strEmailPassword = config["EmailPassword"];
            string strEmailFrom = config["EmailFrom"];
            string strEmailRecipients = config["EmailRecipients"];

            EmailCredentials emailCreds = gnaT.BuildEmailCredentials(
                strEmailLogin,
                strEmailPassword,
                strEmailFrom,
                strEmailRecipients);
            #endregion

            #region Variables
            string strATSname = "";
            string strSettop = "";
            string strATSexists = "No";
            string strExitFlag = "No";
            string strInformation = "";
            string strAlarmFile = "NoDataAlarmState.txt";
            string alarmLog = strAlarmsFolder + strAlarmFile;
            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
            #endregion



            #region Environment Check
            try
            {


                //==== Console settings
                Console.OutputEncoding = System.Text.Encoding.Unicode;
                //CultureInfo culture;

                #region Licenses
                Console.WriteLine("Validating the software license...");
                LicenseValidator.ValidateLicense("ALSTWD", licenseCode);
                Console.WriteLine(strTab1 + "Validated");
                ExcelPackage.License.SetCommercial("14XO1NhmOmVcqDWhA0elxM72um6vnYOS8UiExVFROZuRPn1Ddv5fRV8fiCPcjujkdw9H18nExINNFc8nmOjRIQEGQzVDRjMz5wdPAJkEAQEA");  //valid to 23.03.2026
                #endregion


                gnaT.WelcomeMessage($"AlarmStateWatchdog {BuildInfo.BuildDateString()}");

                string strNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                Console.WriteLine("");
                Console.WriteLine("1. Check system environment");
                Console.WriteLine(strTab1 + "Project: " + strProjectTitle);
                Console.WriteLine(strTab1 + "Master workbook: " + strMasterWorkbookFullPath);

                if (strFreezeScreen == "Yes")
                {
                    gnaDBAPI.testDBconnection(strDBconnection);

                    string strProjectID = gnaDBAPI.getProjectID(strDBconnection, strProjectTitle);
                    if (strProjectID == "Missing")
                    {
                        Console.WriteLine("\n" + strTab1 + "**** " + strProjectTitle + " is missing ****");
                        goto ThatsAllFolks;
                    }
                    else
                    {
                        Console.WriteLine(strTab1 + strProjectTitle + " found");
                    }

                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                }
                else
                {
                    Console.WriteLine(strTab1 + "Existance of workbook & worksheets is not checked");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError: \n" + ex.Message);
            }
            #endregion

            #region Previous Alarm State 
            // read the ats and settop from the config file and write to AlarmState
            Console.WriteLine("2. Extract ATS and settop data");
            Console.WriteLine(strTab1 + "Extract from config file");
            // instantiate the list
            var alarmstate = new List<AlarmState>();

            for (int ATScounter = 1; ATScounter <= 9; ATScounter++)
            {
                string strATSdata = config[$"ATSdetails{ATScounter}"] ?? string.Empty;
                if (strATSdata != "")
                {

                    // Split into fields
                    string[] parts = strATSdata.Split(',');
                    if (parts.Length >= 5) // safety: need at least 5 elements
                    {
                        // Create and populate a new AlarmState
                        var state = new AlarmState
                        {
                            ATSname = parts[0],    // first element
                            Settop = parts[4]  // fifth element
                        };
                        // Add to the list
                        alarmstate.Add(state);
                    }
                }
                else
                {
                    continue;
                }
            }

            Console.WriteLine($"{strTab1}Verify ATS list in {strExcelFile}");

            foreach (var state in alarmstate)
            {

                strATSname = state.ATSname;
                strATSexists = gnaSpreadsheetAPI.CheckStringInColumn(strExcelPath, strExcelFile, strSurveyWorksheet, "11", strATSname);
                if (strATSexists == "Missing")
                {
                    strExitFlag = "Yes";
                    strInformation = strInformation + strATSname + ",";
                }
            }

            if (strExitFlag == "Yes")
            {
                strInformation = strInformation.TrimEnd(',');
                Console.WriteLine($"\n{strTab2}The following ATS are missing from {strExcelFile}:");
                Console.WriteLine($"{strTab3} {strInformation}");
                goto ThatsAllFolks;
            }
            else
            {
                Console.WriteLine($"{strTab2}All ATS verified.");
            }

            // prepare the data for the alarm txt file.
            // Assume: var ats = new List<AlarmState>();
            string strDummyString = string.Empty;

            foreach (var ats in alarmstate)
            {
                strATSname = ats.ATSname ?? string.Empty;
                strSettop = ats.Settop ?? string.Empty;
                strDummyString = strDummyString + strATSname + "(" + strSettop + "):Alarm,";
            }

            // Trim the last comma
            strDummyString = strDummyString.TrimEnd(',');
            // Append date and time
            string strAlarmTime = "2000-01-01 00:01";
            strDummyString = strAlarmTime + "/" + strDummyString;

            if (!File.Exists(alarmLog))
            {
                using (var sw = new StreamWriter(alarmLog, false))
                {
                    sw.WriteLine(strDummyString);
                }
            }

            // extract the state of each ATS from the alarm file and write to AlarmState as "PreviousState"
            string strPreviousDateTime = gnaSpreadsheetAPI.obtainPreviousAlarmTime(alarmLog);
            alarmstate = gnaSpreadsheetAPI.obtainPreviousAlarmState(alarmLog, alarmstate);
            #endregion

            #region SensorID

            Console.WriteLine("3. Extract point names ");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("4. Extract SensorID");
            string[,] strSensors = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

            Console.WriteLine("5. Write SensorID to workbook");
            gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensors, strFirstDataRow);

            #endregion

            Console.WriteLine("6. Extract sensor info from workbook");

            var sensorList = gnaSpreadsheetAPI.getSensorInfo(strExcelPath, strExcelFile, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("7. Generate timeblock Start and End ");
            double dblStartTimeOffset = -1.0 * dblAlarmWindowHrs;
            string strLocalStartTime = DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm") + ":00";
            string strLocalEndTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm") + ":00";




            #region TestData
            // for testing only - comment out for live operation
            //strLocalStartTime = "'2025-05-24 00:00'";
            //strLocalEndTime = "'2025-05-24 06:00'";
            #endregion

            Console.WriteLine($"{strTab1}Start: {strLocalStartTime}");
            Console.WriteLine($"{strTab1}  End: {strLocalEndTime}");

            string strTimeBlockStartUTC = "'" + gnaT.convertLocalToUTC(strLocalStartTime) + "'";
            string strTimeBlockEndUTC = "'" + gnaT.convertLocalToUTC(strLocalEndTime) + "'";

            Console.WriteLine("8. Retrieve prism read state ");
            // sensorList assumed to be filled already from getSensorInfo()
            sensorList = gnaDBAPI.retrieveIsSensorRead(strDBconnection, sensorList, strTimeBlockStartUTC, strTimeBlockEndUTC);

            strAlarmTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            string strCurrentATSstate = strAlarmTime + "/";


            foreach (var state in alarmstate)
            {
                int yesCount = 0;
                int noCount = 0;

                foreach (var sensor in sensorList)
                {
                    if (sensor.ATS == state.ATSname)
                    {
                        if (sensor.Read == "Yes")
                        {
                            yesCount++;
                        }
                        else if (sensor.Read == "No")
                        {
                            noCount++;
                        }
                    }
                }

                // Assign totals to the AlarmState entry
                state.Yes = yesCount;
                state.No = noCount;

                if (state.Yes >= iNoOfSuccessfulReadings)
                {
                    state.currentAlarmState = "OK";
                }
                else
                {
                    state.currentAlarmState = "Alarm";
                }


                strCurrentATSstate = strCurrentATSstate + state.ATSname + "(" + state.Settop + "):" + state.currentAlarmState + ",";

                if (state.currentAlarmState == state.previousAlarmState)
                {
                    state.StateChange = "NoChange";
                }
                else
                {
                    state.StateChange = "Changed";
                }

            }


            strCurrentATSstate = strCurrentATSstate.TrimEnd(',');


            try
            {
                File.WriteAllText(alarmLog, strCurrentATSstate + Environment.NewLine);
                Console.WriteLine($"{strTab1}Alarm log updated.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{strTab1}Warning: failed to update alarm log: {ex.Message}");
            }


            //foreach (var state1 in alarmstate)
            //{
            //    Console.WriteLine(
            //        "ATSname: " + state1.ATSname +
            //        ", Settop: " + state1.Settop +
            //        ", Yes: " + state1.Yes +
            //        ", No: " + state1.No
            //        + ", PreviousState: " + state1.previousAlarmState +
            //        ", CurrentState: " + state1.currentAlarmState +
            //        ", StateChange: " + state1.StateChange
            //    );
            //}


            // === CONTINUE FROM HERE: build notifications and persist state ============================

            Console.WriteLine("9. Prepare messages");
            // Prepare the message
            string messageBalance = string.Empty;
            string strShortMessage = string.Empty;

            // Evaluate change/health conditions
            bool allNoChange = alarmstate.All(s => s.StateChange == "NoChange");
            bool anyChanged = alarmstate.Any(s => s.StateChange == "Changed");
            bool allCurrentOk = alarmstate.All(s => s.currentAlarmState == "OK");

            if (allNoChange)
            {
                // Case 1: nothing changed at all
                messageBalance = "No change";
            }
            else if (anyChanged && allCurrentOk)
            {
                // Case 2: at least one changed AND everyone is now OK
                messageBalance = "NoData alarm cancelled";
            }
            else
            {
                // Case 3: there are changes (but not all OK) OR any other mixed case
                // Requirement: list ALL ATS currently in Alarm (not only those that changed)
                var inAlarm = alarmstate.Where(s => s.currentAlarmState == "Alarm").ToList();

                if (inAlarm.Count == 0)
                {
                    // Safety fallback; theoretically covered by previous branch, but keep robust.
                    messageBalance = "No change";
                }
                else
                {
                    var sb = new System.Text.StringBuilder();
                    foreach (var s in inAlarm)
                    {
                        sb.AppendLine($"{s.ATSname}({s.Settop}):{s.currentAlarmState}");
                    }
                    messageBalance = sb.ToString().TrimEnd();
                }
            }

            // Final message composition 
            string SMSmessage = $"{strSMSTitle}\n{messageBalance}";
            strShortMessage = messageBalance;


            // Daily status report
            string emailMessage = "No change";
            bool generateReport = gnaSpreadsheetAPI.generateDailyReport(strPreviousDateTime, strDailyStatusReportTime);

            if (generateReport)
            {
                strShortMessage = "Daily status report";
                // === Build full status message for email ============================

                messageBalance = string.Empty;

                var sbEmail = new System.Text.StringBuilder();
                foreach (var s in alarmstate)
                {
                    if (s.currentAlarmState == "Alarm")
                    {
                        int total = s.Yes + s.No;
                        sbEmail.AppendLine($"{s.ATSname}({s.Settop}):{s.currentAlarmState}({s.Yes}/{total})");
                    }
                    else
                    {
                        sbEmail.AppendLine($"{s.ATSname}({s.Settop}):{s.currentAlarmState}");
                    }
                }
                messageBalance = sbEmail.ToString().TrimEnd();
                strShortMessage = messageBalance;

                // Final message composition
                emailMessage = $"{strProjectTitle}\n{messageBalance}";

                // Update the sms message as well
                string smsHeader = strSMSTitle + " status";
                SMSmessage = $"{smsHeader}\n{messageBalance}";
                emailMessage += $"\r\nAlarm triggers:\r\n1.Less than {iNoOfSuccessfulReadings} targets observed in past {dblAlarmWindowHrs}hrs.\r\n2.T4D server fails to process data in past {dblAlarmWindowHrs} hrs.";

                // Add copyright notice
                emailMessage = gnaT.addCopyright("AlarmStateWatchdog", emailMessage);

            }

            if (!string.IsNullOrEmpty(strShortMessage))
            {
                Console.WriteLine(strTab1 + strShortMessage);
                if (generateReport)
                {
                    Console.WriteLine(strTab1 + "Daily report generated");
                }
                else
                {
                    Console.WriteLine(strTab1 + "No daily report");
                    
                }
            }
           
            string strSendSMSandEmail = "No";
            if (generateReport ||
                (SMSmessage?.IndexOf("No change", StringComparison.OrdinalIgnoreCase) < 0))
            {
                strSendSMSandEmail = "Yes";
            }

            string emailHeader = strProjectTitle;
            if (strSendSMSandEmail == "Yes")
            {
                string strToday = DateTime.Today.ToString("yyyy-MM-dd");
                if (generateReport) {
                    emailHeader += ": Daily status report (" + DateTime.Today.ToString("yyyy-MM-dd") + " " + strDailyStatusReportTime + ")";
                }
                else
                {

                    emailHeader += ": Status change (" + DateTime.Today.ToString("yyyy-MM-dd") + " " + DateTime.Now.ToString("HH:mm") + ")";
                    emailMessage = gnaT.addCopyright("AlarmStateWatchdog", SMSmessage);
                }

                    string resultMsg = gnaT.SendEmailToRecipients(
                        emailCreds,
                        emailHeader,
                        emailMessage);
                Console.WriteLine(strTab1 + resultMsg);
                gnaT.updateSystemLogFile(strStatusFolder, strProjectTitle + ": "+resultMsg);

                bool smsSuccess = gnaT.sendSMSArray(SMSmessage, smsMobile);
                if (smsSuccess = true)
                {
                    Console.WriteLine(strTab1 + "SMS array sent to " + strMobileNumbers);
                    gnaT.updateSystemLogFile(strStatusFolder, strSMSTitle + ": Status SMS to "+strMobileNumbers);
                }
                else
                {
                    Console.WriteLine(strTab1 + "SMS array failed");
                    gnaT.updateSystemLogFile(strStatusFolder, strSMSTitle + ": SMS failed " + strMobileNumbers);
                }
            }
            else
            {
                Console.WriteLine(strTab1 + "No sms or emails sent");
            }








ThatsAllFolks:
            gnaT.freezeScreen(strFreezeScreen);
            Console.WriteLine("\nTask complete");
            Environment.Exit(0);
        }
            catch (Exception ex)
            {
                File.WriteAllText("fatal_crash.log", ex.ToString());
            }
}
    }
}