using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Vqs.Excel;
using Vqs.ExcelTest.Models.Horizontal;
using Vqs.ExcelTest.Models.Vertical;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Vqs.Excel.Test
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            string inputFilePath = Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.FullName,
                        "DemoFile",
                        "ExcelDemo.xlsx");
            FileInfo intpuFileInfo = new FileInfo(inputFilePath);

            string outputFilePath = Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.FullName,
                        "DemoFile",
                        "output.txt");
            using (StreamWriter outputFile = File.CreateText(outputFilePath))
            {
                // Grab demo file
                using (var excel = new ExcelPackage(intpuFileInfo)) // If missing replace FileInfo constructor parameter with your own file
                {
                    // Horizontal
                    outputFile.WriteLine("Horizontal mapping");

                    // Get sheet with horizontal mapping
                    var sheet = excel.Workbook.Worksheets.First();

                    // Get list of teams based on automatically mapping of the header row
                    //outputFile.WriteLine("List of teams based on automatically mapping of the header row");
                    //var teams = sheet.GetRecords<Team>();
                    //foreach (var team in teams)
                    //{
                    //    outputFile.WriteLine($"{team.Name} - {team.FoundationYear} - {team.Titles}");
                    //}

                    //// Get specific record from sheet based on automatically mapping the header row
                    //outputFile.WriteLine("Team based on automatically mapping of the header row");
                    //var teamRec = sheet.GetRecord<Team>(2);
                    //outputFile.WriteLine($"{teamRec.Name} - {teamRec.FoundationYear} - {teamRec.Titles}");

                    //// Remove HeaderRow
                    //sheet.DeleteRow(1);

                    // Get list of teams based on mapping using attributes
                    outputFile.WriteLine("List of teams based on mapping using attributes");
                    var teamsAttr = sheet.GetRecords<TeamAttributes>();
                    int failcount = GenericValidator.TryValidate(teamsAttr);
                    foreach (var team in teamsAttr)
                    {
                        bool isValid = team.IsValid;
                        string msg = string.Format($"{team.Name},{team.Designation},{team.DOB},{team.Points}");

                        if (!isValid)
                        {
                            msg += ",FAILED," + team.ErrorMessage;
                            failcount++;
                        }
                        else
                        {
                            msg += ",PASSED";
                        }

                        outputFile.WriteLine(msg);
                    }

                    if(failcount > 0)
                    {
                        var resultSheet = excel.Workbook.Worksheets.AddOrReplace("Upload Result");
                        resultSheet.Cells["A1"].LoadFromCollection<TeamAttributes>(teamsAttr, true);
                        excel.Save();
                    }

                    // Get specific record from sheet based on mapping using attributes
                    //outputFile.WriteLine("Team based on mapping using attributes");
                    //var teamAttr = sheet.GetRecord<TeamAttributes>(1);
                    //outputFile.WriteLine($"{teamAttr.Name} - {teamAttr.FoundationYear} - {teamAttr.Titles}");

                    //// Get list of teams based on user created map
                    //outputFile.WriteLine("List of teams based on user created map");
                    //var teamsMap = sheet.GetRecords(TeamMap.Create());
                    //foreach (var team in teamsMap)
                    //{
                    //    outputFile.WriteLine($"{team.Name} - {team.FoundationYear} - {team.Titles}");
                    //}

                    //// Get specific record from sheet based on user created map
                    //outputFile.WriteLine("Team based on user created map");
                    //var teamMap = sheet.GetRecord<Team>(1, TeamMap.Create());
                    //outputFile.WriteLine($"{teamMap.Name} - {teamMap.FoundationYear} - {teamMap.Titles}");

                    //// Vertical
                    //outputFile.WriteLine("Vertical mapping");

                    //// Get sheet with vertical mapping
                    //var vsheet = excel.Workbook.Worksheets.Skip(1).Take(1).First();

                    //// Get list of teams based on automatically mapping of the header row
                    //outputFile.WriteLine("List of teams based on automatically mapping of the header row");
                    //var vteams = vsheet.GetRecords<VTeam>();
                    //foreach (var vteam in vteams)
                    //{
                    //    outputFile.WriteLine($"{vteam.Name} - {vteam.FoundationYear} - {vteam.Titles}");
                    //}

                    //// Get specific record from sheet based on automatically mapping the header row
                    //outputFile.WriteLine("Team based on automatically mapping of the header row");
                    //var vteamRec = vsheet.GetRecord<VTeam>(2);
                    //outputFile.WriteLine($"{vteamRec.Name} - {vteamRec.FoundationYear} - {vteamRec.Titles}");

                    //// Remove HeaderRow
                    //vsheet.DeleteColumn(1);

                    //// Get list of teams based on mapping using attributes
                    //outputFile.WriteLine("List of teams based on mapping using attributes");
                    //var vteamsAttr = vsheet.GetRecords<VTeamAttributes>();
                    //foreach (var vteam in vteamsAttr)
                    //{
                    //    outputFile.WriteLine($"{vteam.Name} - {vteam.FoundationYear} - {vteam.Titles}");
                    //}

                    //// Get specific record from sheet based on mapping using attributes
                    //outputFile.WriteLine("Team based on mapping using attributes");
                    //var vteamAttr = vsheet.GetRecord<VTeamAttributes>(1);
                    //outputFile.WriteLine($"{vteamAttr.Name} - {vteamAttr.FoundationYear} - {vteamAttr.Titles}");

                    //// Get list of teams based on user created map
                    //outputFile.WriteLine("List of teams based on user created map");
                    //var vteamsMap = vsheet.GetRecords(VTeamMap.Create());
                    //foreach (var vteam in vteamsMap)
                    //{
                    //    outputFile.WriteLine($"{vteam.Name} - {vteam.FoundationYear} - {vteam.Titles}");
                    //}

                    //// Get specific record from sheet based on user created map
                    //outputFile.WriteLine("Team based on user created map");
                    //var vteamMap = vsheet.GetRecord<VTeam>(1, VTeamMap.Create());
                    //outputFile.WriteLine($"{vteamMap.Name} - {vteamMap.FoundationYear} - {vteamMap.Titles}");
                }
            }
        }
    }
}