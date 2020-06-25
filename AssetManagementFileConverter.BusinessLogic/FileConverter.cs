using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;

namespace MPSAssestManagementFileConverter.BusinessLogic
{
    public class FileConverter
    {
        ////INPUT 
        ////A: schoolName 
        ////B: lastName    
        ////C: firstName 
        ////D: studentNumber   
        ////E: grade 
        ////F: birthdate   
        ////G: email 
        ////H: homeroom    
        ////I: NGACohortEndYear


        //OUTPUT CSV
        //First Name
        //Last Name
        //email
        //username
        //Location
        //phoneNumber
        //Job Title
        //EmployeeNumber
        //Company

        //Assumptions: 
        //First sheet is sheet with applicable data
        //Column Order is always the same 
        //No blank lines to eof

        public void ConvertFile(string inputFileNameAndPath, string outputFileNameAndPath, string companyName, string outputHeader)
        {
            var wb = new XLWorkbook(inputFileNameAndPath);
            var ws = wb.Worksheet(1);

            //If we need category names in future - for more dynamic input file
            //var categories = GetWorksheetHeaders(ws);

            //HeaderRow
            var colSchoolName = 1;
            var colLastName = 2;
            var colFirstName = 3;
            var colStudentNumber = 4;
            var colEmail = 7;


            // Look for the first row used
            var firstRowUsed = ws.FirstRowUsed();

            // Narrow down the row so that it only includes the used part
            var categoryRow = firstRowUsed.RowUsed();

            // Move to the next row 
            var currentRow = categoryRow.RowBelow();

            //var outputFileRecords = new List<OutputFileRecord>();

            using (var w = new StreamWriter(outputFileNameAndPath))
            {
                //var headerLineText = "First Name,Last Name,email,Username,Location,Phone Number,Job Title,Employee Number,Company";
                    w.WriteLine(outputHeader);
                    w.Flush();
                
                // Get all rows
                while (!currentRow.Cell(colSchoolName).IsEmpty())
                {
                    var outputFileRecord = ConvertExcelRowToOutputFileRecord(companyName, colSchoolName, colLastName, colFirstName, colStudentNumber, colEmail, currentRow);
                    //outputFileRecords.Add(outputFileRecord);
                    
                    var line = $"{outputFileRecord.FirstName},{outputFileRecord.LastName}," +
                        $"{outputFileRecord.Email},{outputFileRecord.Username}," +
                        $"{outputFileRecord.Location},{outputFileRecord.PhoneNumber}," +
                        $"{outputFileRecord.JobTitle},{outputFileRecord.EmployeeNumber}," +
                        $"{ outputFileRecord.Company}";

                    w.WriteLine(line);
                    w.Flush();
                    currentRow = currentRow.RowBelow();
                }
            }
           

            
        }

        private string EscapeCSVText(string data)
        {
            if (data.Contains("\""))
            {
                data = data.Replace("\"", "\"\"");
            }

            if (data.Contains(","))
            {
                data = String.Format("\"{0}\"", data);
            }

            if (data.Contains(System.Environment.NewLine))
            {
                data = String.Format("\"{0}\"", data);
            }

            return data;
        }

        private OutputFileRecord ConvertExcelRowToOutputFileRecord(string companyName, int colSchoolName, int colLastName, int colFirstName, int colStudentNumber, int colEmail, IXLRangeRow currentRow)
        {
            var firstName = currentRow.Cell(colFirstName).GetString();
            var lastName = currentRow.Cell(colLastName).GetString();
            var schoolName = currentRow.Cell(colSchoolName).GetString();
            var studentNumber = currentRow.Cell(colStudentNumber).GetString();
            var email = currentRow.Cell(colEmail).GetString();

            MailAddress addr = new MailAddress(email);
            string username = addr.User;

            var outputFileRecord = new OutputFileRecord
            {
                Company = EscapeCSVText(companyName),
                Email = email,
                EmployeeNumber = studentNumber,
                FirstName = EscapeCSVText(firstName),
                JobTitle = string.Empty,
                LastName = EscapeCSVText(lastName),
                Location = EscapeCSVText(schoolName),
                PhoneNumber = string.Empty,
                Username = username
            };
            return outputFileRecord;
        }

        private List<string> GetWorksheetHeaders(IXLWorksheet ws)
        {
            var categories = new List<string>();
            // Look for the first row used
            var firstRowUsed = ws.FirstRowUsed();

            // Narrow down the row so that it only includes the used part
            var categoryRow = firstRowUsed.RowUsed();

            //Get all categoreis
            foreach (var cell in categoryRow.Cells())
            {
                categories.Add(cell.GetString());
            }
            return categories;

        }
    }
}
