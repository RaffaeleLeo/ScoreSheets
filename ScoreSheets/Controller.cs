using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Xml.Linq;
using System.Drawing;

namespace ScoreSheets
{
    /// <summary>
    /// Does all the backend work
    /// takes in spreadsheets, scans the info we want, and compares them
    /// </summary>
    public class Controller
    {
        View view;

        //for regionals
        string regionalsRegistranceFile;
        string fullSeasonEventsFile;
        string USAClimbingMebershipFile;
        string collegiateMembershipFile;

        List<string> registranceNames;
        List<string> registranceNumbers;

        List<string> fullSeasonEventsNames;
        List<string> USACMemberNames;
        List<string> USACMemberNumbers;
        List<string> collegiateMembers;

        //for divisionals or nationals
        private string fileToCheck;
        private string referenceFile;

        List<string> members;
        List<string> names;

        List<string> nameCheck;
        List<string> memberCheck;
        List<string> qualifiedCheck;

        /// <summary>
        /// Constructs the controller taking a Model gui as a paramater
        /// </summary>
        /// <param name="view"></param>
        public Controller(View view)
        {
            this.view = view;
            //divisionals / nationals stuff
            fileToCheck = "";
            referenceFile = "";

            members = new List<string>();
            names = new List<string>();

            nameCheck = new List<string>();
            memberCheck = new List<string>();
            qualifiedCheck = new List<string>();

            //for divisionals or nationals registrance
            view.SelectFilePressed += SelectFile1;

            //for regionals registrance
            view.RegionalsCheckPressed += SelectFile2;

            //more regionals stuff
            regionalsRegistranceFile = "";
            fullSeasonEventsFile = "";
            USAClimbingMebershipFile = "";
            collegiateMembershipFile = "";

            registranceNames = new List<string>();
            registranceNumbers = new List<string>();

            fullSeasonEventsNames = new List<string>();
            USACMemberNames = new List<string>();
            USACMemberNumbers = new List<string>();
            collegiateMembers = new List<string>();
        }

        //.............................................................
        //CODE FOR REGIONALS CHECK
        //.............................................................

        /// <summary>
        /// For regionals check, scan through 4 different files:
        /// 1. Regionals Registrance
        /// 2. All Seasons Events
        /// 3. USAC membership check
        /// 4. Collegiate membership check
        /// </summary>
        public void SelectFile2()
        {
            //have the user select the excel sheet for regionals registrance
            var fileContent = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select Regional Registrance Sheet";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    regionalsRegistranceFile = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }

            //have the user select the excel sheet for the full seasons events
            var fileContent1 = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select Full Seasons Events Sheet";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    fullSeasonEventsFile = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }

            //have the user select the excel sheet for USA climbing membership check
            var fileContent2 = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select USA Climbing Membership Check Sheet";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    USAClimbingMebershipFile = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }

            //have the user select the excel sheet for collegiate membership status check
            var fileContent3 = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select The Collegiate Membership Check Sheet";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    collegiateMembershipFile = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }

            //this is kind of messy but the next 4 methods are for scanning
            //through each sheet and extracting the information we want
            ScanThroughRegionalsRegistrance();
            ScanThroughFullSeasonsEvents();
            ScanThroughUSACMembership();
            ScanThroughCollegiateMembership();
            //and finally compare the sheets
            CheckRegionalsSheets();
        }

        /// <summary>
        /// Scans through the regionals registrance sheet and extracts the 
        /// names and numbers of the competitors listed
        /// </summary>
        public void ScanThroughRegionalsRegistrance()
        {
            //get the excel file
            FileInfo newFile = new FileInfo(regionalsRegistranceFile);
            
            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                //find what columns the info we need is located
                int namesColumn = 0;
                int memberColumn = 0;
                for (int column = start.Row; column <= end.Column; column++)
                {
                    if (worksheet.Cells[1, column].GetValue<string>() == null)
                    {

                    }
                    else if (worksheet.Cells[1, column].GetValue<string>().Equals("Participant: FullName (First Last)"))
                    {
                        namesColumn = column;
                    }
                    else if (worksheet.Cells[1, column].GetValue<string>().Equals("Member No."))
                    {
                        memberColumn = column;
                    }
                }

                //now read through the rest of the information
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        if (col == namesColumn)
                        {
                            registranceNames.Add(worksheet.Cells[row, col].GetValue<string>());
                        }

                        if (col == memberColumn)
                        {
                            registranceNumbers.Add(worksheet.Cells[row, col].GetValue<string>());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// scans through the full seasons events sheet
        /// </summary>
        public void ScanThroughFullSeasonsEvents()
        {
            //get the excel file
            FileInfo newFile = new FileInfo(fullSeasonEventsFile);

            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                //find what columns the info we need is located
                int namesColumn = 0;

                for (int column = start.Row; column <= end.Column; column++)
                {
                    if (worksheet.Cells[1, column].GetValue<string>() == null)
                    {

                    }
                    else if (worksheet.Cells[1, column].GetValue<string>().Equals("Participant: FullName (First Last)"))
                    {
                        namesColumn = column;
                    }
                }

                //now read through the rest of the information
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        if (col == namesColumn)
                        {
                            fullSeasonEventsNames.Add(worksheet.Cells[row, col].GetValue<string>());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// scans through the USAC membership check sheet
        /// </summary>
        public void ScanThroughUSACMembership()
        {
            //get the excel file
            FileInfo newFile = new FileInfo(USAClimbingMebershipFile);

            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                //find what columns the info we need is located
                int namesColumn = 0;
                int memberColumn = 0;
                for (int column = start.Row; column <= end.Column; column++)
                {
                    if(worksheet.Cells[1, column].GetValue<string>() == null)
                    {

                    }
                    else if (worksheet.Cells[1, column].GetValue<string>().Equals("Participant: FullName (First Last)"))
                    {
                        namesColumn = column;
                    }
                    else if (worksheet.Cells[1, column].GetValue<string>().Equals("Member No."))
                    {
                        memberColumn = column;
                    }
                }

                //now read through the rest of the information
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        if (col == namesColumn)
                        {
                            USACMemberNames.Add(worksheet.Cells[row, col].GetValue<string>());
                        }

                        if (col == memberColumn)
                        {
                            USACMemberNumbers.Add(worksheet.Cells[row, col].GetValue<string>());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// scans through the collegiate membership check sheet
        /// </summary>
        public void ScanThroughCollegiateMembership()
        {
            //get the excel file
            FileInfo newFile = new FileInfo(collegiateMembershipFile);

            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                int namesColumn = 0;

                for (int column = start.Row; column <= end.Column; column++)
                {
                    if (worksheet.Cells[1, column].GetValue<string>() == null)
                    {

                    }
                    else if (worksheet.Cells[1, column].GetValue<string>().Equals("Participant: FullName (First Last)"))
                    {
                        namesColumn = column;
                    }
                }

                //now read through the rest of the information
                for (int row = start.Row; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        if (col == namesColumn)
                        {
                            collegiateMembers.Add(worksheet.Cells[row, col].GetValue<string>());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// THIS IS A HELPER METHOD USED FOR TESTING ONLY, instead of using the global variables it
        /// takes in its own lists as paramaters.
        /// </summary>
        /// <param name="people"></param>
        /// <param name="memberNumbers"></param>
        /// <param name="allSeason"></param>
        /// <param name="USACPeople"></param>
        /// <param name="USAClimbingMember"></param>
        /// <param name="collegiateMember"></param>
        /// <returns></returns>
        public List<string> CheckRegionalsTest(List<string> people, List<string> memberNumbers, List<string> allSeason, 
                                               List<string> USACPeople, List<string> USAClimbingMember, List<string> collegiateMember)
        {
            List<string> writeToFile = new List<string>();
            List<int> numbers = new List<int>();
            bool errorFound = false;
            //scan through all names of people who signed up for regionals
            for (int i = 0; i < people.Count(); i++)
            {
                errorFound = false;
                //check if name is in USAC membership check sheet
                if (!USACPeople.Contains(people[i]))
                {
                    //if they are not in the usac member names but the 
                    //number is correct they signed up under the wrong name
                    if (USAClimbingMember.Contains(memberNumbers[i]))
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("[" + num + "] The member [" + memberNumbers[i] + "] registered under the wrong name.");
                    }
                    else
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("[" + num + "] The member [" + people[i] + " " + memberNumbers[i] + "] could" +
                                        " not be found in the USAC members sheet.");
                    }
                    errorFound = true;
                }
                //see if there name comes up in the full season events names sheet
                if (!allSeason.Contains(people[i]))
                {
                    int num = i + 2;
                    numbers.Add(num);
                    writeToFile.Add("[" + num + "] The member [" + people[i] + " " + memberNumbers[i] + "] could" +
                                        " not be found in the full seasons events sheet.");
                    errorFound = true;
                }
                //see if it comes up in the collegiate member check
                if (!collegiateMember.Contains(people[i]))
                {
                    int num = i + 2;
                    numbers.Add(num);
                    writeToFile.Add("[" + num + "] The member [" + people[i] + " " + memberNumbers[i] + "] could" +
                                        " not be found in the collegiate members sheet.");
                    errorFound = true;
                }

                if (errorFound)
                {
                    writeToFile.Add("--");
                }
            }
            return writeToFile;
        }

        /// <summary>
        /// check the regionals signees against the other spreadsheets
        /// </summary>
        public void CheckRegionalsSheets()
        {
            List<string> writeToFile = new List<string>();
            List<int> numbers = new List<int>();
            bool errorFound = false;
            //scan through all names of people who signed up for regionals
            for(int i = 0; i < registranceNames.Count(); i++)
            {
                errorFound = false;
                //check if name is in USAC membership check sheet
                if (!USACMemberNames.Contains(registranceNames[i]))
                {
                    //if they are not in the usac member names but the 
                    //number is correct they signed up under the wrong name
                    if(USACMemberNumbers.Contains(registranceNumbers[i]))
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("[" + num + "] The member [" + registranceNumbers[i] + "] registered under the wrong name.");
                    }
                    else
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("[" + num + "] The member [" + registranceNames[i] + " " + registranceNumbers[i] + "] could" +
                                        " not be found in the USAC members sheet.");
                    }
                    errorFound = true;
                }
                //see if there name comes up in the full season events names sheet
                if (!fullSeasonEventsNames.Contains(registranceNames[i]))
                {
                    int num = i + 2;
                    numbers.Add(num);
                    writeToFile.Add("[" + num + "] The member [" + registranceNames[i] + " " + registranceNumbers[i] + "] could" +
                                        " not be found in the full seasons events sheet.");
                    errorFound = true;
                }
                //see if it comes up in the collegiate member check
                if (!collegiateMembers.Contains(registranceNames[i]))
                {
                    int num = i + 2;
                    numbers.Add(num);
                    writeToFile.Add("[" + num + "] The member [" + registranceNames[i] + " " + registranceNumbers[i] + "] could" +
                                        " not be found in the collegiate members sheet.");
                    errorFound = true;
                }

                if(errorFound)
                {
                    writeToFile.Add("--");
                }
            }

            string file = Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory);
            file = file + "\\errors.txt";

            FileInfo newfile = new FileInfo(regionalsRegistranceFile);

            //now read through the signees sheet and highlight any problems found in red
            using (ExcelPackage package = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                //reset the spreadsheet
                for (int row = start.Row + 1; row <= end.Row; row++)
                {
                    worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                }
                //loop through all the problems found
                foreach (int row in numbers)
                {
                    worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Red);
                }

                package.Save();
            }


            System.IO.File.WriteAllLines(file, writeToFile);
            System.Windows.Forms.MessageBox.Show("All Done!");
        }





        ///.....................................................................................................
        ///CODE FOR DIVISIONALS/NATIONALS CHECK:
        ///.....................................................................................................
        ///




        /// <summary>
        /// has the user select two files and saves the paths in field variables
        /// </summary>
        private void SelectFile1()
        {
            var fileContent = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select Registration Sheet";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    fileToCheck = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }

            var fileContent_2 = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select Comp Results";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    referenceFile = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent_2 = reader.ReadToEnd();
                    }
                }
            }

            //ScanThroughCSVToCheck();
            ScanThroughSighneesXML();
            ScanThroughRegistrantsXML();
            CompareSheets();
        }

        ///................................................
        ///FOR READING THROUGH EXCEL FILES, CODE FOR REGIONALS/NATIONALS CHECK
        ///................................................

        /// <summary>
        /// scans through an xml file using epplus
        /// </summary>
        private void ScanThroughSighneesXML()
        {
            //get the excel file
            FileInfo newFile = new FileInfo(fileToCheck);

            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                //find what columns the info we need is located
                int namesColumn = 0;
                int memberColumn = 0;
                for (int column = start.Row; column <= end.Column; column++)
                {
                    if (worksheet.Cells[1, column].GetValue<string>().Equals("Participant: FullName (First Last)"))
                    {
                        namesColumn = column;
                    }

                    if (worksheet.Cells[1, column].GetValue<string>().Equals("Member No."))
                    {
                        memberColumn = column;
                    }
                }

                //now read through the rest of the information
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        if (col == namesColumn)
                        {
                            names.Add(worksheet.Cells[row, col].GetValue<string>());
                        }

                        if (col == memberColumn)
                        {
                            members.Add(worksheet.Cells[row, col].GetValue<string>());
                        }
                    }
                }
            }
            
        }

        /// <summary>
        /// scans through the results list and saves the names, member info, and weather they qualified
        /// </summary>
        private void ScanThroughRegistrantsXML()
        {
            //get the excel file
            FileInfo newFile = new FileInfo(referenceFile);

            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                //find what columns the info we need is located
                int namesColumn = 0;
                int memberColumn = 0;
                int qualifiedColumn = 0;
                for (int column = start.Row; column <= end.Column; column++)
                {
                    if (worksheet.Cells[1, column].GetValue<string>().Equals("Participant: FullName (First Last)"))
                    {
                        namesColumn = column;
                    }

                    if (worksheet.Cells[1, column].GetValue<string>().Equals("Member No."))
                    {
                        memberColumn = column;
                    }

                    if(worksheet.Cells[1, column].GetValue<string>().Equals("Qualified?"))
                    {
                        qualifiedColumn = column;
                    }
                }

                //now read through the rest of the information
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        if (col == namesColumn)
                        {
                            nameCheck.Add(worksheet.Cells[row, col].GetValue<string>());
                        }

                        if (col == memberColumn)
                        {
                            memberCheck.Add(worksheet.Cells[row, col].GetValue<string>());
                        }

                        if (col == qualifiedColumn)
                        {
                            if(worksheet.Cells[row, col].GetValue<string>() == null)
                            {
                                qualifiedCheck.Add("");
                            }
                            else
                            {
                                qualifiedCheck.Add("Yes");
                            }
                        }
                    }
                }
            }
            
        }

        //now that we have all the information lets compare it
        private void CompareSheets()
        {
            List<string> writeToFile = new List<string>();
            List<string> qualifiedList = new List<string>();
            List<int> numbers = new List<int>();

            qualifiedList.Add("QUALIFY CHECK: everything below this is checking to make sure the competitors have qualified");
            qualifiedList.Add("MAKE SURE ALL THE ERRORS ABOVE ARE FIXED OR THIS WONT SHOW ALL THE CORRECT RESULTS");
            qualifiedList.Add("--");

            //loop through all uncertain names
            for (int i = 1; i < names.Count; i++)
            {
                //see if the name is found in the database
                if (!nameCheck.Contains(names[i]))
                {
                    //if not see if that names number corresponds with a member in the database
                    if (memberCheck.Contains(members[i]))
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("[" + num + "] The member [" + members[i] + "] registered under the wrong name.");
                        writeToFile.Add("--");
                    }
                    else
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("[" + num + "] The member [" + names[i] + "] with membership number [" + members[i] + "] could not be found in the register check.");
                        writeToFile.Add("--");
                    }
                }
                else
                {
                    //make sure they are qualified
                    for (int j = 0; j < memberCheck.Count; j++)
                    {
                        //find the boi
                        if (members[i].Equals(memberCheck[j]))
                        {
                            //make sure hes qualified
                            if (!qualifiedCheck[j].ToString().Equals("Yes"))
                            {
                                int num = i + 2;
                                numbers.Add(num);
                                qualifiedList.Add("[" + num + "] The competitor [" + names[i] + "] [" + members[i] + "] failed to qualify :/");
                                qualifiedList.Add("--");
                            }
                        }
                    }
                }
            }

            //now go through membership numbers
            for (int i = 0; i < members.Count; i++)
            {
                //see if the membership number is found in the database
                if (!memberCheck.Contains(members[i]))
                {
                    //if not see if the number corresponds with a name in the database
                    if (nameCheck.Contains(names[i]))
                    {
                        int num = i + 2;
                        numbers.Add(num);
                        writeToFile.Add("The member [" + names[i] + "] registered under the wrong USA climbing account" +
                                        " (located at row [" + num + "] within the divisionals registrans chart)");
                        writeToFile.Add("--");
                    }
                }
            }

            string file = Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory);
            file = file + "\\errors.txt";

            foreach(string s in qualifiedList)
            {
                writeToFile.Add(s);
            }
            FileInfo newFile = new FileInfo(fileToCheck);

            //read through it, save all the values we want into arrays
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                

                foreach (int row in numbers)
                {
                    worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Red);
                }

                package.Save();
            }
            

            System.IO.File.WriteAllLines(file, writeToFile);
            System.Windows.Forms.MessageBox.Show("All Done!");
        }
    }
}
