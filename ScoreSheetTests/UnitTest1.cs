using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScoreSheets;

namespace ScoreSheetTests
{
    [TestClass]
    public class UnitTest1
    {
        private static Random random = new Random();

        /// <summary>
        /// basic test, no errors in spreadsheets
        /// </summary>
        [TestMethod]
        public void NoErrors()
        {
            Controller controller = new Controller(new View());
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            List<string> people = new List<string>();
            List<string> numbers = new List<string>();
            List<string> USACpeople = new List<string>();
            List<string> USACNumbers = new List<string>();
            List<string> collMembers = new List<string>();
            List<string> allSeason = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);
            }

            for(int i = 0; i < 100; i++)
            {
                string number = random.Next().ToString();

                numbers.Add(number);
                USACNumbers.Add(number);
            }

            List<string> empty = controller.CheckRegionalsTest(people, numbers, allSeason, USACpeople, USACNumbers, collMembers);

            Assert.IsTrue(empty.Count == 0);
        }

        /// <summary>
        /// places a few people into the regionals check sheet and sees if the program detects them
        /// </summary>
        [TestMethod]
        public void BasicErrorsAll()
        {
            Controller controller = new Controller(new View());
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            List<string> people = new List<string>();
            List<string> numbers = new List<string>();
            List<string> USACpeople = new List<string>();
            List<string> USACNumbers = new List<string>();
            List<string> collMembers = new List<string>();
            List<string> allSeason = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);
            }

            List<string> checkResult = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);

                int num = i + 102;
                int place = i + 100;
                checkResult.Add("[" + num + "] The member [" + people[place] + " " + numbers[place] + "] could" +
                                        " not be found in the USAC members sheet.");
                checkResult.Add("[" + num + "] The member [" + people[place] + " " + numbers[place] + "] could" +
                                        " not be found in the full seasons events sheet.");
                checkResult.Add("[" + num + "] The member [" + people[place] + " " + numbers[place] + "] could" +
                                        " not be found in the collegiate members sheet.");
                checkResult.Add("--");
            }

            List<string> result = controller.CheckRegionalsTest(people, numbers, allSeason, USACpeople, USACNumbers, collMembers);

            foreach (string s in result)
            {
                Assert.IsTrue(checkResult.Contains(s));
            }
        }

        /// <summary>
        /// checks if it detects errors for USACMembers who registered under the wrong name only
        /// </summary>
        [TestMethod]
        public void BasicErrorsWrongName()
        {
            Controller controller = new Controller(new View());
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            List<string> people = new List<string>();
            List<string> numbers = new List<string>();
            List<string> USACpeople = new List<string>();
            List<string> USACNumbers = new List<string>();
            List<string> collMembers = new List<string>();
            List<string> allSeason = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);
            }

            List<string> checkResult = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);

                int num = i + 102;
                int place = i + 100;
                checkResult.Add("[" + num + "] The member [" + numbers[place] + "] registered under the wrong name.");
                checkResult.Add("--");
            }

            List<string> result = controller.CheckRegionalsTest(people, numbers, allSeason, USACpeople, USACNumbers, collMembers);

            foreach (string s in result)
            {
                Assert.IsTrue(checkResult.Contains(s));
            }
        }

        /// <summary>
        /// checks if it detects errors for people who could not be found on the USAC member sheet
        /// </summary>
        [TestMethod]
        public void BasicErrorsUSACMember()
        {
            Controller controller = new Controller(new View());
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            List<string> people = new List<string>();
            List<string> numbers = new List<string>();
            List<string> USACpeople = new List<string>();
            List<string> USACNumbers = new List<string>();
            List<string> collMembers = new List<string>();
            List<string> allSeason = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);
            }

            List<string> checkResult = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);

                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);

                int num = i + 102;
                int place = i + 100;
                checkResult.Add("[" + num + "] The member [" + people[place] + " " + numbers[place] + "] could" +
                                " not be found in the USAC members sheet.");
                checkResult.Add("--");
            }

            List<string> result = controller.CheckRegionalsTest(people, numbers, allSeason, USACpeople, USACNumbers, collMembers);

            foreach (string s in result)
            {
                Assert.IsTrue(checkResult.Contains(s));
            }
        }

        /// <summary>
        /// checks if it detects errors for people who could not be found on the collegiate members sheet
        /// </summary>
        [TestMethod]
        public void BasicErrorsCollegiateMembers()
        {
            Controller controller = new Controller(new View());
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            List<string> people = new List<string>();
            List<string> numbers = new List<string>();
            List<string> USACpeople = new List<string>();
            List<string> USACNumbers = new List<string>();
            List<string> collMembers = new List<string>();
            List<string> allSeason = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);
            }

            List<string> checkResult = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);

                int num = i + 102;
                int place = i + 100;
                checkResult.Add("[" + num + "] The member [" + people[place] + " " + numbers[place] + "] could" +
                                        " not be found in the collegiate members sheet.");
                checkResult.Add("--");
            }

            List<string> result = controller.CheckRegionalsTest(people, numbers, allSeason, USACpeople, USACNumbers, collMembers);

            foreach (string s in result)
            {
                Assert.IsTrue(checkResult.Contains(s));
            }
        }

        /// <summary>
        /// checks if it detects errors for people who could not be found on the all season sheet
        /// </summary>
        [TestMethod]
        public void BasicErrorsAllSeason()
        {
            Controller controller = new Controller(new View());
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            List<string> people = new List<string>();
            List<string> numbers = new List<string>();
            List<string> USACpeople = new List<string>();
            List<string> USACNumbers = new List<string>();
            List<string> collMembers = new List<string>();
            List<string> allSeason = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);
                allSeason.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);
            }

            List<string> checkResult = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string name = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                people.Add(name);
                USACpeople.Add(name);
                collMembers.Add(name);

                string number = random.Next().ToString();
                numbers.Add(number);
                USACNumbers.Add(number);

                int num = i + 102;
                int place = i + 100;
                checkResult.Add("[" + num + "] The member [" + people[place] + " " + numbers[place] + "] could" +
                                        " not be found in the full seasons events sheet.");
                checkResult.Add("--");
            }

            List<string> result = controller.CheckRegionalsTest(people, numbers, allSeason, USACpeople, USACNumbers, collMembers);

            foreach (string s in result)
            {
                Assert.IsTrue(checkResult.Contains(s));
            }
        }
    }
}
