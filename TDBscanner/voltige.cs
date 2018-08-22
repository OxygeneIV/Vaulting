using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Framework.PageObjects;
using Framework.WaitHelpers;
using Framework.WebDriver;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Tests.Base;
using Viedoc.viedoc.pages.components.elements;

namespace Tests.Voltige
{
    [Locator(How.Sizzle, "body:has(#email)")]
    public class LoginPage : PageObject
    {
        [Locator("#email")] public TextField email;

        [Locator("#password")] public PasswordField password;

        [Locator("[type=submit]")] public Button SubmitButton;
    }

    /// <summary>
    /// Viedoc Details Page
    /// </summary>
    [Locator(How.Sizzle, "body:has(table:first)")]
    public class CompetitionPage : PageObject
    {
        [Locator(How.Sizzle, "table:first")] public Table ClassesTable;
    }

    [Locator(How.Sizzle, "body:has(h4:contains(Allmänt))")]
    public class CompetitorPage : PageObject
    {
        [Locator(".row p a[href*='/clubs/']")]
        public Link ClubLink;

        public string ClubLinkText => ClubLink.TrimmedText;
    }

    /// <summary>
    /// Viedoc Details Page
    /// </summary>
    [Locator(How.Sizzle, "body:has(table:first.tablesorter)")]
    public class ClassPage : PageObject
    {
        [Locator(How.Sizzle, "table:first.tablesorter")] public Table CompetitorTable;
    }


    [Locator(How.Sizzle, "body:has(.pills)")]
    public class EkipagePage : PageObject
    {
        [Locator(How.Sizzle, "table:first")] public Table EkipageTable;

        [Locator("[href*=note]")]
        public Link NoteLink;

        public NotePage OpenNotePage()
        {
            NoteLink.Click();
            var page = PageObjectFactory.Init<NotePage>(this.WebDriver);
            Wait.UntilOrThrow(()=>page.Displayed);
            return page;
        }
    }

    [Locator(How.Sizzle, "body:has(#note)")]
    public class NotePage : PageObject
    {
        private TextAreaField NoteField;

        public string GetNoteText()
        {
            var txt = NoteField.TextValue.Trim();
            var trimmed = txt.Replace(System.Environment.NewLine, ", ");
            return trimmed;
        }

    }

    public class Ekipage
    {
        public int EkipageId;
        public int LinförareId;
        public int HorseId;
        public int KlubbId;
    }

    public class CompClass
    {
        public int number;
        public string name;
        public int nofComp;
        public int TDBid;
    }

    [TestClass]
    public class VoltigeTests : ViedocTestbase
    {

        public List<CompClass> classes;


        [TestMethod]
        [TestCategory("VOLTIGE")]
        [Browser(false)]
        public void ReadClasses()
        {
            // Login
            string tdbUrl = "https://tdb.ridsport.se/login";

            //string compUrl = "https://tdb.ridsport.se/clubs/223/meetings/43640";
            // SM  compUrl =    "https://tdb.ridsport.se/meetings/47124";
            //string meetingUrl = "https://tdb.ridsport.se/meetings/45646";
            string meetingUrl = "https://tdb.ridsport.se/meetings/47124";

            // Open Browser
            var driver = CreateBrowserInstance(Driver.Browser.Chrome);

            // Goto TDB
            driver.Navigate().GoToUrl(tdbUrl);

            // Login
            LoginPage l = PageObjectFactory.Init<LoginPage>(driver);
            //l.email.SetText("annaomagnus@hotmail.com");
            //l.password.SetText("berlin96");
            l.email.SetText("oxygeneiv@hotmail.com");
            l.password.SetText("xfiles67");
            l.SubmitButton.Click();

            // Inistatiate the Competition page
            CompetitionPage c = PageObjectFactory.Init<CompetitionPage>(driver);

            // Goto Meeting
            c.WebDriver.Navigate().GoToUrl(meetingUrl);

            // Chec we have a table of classes
            Wait.UntilOrThrow(() => c.ClassesTable.Displayed);

            var rows = c.ClassesTable.Rows.ToList();
            var numberOfClasses = rows.Count;

            // All ints are DB Ids, not what is diaplyed in the table 1, 2, 3.1, 3.2, 4  etc

            Dictionary<int, string> _classes = new Dictionary<int, string>();
            Dictionary<int, string> _linf = new Dictionary<int, string>();
            Dictionary<int, string> _horse = new Dictionary<int, string>();
            Dictionary<int, string> _clubs = new Dictionary<int, string>();
            Dictionary<int, string> _comp = new Dictionary<int, string>();
            Dictionary<int, string> _ekipage = new Dictionary<int, string>();


            // Loop through the classes

            for (int i = 0; i < numberOfClasses; i++)
            {
                c = PageObjectFactory.Init<CompetitionPage>(driver);

                // Get the classes table
                Wait.UntilOrThrow(() => c.ClassesTable.Displayed);

                // Fetch the rows
                rows = c.ClassesTable.Rows.ToList();

                // Set current row
                var curClassrow = rows[i];

                // DATA
                var classnr = curClassrow.GetCellText("Nr"); // Class Nr

                var classnamn = curClassrow.GetCellText("Namn"); // Class Name
                classnamn = classnamn.Replace(System.Environment.NewLine, " ");

                var classanmalda = curClassrow.GetCellText("Anmälda");
                int anm = Int32.Parse(classanmalda);


                // Click on the Anmälda link to get easy url access
                curClassrow.GetCell("Anmälda").LinkClick();

                // Now we have the class details Page
                ClassPage classpage = PageObjectFactory.Init<ClassPage>(driver);
                Wait.UntilOrThrow(() => classpage.Displayed);

                // Check # anmälda
                var comprows = classpage.CompetitorTable.Rows.ToList();

                // No competitors, goto next class

                var classId = classpage.WebDriver.Url;
                var classurls = classId.Split('/');
                var classIdNum = Int32.Parse(classurls[classurls.Length - 2]);
                _classes[classIdNum] = classnr + "|" + classnamn;



                var checkCounter = 0;

                if (anm > 0)
                { 
                    // Loop through the competitors
                    for (var j = 0; j < comprows.Count; j++)
                    {
                        classpage = PageObjectFactory.Init<ClassPage>(driver);
                        Wait.UntilOrThrow(() => classpage.CompetitorTable.Displayed);

                        //// Fetch class Id Num in DB
                        //var classId = classpage.WebDriver.Url;
                        //var classurls = classId.Split('/');
                        //var classIdNum = Int32.Parse(classurls[classurls.Length - 2]);


                        comprows = classpage.CompetitorTable.Rows.ToList();
                        var curCompsrow = comprows[j];

                        // Avanmäld ?
                        var status = curCompsrow.GetCell("Status");
                        if (status.Text.ToLower().Contains("avanmäld"))
                            continue;

                        var linfCell = curCompsrow.GetCell("Linförare");
                        var horseCell = curCompsrow.GetCell("Häst");
                        var clubCell = curCompsrow.GetCell("Klubb");


                        checkCounter++;

                        var linfId = linfCell.LinkUrl;
                        var horseId = horseCell.LinkUrl;
                        var clubId = clubCell.LinkUrl;
                        var ekipageId = curCompsrow.GetCell(5).LinkUrl;

                        var linfIdNum = Int32.Parse(linfId.Split('/').Last());
                        var horseIdNum = Int32.Parse(horseId.Split('/').Last());
                        var clubIdNum = Int32.Parse(clubId.Split('/').Last());
                        var ekipageIdNum = Int32.Parse(ekipageId.Split('/').Last());

                        var linftext = linfCell.Text.Trim();
                        var horseCellText = horseCell.Text.Trim();
                        var clubCellText = clubCell.Text.Trim();

                        //_classes[classIdNum] = classnr + "|" + classnamn;
                        _linf[linfIdNum] = linftext;
                        _horse[horseIdNum] = horseCellText;
                        _clubs[clubIdNum] = clubCellText;


                        // Open tävlande
                        curCompsrow.GetCell(5).LinkClick();

                        EkipagePage ekipagePage = PageObjectFactory.Init<EkipagePage>(driver);
                        Wait.UntilOrThrow(() => ekipagePage.Displayed);

                        // Get the note
                        var notepage = ekipagePage.OpenNotePage();
                        var notetext = notepage.GetNoteText();
                        driver.Navigate().Back();

                        ekipagePage = PageObjectFactory.Init<EkipagePage>(driver);
                        Wait.UntilOrThrow(() => ekipagePage.Displayed);


                        var table = ekipagePage.EkipageTable;
                        var ekrows = table.Rows;
                        var voltigorercell = ekrows.Last().GetCell(2);
                        var links = voltigorercell.Links;
                        var n = links.Count();

                        bool individuell = n == 1;

                        var voltigorids = links.Select(voltid => voltid.GetAttribute("href")).ToList();
                        var voltigoridsNums = voltigorids.Select(v => Int32.Parse(v.Split('/').Last())).ToList();

                        var voltigoridsClubs = voltigorids.Select(v => Int32.Parse(v.Split('/')[4])).ToList();


                        var voltigornames = links.Select(voltid => voltid.Text).ToList();
                        var nns = String.Join(",", voltigornames);
                        var nums = String.Join(",", voltigoridsNums);

                        var nnn = "";
                        for (int k = 0; k < voltigorids.Count; k++)
                        {
                            _comp[voltigoridsNums[k]] = voltigornames[k];
                            nnn = nnn + "|" + voltigoridsNums[k] + "|" + voltigornames[k];
                        }

                        // Open first tävlande and register club
                        var firstVolter = links.First();

                        firstVolter.Click();
                        clubIdNum = voltigoridsClubs.First();


                        var compPage = PageObjectFactory.Init<CompetitorPage>(ekipagePage.WebDriver);
                        Wait.UntilOrThrow(() => compPage.ClubLink.Displayed);
                        var clubIdName = compPage.ClubLinkText.Trim();
                        if (!_clubs.ContainsKey(clubIdNum))
                        {
                            // Not yet defined
                            Log.Info($"Adding club ${clubIdName} (${clubIdNum})");
                            _clubs[clubIdNum] = clubIdName;
                        }

                        compPage.WebDriver.Navigate().Back();

                        Trace.WriteLine(classIdNum + "|" + _classes[classIdNum] + "|" + linfIdNum + "|" +
                                        _linf[linfIdNum] +
                                        "|" + horseIdNum + "|" + _horse[horseIdNum] + "|" + clubIdNum + "|" +
                                        _clubs[clubIdNum] + "|" + notetext + nnn);

                        driver.Navigate().Back();
                    }
            }

            Assert.AreEqual(anm, checkCounter,"anmäkda vs checkcount");

                driver.Navigate().Back();

            }

            foreach (KeyValuePair<int, string> kvp in _classes)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _linf)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _horse)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _clubs)
            {
                string country = "SE";
                string club = kvp.Value;
                string clubLowerCase = club.ToLower();

                switch (clubLowerCase)
                {
                    // Check country
                    case "denmark":
                        club = "??";
                        country = "DK";
                        break;
                    case "norway":
                        club = "??";
                        country = "NO";
                        break;
                    case "finland":
                        club = "??";
                        country = "FI";
                        break;
                    case "okänd klubb":
                        club = "??";
                        country = "??";
                        break;
                }

                Trace.WriteLine(kvp.Key + "|" + club + "|" + country);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _comp)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");



        }
    }
}


