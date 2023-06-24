using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.IO;
using System.Xml.Linq;
using OpenQA.Selenium.Support.UI;

class Program
{
    static void Main()
    {
        string inputFile = "C:\\Users\\User\\Desktop\\data2.txt";
        string outputFile = "C:\\Users\\User\\Desktop\\data.csv";
        string eiin = string.Empty;
        string institutName = string.Empty;
        int scTotal = 0;
        int bsTotal = 0;
        int humTotal = 0;
        int scPassedCount = 0;
        int scFailedCount = 0;
        int scGpa = 0;
        int bsPassedCount = 0;
        int bsFailedCount = 0;
        int bsGpa = 0;
        int humPassedCount = 0;
        int humFailedCount = 0;
        int humGpa = 0;
        int count = 0;

        var holdingOrRoadNo = string.Empty;
        var postOffice = string.Empty;
        var postCode = string.Empty;
        var division = string.Empty;
        var zilla = string.Empty;
        var upazilla = string.Empty;
        var ward = string.Empty;
        var mouja = string.Empty;
        var mobile = string.Empty;
        var alternativeMobile = string.Empty;
        var phone = string.Empty;
        var email = string.Empty;
        var website = string.Empty;
        var typeOfInstitute = string.Empty;
        var levelOfInstitute = string.Empty;
        var management = string.Empty;
        var instituteArea = string.Empty;

        try
        {
            // Read the data from the text file
            string[] lines = File.ReadAllLines(inputFile);

            IWebDriver driver = new ChromeDriver();

            // Create the CSV file and write the headers
            using (StreamWriter writer = new StreamWriter(outputFile, true))
            {
                //writer.WriteLine("Eiin,Name, Total, BS Total, BS pass,BS Fail,BS GPA5, SC Total, SC pass,SC Fail,SC GPA5, HUM Total, HUM pass,HUM Fail,HUM GPA5, Holding Or Road No, Post Office, Post Code, Division, Zilla, Upazilla/Thana, Ward, Mouja, Mobile, Alternative Mobile, Phone, Email, Website, Is Website Live, Type Of Institute, Management, Institute Area");

                // Process each line of data
                foreach (string line in lines)
                {
                    // Split the line by spaces
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        // get data using selenium
                        driver.Navigate().GoToUrl("http://202.72.235.210/ban_survey/public/institute-search");
                        var element = driver.FindElement(By.Id("eiin"));
                        element.SendKeys(eiin);
                        element.Submit();
                        var list = new List<string>();

                        var ele = driver.FindElement(By.CssSelector("#firstPage > div > div:nth-child(3) > div.contentBoxBody > table:nth-child(2) > tbody"));
                        var trElement = ele.FindElements(By.TagName("tr"));
                        int counts = 0;
                        foreach (var tr in trElement)
                        {
                            if (counts > 0)
                            {
                                var tdElement = tr.FindElements(By.TagName("td"));
                                foreach (var td in tdElement)
                                {
                                    var lable = td.FindElement(By.TagName("label")).Text;
                                    if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "হোল্ডিং নম্বর/রোড:")
                                        holdingOrRoadNo = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "ডাকঘর:")
                                        postOffice = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "পোস্ট কোড:")
                                        postCode = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "বিভাগ:")
                                        division = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "জেলা:")
                                        zilla = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "উপজেলা/থানা:")
                                        upazilla = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "ওয়ার্ড:")
                                    {
                                        var wardElement = td.FindElement(By.TagName("select"));
                                        if (wardElement.Selected)
                                        {
                                            var wardSelect = new SelectElement(wardElement);
                                            var wardName = wardSelect.SelectedOption?.Text;
                                            ward = wardName;
                                        }
                                    }
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "মৌজা:")
                                        mouja = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "মোবাইল নম্বর:")
                                        mobile = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "বিকল্প মোবাইর নম্বর:")
                                        alternativeMobile = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "ফোন:")
                                        phone = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "ই-মেইল:")
                                        email = td.FindElement(By.TagName("input")).GetAttribute("value");
                                    else if (string.IsNullOrWhiteSpace(lable) == false && lable.Trim() == "ওয়েবসাইট:")
                                        website = td.FindElement(By.TagName("input")).GetAttribute("value");
                                }
                            }
                            counts++;
                        }

                        var typeOfInstituteElement = driver.FindElement(By.CssSelector(".w-100 > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > select:nth-child(2)"));
                        if (typeOfInstituteElement.Selected)
                        {
                            var typeOfInstituteSelect = new SelectElement(typeOfInstituteElement);
                            typeOfInstitute = typeOfInstituteSelect.SelectedOption.Text;
                        }

                        var levelOfInstituteElement = driver.FindElement(By.CssSelector(".w-100 > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > select:nth-child(2)"));
                        if (levelOfInstituteElement.Selected)
                        {
                            var lvlOfInsSelect = new SelectElement(levelOfInstituteElement);
                            levelOfInstitute = lvlOfInsSelect.SelectedOption.Text;
                        }

                        var managementElement = driver.FindElement(By.CssSelector("#firstPage > div:nth-child(1) > div:nth-child(4) > div:nth-child(3) > div:nth-child(2) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > select:nth-child(2)"));
                        if (managementElement.Selected)
                        {
                            var managementSelect = new SelectElement(managementElement);
                            management = managementSelect.SelectedOption.Text;
                        }

                        var instituteAreaElement = driver.FindElement(By.CssSelector("#firstPage > div:nth-child(1) > div:nth-child(4) > div:nth-child(3) > div:nth-child(2) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > select:nth-child(2)"));
                        if (instituteAreaElement.Selected)
                        {
                            var areaSelect = new SelectElement(instituteAreaElement);
                            instituteArea = areaSelect.SelectedOption.Text;
                        }

                        // Write the data to the CSV file
                        writer.WriteLine($"{eiin},{institutName},{bsTotal+scTotal+humTotal},{bsTotal},{bsPassedCount},{bsFailedCount},{bsGpa},{scTotal},{scPassedCount},{scFailedCount},{scGpa},{humTotal},{humPassedCount},{humFailedCount},{humGpa},{holdingOrRoadNo},{postOffice},{postCode},{division},{zilla},{upazilla},{ward},{mouja},{mobile},{alternativeMobile},{phone},{email},{website},{"No"},{typeOfInstitute},{management},{instituteArea}");

                        eiin = string.Empty;
                        institutName = string.Empty;
                        scPassedCount = 0;
                        scFailedCount = 0;
                        scGpa = 0;
                        bsPassedCount = 0;
                        bsFailedCount = 0;
                        bsGpa = 0;
                        humPassedCount = 0;
                        humFailedCount = 0;
                        humGpa = 0;

                        holdingOrRoadNo = string.Empty;
                        postOffice = string.Empty;
                        postCode = string.Empty;
                        division = string.Empty;
                        zilla = string.Empty;
                        upazilla = string.Empty;
                        ward = string.Empty;
                        mouja = string.Empty;
                        mobile = string.Empty;
                        alternativeMobile = string.Empty;
                        phone = string.Empty;
                        email = string.Empty;
                        website = string.Empty;
                        typeOfInstitute = string.Empty;
                        levelOfInstitute = string.Empty;
                        management = string.Empty;
                        instituteArea = string.Empty;
                    }
                    else
                    {
                        string[] parts = line.Split(':');

                        // Extract the institution ID
                        string tagName = parts[0].Trim();

                        if (tagName == "INST")
                        {
                            for (int i = 1; i < parts.Length; i++)
                            {
                                string[] subjectParts = parts[i].Split('-');
                                eiin = subjectParts[0].Trim();
                                institutName = subjectParts[1].Trim().Replace(",", "");
                            }
                        }
                        else if (tagName == "BUSINESS STUDIES")
                        {
                            for (int i = 1; i < parts.Length; i++)
                            {
                                string[] subjectParts = parts[i].Split(';');

                                for (int j = 0; j < subjectParts.Length; j++)
                                {
                                    string[] subPart = subjectParts[j].Split("=");
                                    if (subPart != null &&  subPart.Length > 0)
                                    {
                                        string tagStr = subPart[0].Trim();
                                        if (tagStr == "PASSED")
                                            bsPassedCount = int.Parse(subPart[1]);
                                        else if (tagStr == "NOT PASSED")
                                            bsFailedCount = int.Parse(subPart[1]);
                                        else if (tagStr == "GPA5")
                                            bsGpa = int.Parse(subPart[1]);
                                    }
                                }
                            }
                            bsTotal = bsPassedCount + bsFailedCount;
                        }
                        else if (tagName == "HUMANITIES")
                        {
                            for (int i = 1; i < parts.Length; i++)
                            {
                                string[] subjectParts = parts[i].Split(';');

                                for (int j = 0; j < subjectParts.Length; j++)
                                {
                                    string[] subPart = subjectParts[j].Split("=");
                                    if (subPart != null &&  subPart.Length > 0)
                                    {
                                        string tagStr = subPart[0].Trim();
                                        if (tagStr == "PASSED")
                                            humPassedCount = int.Parse(subPart[1]);
                                        else if (tagStr == "NOT PASSED")
                                            humFailedCount = int.Parse(subPart[1]);
                                        else if (tagStr == "GPA5")
                                            humGpa = int.Parse(subPart[1]);
                                    }
                                }
                            }
                            humTotal = humPassedCount + humFailedCount;
                        }
                        else if (tagName == "SCIENCE")
                        {
                            for (int i = 1; i < parts.Length; i++)
                            {
                                string[] subjectParts = parts[i].Split(';');

                                for (int j = 0; j < subjectParts.Length; j++)
                                {
                                    string[] subPart = subjectParts[j].Split("=");
                                    if (subPart != null &&  subPart.Length > 0)
                                    {
                                        string tagStr = subPart[0].Trim();
                                        if (tagStr == "PASSED")
                                            scPassedCount = int.Parse(subPart[1]);
                                        else if (tagStr == "NOT PASSED")
                                            scFailedCount = int.Parse(subPart[1]);
                                        else if (tagStr == "GPA5")
                                            scGpa = int.Parse(subPart[1]);
                                    }
                                }
                            }
                            scTotal = scPassedCount + scFailedCount;
                        }
                    }
                }
                //writer.WriteLine($"{eiin},{institutName},{bsPassedCount},{bsFailedCount},{bsGpa},{scPassedCount},{scFailedCount},{scGpa},{humPassedCount},{humFailedCount},{humGpa}");
            }

            Console.WriteLine("CSV file created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
