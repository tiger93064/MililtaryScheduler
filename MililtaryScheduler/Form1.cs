using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.IO;       //使用System.IO.MemoryStream
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;

//MilitaryShcheduler is a Crew-Scheduler application that assists Yuntech Military group to dispatch crews in 4 certain preiod.
//Support crossing years arrangement, web crawler from yuntech calendar webpage and auto-detected takeoffType, editable crew lists within dynamic arrangement algo
//  , export formatted Excel file.
//
//The porgram can divide into serval parts: generate selected year calander, web crawler from yuntech calender webpage then assgin to local calender, process exception event to support requirment
//  specific three-day continous holidays(takeofftype=2), dispatch crew by textbox name list, print calendar to screen. Export excel file.
//
//All codes below credited to Guanting, Liu 12,10, 2019.

namespace MililtaryScheduler
{
    public partial class Form1 : Form
    {
        int[] dayOfaMonth = { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        string[] weekdayCHTtable = { "日", "一", "二", "三", "四", "五", "六" };
        List<string> weekdayENGtable = new List<string>() { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
        List<day[,]> cals = new List<day[,]>();
        List<int> comb1Data,comb2Data;

        List<List<day>> takeoffDaysYear = new List<List<day>>();
        List<List<day>> excepofDaysYear = new List<List<day>>();

        List<day> excepofAnniverasry = new List<day>();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime d = new DateTime(2020, 1, 1);                                                               //dynamicly set comboBox selection from 2019 to "now year + 1".

            comb1Data = new List<int>() { 2019 };                          //generate combox1 dataSrc.
            for (int i = comb1Data[0] + 1; i <= int.Parse(DateTime.Now.ToString("yyyy")) + 1; i++) comb1Data.Add(i);
            comboBox1.DataSource = comb1Data;
            comboBox1.SelectedIndex = 0;
            //comboBox1.SelectedIndex = comb1Data.IndexOf(int.Parse(DateTime.Now.ToString("yyyy")));
                                                                                                                  //
            generateCal();                                                                                        //generate calendar initially.
            
            

        }
        private void generateCal()
        {
            cals = new List<day[,]>();
            DateTime dtemp = new DateTime(int.Parse(comboBox1.Text), 1, 1);                                       //set pivot to 2019/1/1 weekday number by English lookup table.
            int pivot = weekdayENGtable.FindIndex(a => a.Contains(dtemp.DayOfWeek.ToString()));                   //

            for (int i = int.Parse(comboBox1.Text); i <= int.Parse(comboBox2.Text); i++) {                        //generate cal from combobox1 selected year to combobox2 selected years.
                                                                                                                  //detect Leap year(閏年)
                if ((i % 4 == 0 && i % 100 != 0) || (i % 400 == 0 && i % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };               //
                day[,] calendar = new day[13, 32];
                for (int x = 1; x < 13; x++) {                                                                    //assgin weekday to local calendar within pivot.
                    for (int y = 1; y <= dayOfaMonth[x]; y++) {
                        calendar[x, y] = new day(x, y, weekdayCHTtable[pivot % 7]);
                        pivot++;

                    }
                }
                cals.Add(calendar);                                                                                //cals is an array saved generated year.
            }
            
            //foreach (day[,] t in cals) Console.WriteLine(t[1, 1].weekday);
            
        }
        private void webCrawler()
        {                                                                                                         //web crawler section from yuntech online calendar
                                                                                                                  //here:"https://events.yuntech.edu.tw/index.php?&y=2020&view=YunTech&"
                                                                                                                  //this method is valid until webpage structure changed.
            takeoffDaysYear = new List<List<day>>();
            excepofDaysYear = new List<List<day>>();

            for (int i = int.Parse(comboBox1.Text); i <= int.Parse(comboBox2.Text); i++) {                        //loop from combox1 selection to combox2 selection.
                List<day> takeOffDays = new List<day>();                                      //yuntech holidays.
                List<day> excepTakeoffDays = new List<day>();                                 //補行上班.

                WebClient wC = new WebClient();
                MemoryStream memoryStream  = new MemoryStream(wC.DownloadData("https://events.yuntech.edu.tw/index.php?&y="+i.ToString()+"&view=YunTech&"));        //load page by i(year number).
                //Console.WriteLine("Hi");
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(memoryStream, Encoding.UTF8);
                //textBox2.Text = doc.Text;

                HtmlNode node = doc.DocumentNode.SelectSingleNode("//div[@class='container']/div[@class='row']/div[@class='col-12']/div[@class='yuntech_cal']/div[@class='container']");
                //Console.WriteLine(node.InnerHtml);

                HtmlNodeCollection nameNodes = doc.DocumentNode.SelectNodes("//div[@class='container']/div[@class='row']/div[@class='col-12']/div[@class='yuntech_cal']/div[@class='container']/div");
                //Console.WriteLine(nameNodes.Count);
                //Console.WriteLine();
                foreach (HtmlNode n in nameNodes)
                {

                    HtmlNodeCollection nodesOfevents = n.SelectNodes("./div[2]/div/div/div");
                    if (nodesOfevents == null) continue;
                    //Console.WriteLine(nodesOfevents.Count);
                    foreach (HtmlNode nt in nodesOfevents)
                    {
                        if (nt.Attributes["class"].Value == "YunTech_calendar_holiday row")                      //specific holiday by certain html class
                        {
                            //Console.WriteLine(nt.SelectSingleNode("./div[1]").InnerText + " " + nt.SelectSingleNode("./div[2]").InnerText);
                            takeOffDays.Add(new day(nt.SelectSingleNode("./div[1]").InnerText, nt.SelectSingleNode("./div[2]").InnerText));
                        }
                        else
                        {                                                                                        //specific 補行上班日, because Sat. Sun. initaillize as holiday(takeofftype 1).
                            if (nt.SelectSingleNode("./div[2]").InnerText.Contains("補行"))
                            {
                                //Console.WriteLine(nt.SelectSingleNode("./div[1]").InnerText + " " + nt.SelectSingleNode("./div[2]").InnerText);
                                excepTakeoffDays.Add(new day(nt.SelectSingleNode("./div[1]").InnerText, nt.SelectSingleNode("./div[2]").InnerText));


                            }                                                                                   //
                                                                                                                //specific 校慶, bcuz html class is holiday but militarySection works normally on that day.
                            else if (nt.SelectSingleNode("./div[2]").InnerText.Contains("校慶") && ( nt.SelectSingleNode("./div[2]").InnerText.Contains("大會")|| nt.SelectSingleNode("./div[2]").InnerText.Contains("園遊"))){
                                excepofAnniverasry.Add(new day(nt.SelectSingleNode("./div[1]").InnerText, nt.SelectSingleNode("./div[2]").InnerText));
                                excepofAnniverasry[excepofAnniverasry.Count-1].weekday = cals[i - int.Parse(comboBox1.Text)][excepofAnniverasry[excepofAnniverasry.Count - 1].month, excepofAnniverasry[excepofAnniverasry.Count - 1].date].weekday;
                            }
                            if (nt.SelectSingleNode("./div[2]").InnerText.Contains("運動會")&& !nt.SelectSingleNode("./div[2]").InnerText.Contains("補")) {     //trickly assgin to calendar process later.
                                day d = new day((nt.SelectSingleNode("./div[1]").InnerText), nt.SelectSingleNode("./div[2]").InnerText);
                                cals[i - int.Parse(comboBox1.Text)][d.month, d.date].sEvent = d.sEvent;
                            }
                            if (nt.SelectSingleNode("./div[2]").InnerText.Contains("勞動"))                                                                      //trickly assgin to calendar process later.
                            {
                                day d = new day((nt.SelectSingleNode("./div[1]").InnerText), nt.SelectSingleNode("./div[2]").InnerText);
                                cals[i - int.Parse(comboBox1.Text)][d.month, d.date].sEvent = d.sEvent;
                            }
                        }
                        

                    }

                    // HtmlNodeCollection nodeOfDays = n.SelectSingleNode("div[@class='col - xl - 9 col - lg - 9 col - md - 8 col - sm - 6 col - 12']/div[@class='w-100']/div").SelectNodes("/div");
                    //Console.WriteLine(nodeOfDays.Count);
                }

                foreach (day d in takeOffDays)                                      //assign weekday to takeoffDays[].
                {
                    d.weekday = cals[i - int.Parse(comboBox1.Text)][d.month, d.date].weekday;
                }
                foreach (day d in excepTakeoffDays)                                //assign weekday to excepTakeoffDays[].
                {
                    d.weekday = cals[i - int.Parse(comboBox1.Text)][d.month, d.date].weekday;
                    d.takeoffType = 0;
                }

                takeoffDaysYear.Add(takeOffDays);
                excepofDaysYear.Add(excepTakeoffDays);

            }
        }

        private void assignWCresult()                                                                           //assign wecCrawler result to local calendar.
        {
            for (int i = 0; i <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); i++)
            {
                foreach (day d in takeoffDaysYear[i])
                {
                    cals[i][d.month, d.date] = d;
                }
                foreach (day d in excepofDaysYear[i])
                {
                    cals[i][d.month, d.date] = d;
                }
                if(i<excepofAnniverasry.Count) cals[i][excepofAnniverasry[i].month, excepofAnniverasry[i].date] = excepofAnniverasry[i];
            }
        }

        private void doException()
        {
            //"補行上班日" had been change takeofftype to 0 right after create object. so do not process like an exception.

            if(checkBox1.Checked) exp1();               //校慶 treats like takeofftype=0 and the right after it day.takeofftype=0 change to 1.
            exp2();                                     //point out 春節期間 and set takeofftype to 3;
            if (checkBox2.Checked) exp3();                                     //運動會 takeofftype = 0
            if (checkBox3.Checked) exp4();
            if (checkBox4.Checked) exp5();


            void exp1() {
                for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                {
                    bool key = false, toNextY = false;
                    int a = int.Parse(comboBox1.Text) + x;
                    if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                    //Console.WriteLine(x);
                    for (int i = 1; i < 13; i++)
                    {                        
                        for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                        {
                            if (key && cals[x][i, j].takeoffType == 0) {
                                cals[x][i, j].takeoffType = 1;
                                toNextY = true;
                                break;
                            }
                            if (cals[x][i, j].sEvent.Contains("校慶") && (cals[x][i, j].sEvent.Contains("大會") || cals[x][i, j].sEvent.Contains("園遊")))
                            {
                                cals[x][i, j].takeoffType = 0;
                                key = true;
                            }
                        }
                        if (toNextY) break;
                    }
                }
            }
            void exp2()
            {
                for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                {
                    bool key = false, key2 = false, continuer = false;
                    int a = int.Parse(comboBox1.Text) + x;
                    if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    day yesterday = new day(12,31,"y"), yyesterday = new day(12, 30, "yy");

                    for (int i = 1; i < 13; i++)
                    {
                        if (continuer) break;
                        for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                        {
                            if (key) {
                                if (key2) {
                                    if (cals[x][i, j].takeoffType == 1)
                                    {
                                        cals[x][i, j].takeoffType = 3;

                                        yyesterday = yesterday;
                                        yesterday = new day(i, j, "");
                                        continue;
                                    }
                                    else {
                                        continuer = true;
                                        break;
                                    }
                                    
                                    
                                }
                                cals[x][i, j].takeoffType = 3;
                                if ((cals[x][i, j].sEvent.Contains("除夕") || cals[x][i, j].sEvent.Contains("春節") || cals[x][i, j].sEvent.Contains("春假")) && cals[x][i, j].sEvent.Contains("結束"))key2 = true;


                                yyesterday = yesterday;
                                yesterday = new day(i, j, "");
                                continue;
                            }
                            if ((cals[x][i, j].sEvent.Contains("除夕") || cals[x][i, j].sEvent.Contains("春節") || cals[x][i, j].sEvent.Contains("春假")) && cals[x][i, j].sEvent.Contains("開始"))
                            {
                                key = true;
                                cals[x][i, j].takeoffType = 3;
                                
                                if (cals[x][yesterday.month, yesterday.date].takeoffType == 1) cals[x][yesterday.month, yesterday.date].takeoffType = 3;
                                if (cals[x][yyesterday.month, yyesterday.date].takeoffType == 1) cals[x][yyesterday.month, yyesterday.date].takeoffType = 3;
                            }

                            yyesterday = yesterday;
                            yesterday = new day(i, j, "");
                        }
                    }
                }
            }
            void exp3() {
                for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                {
                    int a = int.Parse(comboBox1.Text) + x;
                    if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                    for (int i = 1; i < 13; i++)
                    {
                        for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                        {
                            if (cals[x][i, j].sEvent.Contains("運動會")&& !cals[x][i, j].sEvent.Contains("補")) cals[x][i, j].takeoffType = 0;
                        }
                    }
                }
            }
            void exp4() {
                for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                {
                    int a = int.Parse(comboBox1.Text) + x;
                    if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                    for (int i = 1; i < 13; i++)
                    {
                        for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                        {
                            if (cals[x][i, j].sEvent.Contains("校慶") && cals[x][i, j].sEvent.Contains("補")) cals[x][i, j].takeoffType = 0;
                        }
                    }
                }
            }
            void exp5()
            {
                for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                {
                    int a = int.Parse(comboBox1.Text) + x;
                    if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                    for (int i = 1; i < 13; i++)
                    {
                        for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                        {
                            if (cals[x][i, j].sEvent.Contains("勞動")) cals[x][i, j].takeoffType = 1;
                        }
                    }
                }
            }
        }
        private void findVacation() {

            for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
            {                
                int a = int.Parse(comboBox1.Text) + x;
                if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                int count = 0;
                List<day> temp = new List<day>();
                for (int i = 1; i < 13; i++)
                {
                    for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                    {
                        if (cals[x][i, j].takeoffType == 1)
                        {
                            count++;
                            temp.Add(cals[x][i, j]);
                            if (i == 12 && j == 31 && count >= 3) foreach (day d in temp) cals[x][d.month, d.date].takeoffType = 2;
                        }
                        else
                        {
                            if (count >= 3)
                            {
                                foreach (day d in temp) cals[x][d.month, d.date].takeoffType = 2;
                            }
                            count = 0;
                            temp.Clear();
                        }

                    }
                }
            }

            for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
            {
                int a = int.Parse(comboBox1.Text) + x;
                if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                int count = 0;
                bool breaker = false;
                List<day> tempCrossYear = new List<day>();
                for (int i = 1; i < 13; i++)
                {
                    if (breaker) break;
                    for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                    {
                        if (cals[x][i, j].takeoffType == 1) tempCrossYear.Add(cals[x][i, j]);
                        else {
                            breaker = true;
                            break;
                        }
                    }
                }

                List<day> tempLastYear = new List<day>();
                if (tempCrossYear.Count >= 1) {
                    for (int i = 31; i > 0; i--) {
                        DateTime tempDT = new DateTime(int.Parse(comboBox1.Text) + x-1, 12, i);
                        if (tempDT.DayOfWeek.ToString() == "Sunday" || tempDT.DayOfWeek.ToString() == "Saturday")
                        {
                            tempLastYear.Add(new day(12, i, ""));
                        }
                        else break;
                    }
                }
                if ((tempCrossYear.Count+tempLastYear.Count) >= 3)
                {
                    foreach (day d in tempCrossYear) cals[x][d.month, d.date].takeoffType = 2;
                    foreach (day d in tempLastYear) {
                        if(x!=0) cals[x-1][d.month, d.date].takeoffType = 2;
                    }
                }

            }
        }
        private void dispatchCrew() {
            string[] normalDaysCrew = textBox1.Text.Split(',');
            string[] weekendDaysCrew = textBox3.Text.Split(',');
            string[] holidayDaysCrew = textBox2.Text.Split(',');
            int pivot= 0, pivot1 = 0, pivot2 = 0, pivot3 = 0;

            for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
            {
                int a = int.Parse(comboBox1.Text) + x;
                if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                for (int i = 1; i < 13; i++)
                {
                    for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                    {
                        switch (cals[x][i, j].takeoffType)
                        {
                            case 0:
                                cals[x][i, j].crew = normalDaysCrew[pivot % normalDaysCrew.Length];
                                pivot++;
                                break;
                            case 1:
                                cals[x][i, j].crew = weekendDaysCrew[pivot1 % weekendDaysCrew.Length];
                                pivot1++;
                                break;
                            case 2:
                                cals[x][i, j].crew = holidayDaysCrew[pivot2 % holidayDaysCrew.Length];
                                pivot2++;
                                break;
                            case 3:
                                break;

                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label3.Text = comboBox2.Text +" 年班表";
            generateCal();
            webCrawler();
            assignWCresult();
            doException();
            findVacation();

            dispatchCrew();
            printCalendar();

            button2.Enabled = true;
        }

        private void generateExcel() {
            string pathFile = @"D:\test";

            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();

            // 讓Excel文件可見
            excelApp.Visible = true;

            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);


            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];

            // 設定活頁簿焦點
            wBook.Activate();

            try
            {
                //for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                for (int x = int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
                {
                    //int a = int.Parse(comboBox1.Text) + x;
                    int a = int.Parse(comboBox1.Text) + x;
                    if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                    else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                    Excel.Worksheet newWorksheet;
                    for (int i = 1; i < 13; i++)
                    {
                        //Add a worksheet to the workbook.
                        newWorksheet = (Excel.Worksheet)excelApp.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        //Name the sheet.
                        newWorksheet.Name = a.ToString()+" " + i.ToString()+"月";
                        newWorksheet.Activate();

                        string[] Crews = new string[] { "少校組長\n廖靜婕", "中校教官\n留濰旻", "中校教官\n謝艷芬", "少校教官\n黃麗穎", "中校教官\n林淑真", "校安助理\n彭啟禎", "校安助理\n丁儀偉", "校安助理\n郭宗廷", "校安助理\n胡景龍" };

                        wRange = newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[1, Crews.Count()+3]].merge();
                        excelApp.Cells[1, 1] = "國立雲林科技大學"+(a-1911).ToString()+"年"+i.ToString()+"月份軍訓教官及校安人員輪班表(預排)";
                        excelApp.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        excelApp.Cells[1, 1].Font.Name = "DFKai-SB";
                        excelApp.Cells[1, 1].Font.Size = 14;


                        wRange = newWorksheet.Range[newWorksheet.Cells[2, 1], newWorksheet.Cells[2,2]].merge();

                        wRange = newWorksheet.Range[newWorksheet.Cells[2, 3], newWorksheet.Cells[2, Crews.Count() + 3]];
                        wRange.Select();
                        wRange.Rows.AutoFit();
                        for (int j = 3; j <= Crews.Count()+3; j++)
                        {
                            excelApp.Cells[2, j].Font.Name = "DFKai-SB";
                            excelApp.Cells[2, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            if (j== Crews.Count() + 3)
                            {
                                excelApp.Cells[2, j] = "備註";
                                
                            }
                                
                            else
                            {
                                excelApp.Cells[2, j] = Crews[j - 3];
                                excelApp.Cells[2, j].Font.Size = 11;
                            }
                                

                            

                        }

                        //wRange.Select();
                        //wRange.Font.Color = ColorTranslator.ToOle(Color.White);

                        for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                        {
                            excelApp.Cells[j+2, 1] = j.ToString();
                            excelApp.Cells[j+2, 2] = " "+cals[x][i, j].weekday+" ";
                            if(cals[x][i, j].takeoffType == 1 || cals[x][i, j].takeoffType == 2) excelApp.Cells[j + 2, 2].Font.Color = ColorTranslator.ToOle(Color.Red);


                            for (int k = 3; k < Crews.Count() + 3; k++) {
                                
                                if(cals[x][i, j].takeoffType==1 || cals[x][i, j].weekday.Equals("六") || cals[x][i, j].weekday.Equals("日")) excelApp.Cells[j + 2, k].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                else if(cals[x][i, j].takeoffType == 2 && !cals[x][i, j].weekday.Equals("六") && !cals[x][i, j].weekday.Equals("日")) excelApp.Cells[j + 2, k].Interior.Color = ColorTranslator.ToOle(Color.Orange);
                                else if(cals[x][i, j].takeoffType == 3) excelApp.Cells[j + 2, k].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                                try
                                {
                                    if (excelApp.Cells[2, k].Text.Contains(cals[x][i, j].crew))
                                    {
                                        excelApp.Cells[j + 2, k] = "◎";
                                        excelApp.Cells[j + 2, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    }
                                }
                                catch { }
                            }
                        }
                        string[] footerText = new string[] { "值勤總時數","教官超過80小時可補休\n加班時數", "校安加班總時數\n平日4小時假日12小時", "1-3月校安加班總時數", "承辦人:" };
                        for (int z = 1; z < 6; z++) {
                            wRange = newWorksheet.Range[newWorksheet.Cells[dayOfaMonth[i]+z+2, 1], newWorksheet.Cells[dayOfaMonth[i]+z+2, 3]].merge();
                            excelApp.Cells[dayOfaMonth[i] + z + 2, 1].ColumnWidth = 10;
                            excelApp.Cells[dayOfaMonth[i] + z + 2, 1] = footerText[z - 1];
                            excelApp.Cells[dayOfaMonth[i] + z + 2, 1].Font.Size = 9;
                            if (z==1) excelApp.Cells[dayOfaMonth[i] + z + 2, 1].Font.Size = 12;
                            else if (z==5) excelApp.Cells[dayOfaMonth[i] + z + 2, 1].Font.Size = 11;

                            excelApp.Rows[10].RowHeight = 18;
                        }
                        // 設定第1列資料
                        //excelApp.Cells[1, 1] = "名稱";
                        //excelApp.Cells[1, 2] = "數量";
                        



                        // 設定第5列資料
                        //excelApp.Cells[35, 1] = "總計";
                        // 設定總和公式 =SUM(B2:B4)
                        excelApp.Cells[35, 2].Formula = string.Format("=SUM(B{0}:B{1})", 2, 4);
                        // 設定第5列顏色
                        //wRange = newWorksheet.Range[newWorksheet.Cells[35, 1], newWorksheet.Cells[35, 2]];
                        //wRange.Select();
                        //wRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                        //wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);

                        // 自動調整欄寬
                        wRange = newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[dayOfaMonth[i] + 6, Crews.Count()+3]];
                        wRange.Select();
                        wRange.Columns.AutoFit();
                        wRange.Rows.AutoFit();
                        wRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        wRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                        wRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        //wRange = newWorksheet.Range[newWorksheet.Cells[6, 1], newWorksheet.Cells[6, 2]].merge();

                        //Get the Cells collection.             //method 2
                        Excel.Range cells = newWorksheet.Cells;

                        //Input a string value to a cell of the sheet.
                        //cells.set_Item(i, i, "New_Sheet" + i.ToString());
                        
                    }
                }

                for (int i = 1; i < 6; i++)
                {

                    
            }
                

                
                

                try
                {
                    //另存活頁簿
                    wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }

            //關閉活頁簿
            //wBook.Close(false, Type.Missing, Type.Missing);

            //關閉Excel
            //excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();
        }
        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button2.Enabled = false;

            comb2Data = new List<int>();                          //generate combox1 dataSrc.
            for (int i = comb1Data[comboBox1.SelectedIndex]; i <= int.Parse(DateTime.Now.ToString("yyyy")) + 4; i++) comb2Data.Add(i);
            comboBox2.DataSource = comb2Data;
            comboBox2.SelectedIndex = comb2Data.IndexOf(int.Parse(DateTime.Now.ToString("yyyy")));
            if (comboBox2.SelectedIndex == -1) comboBox2.SelectedIndex = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            generateExcel();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button2.Enabled = false;
        }

        public void printCalendar()
        {
            textBox4.Text = "";
            for (int x = 0; x <= int.Parse(comboBox2.Text) - int.Parse(comboBox1.Text); x++)
            {
                int a = int.Parse(comboBox1.Text) + x;
                if ((a % 4 == 0 && a % 100 != 0) || (a % 400 == 0 && a % 4000 != 0)) dayOfaMonth = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                else dayOfaMonth = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

                for (int i = 1; i < 13; i++)
                {
                    for (int j = 1; j < dayOfaMonth[i] + 1; j++)
                    {
                        textBox4.Text += i + " / " + j + "  (" + cals[x][i, j].weekday + ") " + cals[x][i, j].takeoffType + " " + cals[x][i, j].crew + " " + cals[x][i, j].sEvent + Environment.NewLine;
                        Console.WriteLine(i + " " + j + "  (" + cals[x][i, j].weekday + ") " + cals[x][i, j].takeoffType + " "+cals[x][i,j].crew+" " + cals[x][i, j].sEvent);
                    }
                    if (i == 12) textBox4.Text += "-------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                }
            }
            
        }
    }

    public class day
    {        
        public day(String MonthDate, string Event)
        {
            string[] sArray = MonthDate.Split('/');
            month = int.Parse(sArray[0]);
            date = int.Parse(sArray[1]);
            sEvent = Event;

            takeoffType = 1;
        }
        public day(int m, int d, string w)
        {
            month = m;
            date = d;
            weekday = w;
            sEvent = "";


            if (weekday == "六" || weekday == "日") takeoffType = 1;
            else takeoffType = 0;
        }
        public int month { get; set; }
        public int date { get; set; }
        public string weekday { get; set; }
        public string monDate { get; set; }                 //optional
        public string sEvent { get; set; }                  //optional
        public string crew { get; set; }                    //optional
        public int takeoffType { get; set; }                // workday = 0, holiday = 1, continous holiday = 2, new year holidays = 3  

    }
}
