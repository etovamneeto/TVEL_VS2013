using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using RiskCalculatorLib;
using Excel = Microsoft.Office.Interop.Excel;
//Хреновины имеют значение, когда сравниваешь строки и строки, т.е. когда тебе надо ввести в текстовую ячейку текстовый параметр; В другие, где инты или флоаты, туда не нужны они
//"UPDATE [names] SET [userName]='" + nameBox.Text + "', [age]=" + Convert.ToInt32(ageBox.Text), но + " WHERE [id]=" + Convert.ToInt32(table.Rows[0]["id"]), connection);

namespace TVELtest
{
    public partial class Form1 : Form
    {
        /*-----Описание класса Человек, в котором хранится информация: id, пол, возраст при облучении, дозовая история-----*/
        public class Man
        {
            private int id = 0;
            private byte sex = 0;
            private short ageAtExp = 0;
            private RiskCalculator.DoseHistoryRecord[] doseHistory = null;

            public Man(int id, byte sex, short ageAtExp, RiskCalculator.DoseHistoryRecord[] doseHistory)
            {
                this.id = id;
                this.sex = sex;
                this.ageAtExp = ageAtExp;
                this.doseHistory = doseHistory;
            }

            public int getID() { return this.id; }
            public byte getSex() { return this.sex; }
            public short getAgeAtExp() { return this.ageAtExp; }
            public RiskCalculator.DoseHistoryRecord[] getDoseHistory() { return this.doseHistory; }

            public void setID(int id) { this.id = id; }
            public void setSex(byte sex) { this.sex = sex; }
            public void setAgeAtExp(short ageAtExp) { this.ageAtExp = ageAtExp; }
            public void getDoseHistory(RiskCalculator.DoseHistoryRecord[] doseHistory) { this.doseHistory = doseHistory; }
        }

        /*-----Описание класса Объект, представляющий собой строку таблицы с параметрами: id, пол, доза суммарная, доза внутренняя, возраст при облучении-----*/
        public class dbObject
        {
            private int id = 0;
            private short ageAtExp = 0;
            private double dose = 0;
            private double doseInt = 0;
            private byte sex = 0;
            private int year = 0;

            public dbObject(int id, byte sex, int year, short ageAtExp, double dose, double doseInt)
            {
                this.id = id;
                this.sex = sex;
                this.year = year;
                this.ageAtExp = ageAtExp;
                this.dose = dose;
                this.doseInt = doseInt;
            }

            public void setId(int id) { this.id = id; }
            public void setAgeAtExp(short ageAtExp) { this.ageAtExp = ageAtExp; }
            public void setYear(int year) { this.year = year; }
            public void setDose(double dose) { this.dose = dose; }
            public void setDoseInt(double doseInt) { this.doseInt = doseInt; }
            public void setSex(byte sex) { this.sex = sex; }

            public int getId() { return this.id; }
            public short getAgeAtExp() { return this.ageAtExp; }
            public int getYear() { return this.year; }
            public double getDose() { return this.dose; }
            public double getDoseInt() { return this.doseInt; }
            public byte getSex() { return this.sex; }
        }

        /*-----Описание форм инициализации и инициализация библиотеки с рейтами 2012 года-----*/
        public Form1(String title)
        {
            InitializeComponent();
            this.Text = title;
        }

        String libPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\DataRus2012";

        /*-----Функции для расчета LAR, необходимых для расчета ОРПО*-----*/
        public double getManExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (2 / Math.Pow(10, 6)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (-13 / Math.Pow(10, 4)) * meanAge;
            double constant = 9.36 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (1 / Math.Pow(10, 5)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (-31 / Math.Pow(10, 4)) * meanAge;
            double constant = 17.42 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getManIntLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (-3 / Math.Pow(10, 5)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (23 / Math.Pow(10, 4)) * meanAge;
            double constant = 1.15 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanIntLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (-4 / Math.Pow(10, 5)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (27 / Math.Pow(10, 4)) * meanAge;
            double constant = 5.02 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getOrpo(double lar, double averageDose)
        {
            double orpo = 0;
            orpo = lar * averageDose;
            return orpo;
        }

        public double getOrpo_95(double lar, double averageDose, double deviation)
        {
            double orpo = 0;
            orpo = lar * (averageDose + 1.96 * deviation);
            return orpo;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RiskCalculatorLib.RiskCalculator.FillData(ref libPath);
            this.Width = 300;
            this.Height = 300;
            this.CenterToScreen();
        }

        private void getOrpoButton_Click(object sender, EventArgs e)
        {
            /*-----Инициализация всяких входных параметров, подключения к БД, парсинга в таблицу нужных столбцов-----*/
            String dbPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\dbTvel.mdb";
            String connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbPath;
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Dose], [DoseInt], [Year], [Gender], [BirthYear], [AgeAtExp] FROM [Final] WHERE [Shop]='r3'", connectionString);//Выбор нужных столбцов из нужной таблицы
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet, "Final");
            DataTable table = dataSet.Tables[0];//Из Final в эту таблицу считываются поля, указанные в запросе; Выборка для МСК (shop = r3)

            /*-----Список, в котором хранятся строковые параметры, инентифицирующие возрастные группы-----*/
            List<String> ageGroups = new List<string>();//Строки, в которых указаны возростные группы. Это ключи для дальнейшей связи через словари.
            ageGroups.Add("18-24");
            ageGroups.Add("25-29");
            ageGroups.Add("30-34");
            ageGroups.Add("35-39");
            ageGroups.Add("40-44");
            ageGroups.Add("45-49");
            ageGroups.Add("50-54");
            ageGroups.Add("55-59");
            ageGroups.Add("60-64");
            ageGroups.Add("65-69");
            ageGroups.Add("70+");

            /*-----Список, в котором хранятся нижние границы возрастов для возрастных групп-----*/
            List<int> ageLowerBound = new List<int>();
            ageLowerBound.Add(18);
            ageLowerBound.Add(25);
            ageLowerBound.Add(30);
            ageLowerBound.Add(35);
            ageLowerBound.Add(40);
            ageLowerBound.Add(45);
            ageLowerBound.Add(50);
            ageLowerBound.Add(55);
            ageLowerBound.Add(60);
            ageLowerBound.Add(65);
            ageLowerBound.Add(70);

            /*-----Список, в котором хранятся верхние границы возрастов для возрастных групп-----*/
            List<int> ageUpperBound = new List<int>();
            ageUpperBound.Add(24);
            ageUpperBound.Add(29);
            ageUpperBound.Add(34);
            ageUpperBound.Add(39);
            ageUpperBound.Add(44);
            ageUpperBound.Add(49);
            ageUpperBound.Add(54);
            ageUpperBound.Add(59);
            ageUpperBound.Add(64);
            ageUpperBound.Add(69);
            ageUpperBound.Add(100);

            /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
            List<dbObject> dbRecords = new List<dbObject>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                dbRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
            }

            /*-----Список, в котором хранится пол-----*/
            List<byte> dbSex = new List<byte>();
            for (int i = 0; i < dbRecords.Count; i++)
                dbSex.Add(dbRecords[i].getSex());

            /*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
            byte sexMale = dbSex.Min();
            byte sexFemale = dbSex.Max();

            /*-----Список возрастов облучения из БД для М и Ж-----*/
            List<short> dbManAges = new List<short>();
            for (int i = 0; i < dbRecords.Count; i++)
                if (dbRecords[i].getSex() == sexMale)
                    dbManAges.Add(dbRecords[i].getAgeAtExp());

            List<short> dbWomanAges = new List<short>();
            for (int i = 0; i < dbRecords.Count; i++)
                if (dbRecords[i].getSex() == sexFemale)
                    dbWomanAges.Add(dbRecords[i].getAgeAtExp());

            /*-----Min и Max возраста облучения для М и Ж-----*/
            short manMinAge = dbManAges.Min();
            short manMaxAge = dbManAges.Max();

            short womanMinAge = dbWomanAges.Min();
            short womanMaxAge = dbWomanAges.Max();

            /*-----Массивы списков для мужчин и для женщин, в каждом из которых будут храниться объекты, сгруппированные по возрастам облучения. i = 0 - 18-летние, i = 1 - 19-летние и тд-----*/
            List<dbObject>[] manAgesGroupedArray = new List<dbObject>[manMaxAge - manMinAge + 1];
            for (int i = 0; i < manAgesGroupedArray.Length; i++)
                manAgesGroupedArray[i] = new List<dbObject>();

            List<dbObject>[] womanAgesGroupedArray = new List<dbObject>[womanMaxAge - womanMinAge + 1];
            for (int i = 0; i < womanAgesGroupedArray.Length; i++)
                womanAgesGroupedArray[i] = new List<dbObject>();

            for (int i = manMinAge - manMinAge; i <= manMaxAge - manMinAge; i++)
            {
                for (int k = 0; k < dbRecords.Count; k++)
                {
                    if (dbRecords[k].getSex() == sexMale)
                        if (dbRecords[k].getAgeAtExp() == i + manMinAge)
                        {
                            manAgesGroupedArray[i].Add(dbRecords[k]);
                        }
                }
            }

            for (int i = womanMinAge - womanMinAge; i <= womanMaxAge - womanMinAge; i++)
            {
                for (int k = 0; k < dbRecords.Count; k++)
                {
                    if (dbRecords[k].getSex() == sexFemale)
                        if (dbRecords[k].getAgeAtExp() == i + womanMinAge)
                        {
                            womanAgesGroupedArray[i].Add(dbRecords[k]);
                        }
                }
            }

            /*-----Создание массивов, в которых хранятся суммы средних доз подгрупп, входящих в половозрастную группу-----*/
            double[] manAverDosesExt = new double[manAgesGroupedArray.Length];
            double[] manAverDosesInt = new double[manAgesGroupedArray.Length];
            double[] womanAverDosesExt = new double[womanAgesGroupedArray.Length];
            double[] womanAverDosesInt = new double[womanAgesGroupedArray.Length];

            /*-----Создание массива, в котором хранятся суммы возрастов подгруппы; по факту это ageAtExt * Count-----*/
            double[] manAgeAmountOfSubgroup = new double[manAgesGroupedArray.Length];
            double[] womanAgeAmountOfSubgroup = new double[womanAgesGroupedArray.Length];

            for (int i = 0; i < manAgesGroupedArray.Length; i++)
            {
                for (int n = 0; n < ageGroups.Count; n++)
                    for (int k = 0; k < manAgesGroupedArray[i].Count; k++)
                    {
                        if (manAgesGroupedArray[i][0].getAgeAtExp() >= ageLowerBound[n] && manAgesGroupedArray[i][0].getAgeAtExp() <= ageUpperBound[n])
                        {
                            manAverDosesExt[i] += manAgesGroupedArray[i][k].getDose() - manAgesGroupedArray[i][k].getDoseInt();
                            manAverDosesInt[i] += manAgesGroupedArray[i][k].getDoseInt();
                            manAgeAmountOfSubgroup[i] += manAgesGroupedArray[i][k].getAgeAtExp();
                        }
                    }
                manAverDosesExt[i] = manAverDosesExt[i] / manAgesGroupedArray[i].Count;
                manAverDosesInt[i] = manAverDosesInt[i] / manAgesGroupedArray[i].Count;
            }

            for (int i = 0; i < womanAgesGroupedArray.Length; i++)
            {
                for (int n = 0; n < ageGroups.Count; n++)
                    for (int k = 0; k < womanAgesGroupedArray[i].Count; k++)
                    {
                        if (womanAgesGroupedArray[i][0].getAgeAtExp() >= ageLowerBound[n] && womanAgesGroupedArray[i][0].getAgeAtExp() <= ageUpperBound[n])
                        {
                            womanAverDosesExt[i] += womanAgesGroupedArray[i][k].getDose() - womanAgesGroupedArray[i][k].getDoseInt();
                            womanAverDosesInt[i] += womanAgesGroupedArray[i][k].getDoseInt();
                            womanAgeAmountOfSubgroup[i] += womanAgesGroupedArray[i][k].getAgeAtExp();
                        }
                    }
                womanAverDosesExt[i] = womanAverDosesExt[i] / womanAgesGroupedArray[i].Count;
                womanAverDosesInt[i] = womanAverDosesInt[i] / womanAgesGroupedArray[i].Count;
            }

            /*-----Создание массива, в котором хранится число подгрупп, входящих в возрастную группу-----*/
            double[] manValueOfSubgroups = new double[ageGroups.Count];
            double[] womanValueOfSubgroups = new double[ageGroups.Count];
            /*-----Создание массива, в котором хранятся суммы возрастов всех записей в подгруппе-----*/
            double[] manAgeAmountOfGroup = new double[ageGroups.Count];
            double[] womanAgeAmountOfGroup = new double[ageGroups.Count];
            /*-----Создание массива, в котором хранятся суммы количеств записей для каждой подгруппы, входящей в возрастую группу-----*/
            double[] manAmountOfSubgroupCounts = new double[ageGroups.Count];
            double[] womanAmountOfSubgroupCounts = new double[ageGroups.Count];
            /*-----Создание массивов, в которых хранятся суммы средних доз всех подгрупп, входящих в возрастные группы-----*/
            double[] manAmountOfAverExtDoses = new double[ageGroups.Count];
            double[] manAmountOfAverIntDoses = new double[ageGroups.Count];
            double[] womanAmountOfAverExtDoses = new double[ageGroups.Count];
            double[] womanAmountOfAverIntDoses = new double[ageGroups.Count];

            for (int i = 0; i < ageGroups.Count - 1; i++)
            {
                manValueOfSubgroups[i] = 0;
                manAgeAmountOfGroup[i] = 0;
                manAmountOfSubgroupCounts[i] = 0;
                manAmountOfAverExtDoses[i] = 0;
                manAmountOfAverIntDoses[i] = 0;

                womanValueOfSubgroups[i] = 0;
                womanAgeAmountOfGroup[i] = 0;
                womanAmountOfSubgroupCounts[i] = 0;
                womanAmountOfAverExtDoses[i] = 0;
                womanAmountOfAverIntDoses[i] = 0;
            }

            /*-----Создание массивов списков, которые будут использоваться для расчета среднеквадр. отклонения-----*/
            List<double>[] manArrayForDeviationExt = new List<double>[ageGroups.Count];
            List<double>[] manArrayForDeviationInt = new List<double>[ageGroups.Count];
            List<double>[] womanArrayForDeviationExt = new List<double>[ageGroups.Count];
            List<double>[] womanArrayForDeviationInt = new List<double>[ageGroups.Count];

            for (int i = 0; i < ageGroups.Count; i++)
            {
                manArrayForDeviationExt[i] = new List<double>();
                manArrayForDeviationInt[i] = new List<double>();

                womanArrayForDeviationExt[i] = new List<double>();
                womanArrayForDeviationInt[i] = new List<double>();
            }

            for (int i = 0; i < manAgesGroupedArray.Length; i++)
            {
                for (int k = 0; k < ageGroups.Count; k++)
                {
                    if (manAgeAmountOfSubgroup[i] / manAgesGroupedArray[i].Count >= ageLowerBound[k] && manAgeAmountOfSubgroup[i] / manAgesGroupedArray[i].Count <= ageUpperBound[k])
                    {
                        manValueOfSubgroups[k] += 1;
                        manAmountOfAverExtDoses[k] += manAverDosesExt[i];
                        manAmountOfAverIntDoses[k] += manAverDosesInt[i];
                        manAgeAmountOfGroup[k] += manAgeAmountOfSubgroup[i];
                        manAmountOfSubgroupCounts[k] += manAgesGroupedArray[i].Count;
                        manArrayForDeviationExt[k].Add(manAverDosesExt[i]);
                        manArrayForDeviationInt[k].Add(manAverDosesInt[i]);
                    }
                }
            }

            for (int i = 0; i < womanAgesGroupedArray.Length; i++)
            {
                for (int k = 0; k < ageGroups.Count; k++)
                {
                    if (womanAgeAmountOfSubgroup[i] / womanAgesGroupedArray[i].Count >= ageLowerBound[k] && womanAgeAmountOfSubgroup[i] / womanAgesGroupedArray[i].Count <= ageUpperBound[k])
                    {
                        womanValueOfSubgroups[k] += 1;
                        womanAmountOfAverExtDoses[k] += womanAverDosesExt[i];
                        womanAmountOfAverIntDoses[k] += womanAverDosesInt[i];
                        womanAgeAmountOfGroup[k] += womanAgeAmountOfSubgroup[i];
                        womanAmountOfSubgroupCounts[k] += womanAgesGroupedArray[i].Count;
                        womanArrayForDeviationExt[k].Add(womanAverDosesExt[i]);
                        womanArrayForDeviationInt[k].Add(womanAverDosesInt[i]);
                    }
                }
            }

            /*-----Создание массивов, в которых хранятся среднеквадратические погрешности-----*/
            double[] manDeviationExt = new double[ageGroups.Count];
            double[] manDeviationInt = new double[ageGroups.Count];
            double[] womanDeviationExt = new double[ageGroups.Count];
            double[] womanDeviationInt = new double[ageGroups.Count];
            for (int i = 0; i < ageGroups.Count - 1; i++)
            {
                for (int k = 0; k < manArrayForDeviationExt[i].Count; k++)
                {
                    manDeviationExt[i] += Math.Pow((manArrayForDeviationExt[i][k] - manArrayForDeviationExt[i].Average()), 2);
                }
                for (int k = 0; k < manArrayForDeviationInt[i].Count; k++)
                {
                    manDeviationInt[i] += Math.Pow((manArrayForDeviationInt[i][k] - manArrayForDeviationInt[i].Average()), 2);
                }

                manDeviationExt[i] = Math.Sqrt(manDeviationExt[i] / manArrayForDeviationExt[i].Count);
                manDeviationInt[i] = Math.Sqrt(manDeviationInt[i] / manArrayForDeviationInt[i].Count);


                for (int k = 0; k < womanArrayForDeviationExt[i].Count; k++)
                {
                    womanDeviationExt[i] += Math.Pow((womanArrayForDeviationExt[i][k] - womanArrayForDeviationExt[i].Average()), 2);
                }
                for (int k = 0; k < womanArrayForDeviationInt[i].Count; k++)
                {
                    womanDeviationInt[i] += Math.Pow((womanArrayForDeviationInt[i][k] - womanArrayForDeviationInt[i].Average()), 2);
                }

                womanDeviationExt[i] = Math.Sqrt(womanDeviationExt[i] / womanArrayForDeviationExt[i].Count);
                womanDeviationInt[i] = Math.Sqrt(womanDeviationInt[i] / womanArrayForDeviationInt[i].Count);
            }

            /*-----Создание массивов, в которых хрянятся ОРПО для каждой половозрастной группы-----*/
            double[] manOrpoExt = new double[ageGroups.Count];
            double[] manOrpoInt = new double[ageGroups.Count];
            double[] womanOrpoExt = new double[ageGroups.Count];
            double[] womanOrpoInt = new double[ageGroups.Count];

            /*-----Создание массивов, в которых хрянятся ОРПО-95% для каждой половозрастной группы-----*/
            double[] manOrpoExt_95 = new double[ageGroups.Count];
            double[] manOrpoInt_95 = new double[ageGroups.Count];
            double[] womanOrpoExt_95 = new double[ageGroups.Count];
            double[] womanOrpoInt_95 = new double[ageGroups.Count];

            /*
             * 
             * LAR считается от Зв, а у нас доза в мЗв.
             * Передалать надо!
             * После этого начинаем программировать ИБПО
             * 
             */
            for (int k = 0; k < manAgesGroupedArray.Length; k++)
                for (int i = 0; i < ageGroups.Count; i++)
                {
                    manOrpoExt[i] = getOrpo(getManExtLar(manAgeAmountOfGroup[i] / manAmountOfSubgroupCounts[i]), manAmountOfAverExtDoses[i] / manValueOfSubgroups[i]);
                    manOrpoInt[i] = getOrpo(getManIntLar(manAgeAmountOfGroup[i] / manAmountOfSubgroupCounts[i]), manAmountOfAverIntDoses[i] / manValueOfSubgroups[i]);
                    manOrpoExt_95[i] = getOrpo_95(getManExtLar(manAgeAmountOfGroup[i] / manAmountOfSubgroupCounts[i]), manAmountOfAverExtDoses[i] / manValueOfSubgroups[i], manDeviationExt[i]);
                    manOrpoInt_95[i] = getOrpo_95(getManIntLar(manAgeAmountOfGroup[i] / manAmountOfSubgroupCounts[i]), manAmountOfAverIntDoses[i] / manValueOfSubgroups[i], manDeviationInt[i]);
                    womanOrpoExt[i] = getOrpo(getWomanExtLar(womanAgeAmountOfGroup[i] / womanAmountOfSubgroupCounts[i]), womanAmountOfAverExtDoses[i] / womanValueOfSubgroups[i]);
                    womanOrpoInt[i] = getOrpo(getWomanIntLar(womanAgeAmountOfGroup[i] / womanAmountOfSubgroupCounts[i]), womanAmountOfAverIntDoses[i] / womanValueOfSubgroups[i]);
                    womanOrpoExt_95[i] = getOrpo_95(getWomanExtLar(womanAgeAmountOfGroup[i] / womanAmountOfSubgroupCounts[i]), womanAmountOfAverExtDoses[i] / womanValueOfSubgroups[i], womanDeviationExt[i]);
                    womanOrpoInt_95[i] = getOrpo_95(getWomanIntLar(womanAgeAmountOfGroup[i] / womanAmountOfSubgroupCounts[i]), womanAmountOfAverIntDoses[i] / womanValueOfSubgroups[i], womanDeviationInt[i]);
                }

            ///*-----Вывод в Excel-файл-----*/
            ///*-----Инициализация Excel-файла-----*/
            //Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;
            //excelApp.DisplayAlerts = true;
            //excelApp.StandardFont = "Times-New-Roman";
            //excelApp.StandardFontSize = 12;

            ///*-----Создание рабочей книги с 4 страницами, в которые будет выводиться информация-----*/
            //excelApp.Workbooks.Add(Type.Missing);
            //Excel.Workbook excelWorkbook = excelApp.Workbooks[1];
            //excelApp.SheetsInNewWorkbook = 4;
            //Excel.Worksheet excelWorksheet = null;
            //Excel.Range excelCells = null;

            ///*-----Вывод в столбцы-----*/
            //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
            //excelWorksheet.Name = "Мужчины, ОРПО внеш.";

            ///*-----Описываем ячейку А1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("A1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "Возрастные группы";

            ///*-----Описываем ячейку B1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("B1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО";

            ///*-----Описываем ячейку C1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("C1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО_95";

            //for (int i = 2; i <= ageGroups.Count + 1; i++)
            //{
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
            //    excelCells.Value2 = ageGroups[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
            //    excelCells.Value2 = manOrpoExt[i - 2];
            //    //excelCells.Value2 = 1000 * ((womanAmountOfAverExtDoses[i - 2] / womanValueOfSubgroups[i - 2]) + 1.96 * womanDeviationExt[i - 2]); //на всякий случай выводы для ср эф доз и верхних границ 95% инт
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
            //    excelCells.Value2 = manOrpoExt_95[i - 2];
            //    //excelCells.Value2 = excelCells.Value2 = 1000 * ((womanAmountOfAverIntDoses[i - 2] / womanValueOfSubgroups[i - 2]) + 1.96 * womanDeviationInt[i - 2]);
            //    excelCells.Borders.ColorIndex = 1;

            //}

            //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(2);
            //excelWorksheet.Name = "Мужчины, ОРПО внут.";

            ///*-----Описываем ячейку А1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("A1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "Возрастные группы";

            ///*-----Описываем ячейку B1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("B1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО";

            ///*-----Описываем ячейку C1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("C1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО_95";

            //for (int i = 2; i <= ageGroups.Count + 1; i++)
            //{
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
            //    excelCells.Value2 = ageGroups[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
            //    excelCells.Value2 = manOrpoInt[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
            //    excelCells.Value2 = manOrpoInt_95[i - 2];
            //    excelCells.Borders.ColorIndex = 1;

            //}

            //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(3);
            //excelWorksheet.Name = "Женщины, ОРПО внеш.";

            ///*-----Описываем ячейку А1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("A1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "Возрастные группы";

            ///*-----Описываем ячейку B1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("B1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО";

            ///*-----Описываем ячейку C1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("C1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО_95";

            //for (int i = 2; i <= ageGroups.Count + 1; i++)
            //{
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
            //    excelCells.Value2 = ageGroups[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
            //    excelCells.Value2 = womanOrpoExt[i - 2];  
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
            //    excelCells.Value2 = womanOrpoExt_95[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //}

            //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(4);
            //excelWorksheet.Name = "Женщины, ОРПО внут.";

            ///*-----Описываем ячейку А1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("A1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "Возрастные группы";

            ///*-----Описываем ячейку B1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("B1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО";

            ///*-----Описываем ячейку C1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("C1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "ОРПО_95";

            //for (int i = 2; i <= ageGroups.Count + 1; i++)
            //{
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
            //    excelCells.Value2 = ageGroups[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
            //    excelCells.Value2 = womanOrpoInt[i - 2];    
            //    excelCells.Borders.ColorIndex = 1;
            //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
            //    excelCells.Value2 = womanOrpoInt_95[i - 2];
            //    excelCells.Borders.ColorIndex = 1;
            //}


            //testTextBox.Text = (manAmountOfAverExtDoses[Convert.ToInt32(textBox1.Text)] / manValueOfSubgroups[Convert.ToInt32(textBox1.Text)]).ToString();//getOrpo(getManExtLar(manAgeAmountOfGroup[Convert.ToInt32(textBox1.Text)] / manAmountOfSubgroupCounts[Convert.ToInt32(textBox1.Text)]), (manAmountOfAverExtDoses[Convert.ToInt32(textBox1.Text)] / manValueOfSubgroups[Convert.ToInt32(textBox1.Text)])).ToString();
            //resultTextBox.Text = (manAmountOfAverIntDoses[Convert.ToInt32(textBox1.Text)] / manValueOfSubgroups[Convert.ToInt32(textBox1.Text)]).ToString();
            testTextBox.Text = "ОРПО! " + (dbManAges.Count + dbWomanAges.Count).ToString();
            resultTextBox.Text = "ОРПО! " + dbRecords.Count;
        }

        private void getIbpoButton_Click(object sender, EventArgs e)
        {
            String dbPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\dbTvel.mdb";
            String connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbPath;
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            /*-----Из таблицы Final в эту таблицу считываются поля, указанные в запросе; Выборка для МСК (shop = r3)-----*/
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Dose], [DoseInt], [Year], [Gender], [BirthYear], [AgeAtExp] FROM [Final] WHERE [Shop]='r3'", connectionString);//Выбор нужных столбцов из нужной таблицы
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet, "Final");
            DataTable table = dataSet.Tables[0];

            /*-----Заполнение списка ключей возрастных групп-----*/
            List<String> ageGroups = new List<string>();
            ageGroups.Add("18-24");
            ageGroups.Add("25-29");
            ageGroups.Add("30-34");
            ageGroups.Add("35-39");
            ageGroups.Add("40-44");
            ageGroups.Add("45-49");
            ageGroups.Add("50-54");
            ageGroups.Add("55-59");
            ageGroups.Add("60-64");
            ageGroups.Add("65-69");
            ageGroups.Add("70+");

            /*-----Список, в котором хранятся нижние границы возрастов для возрастных групп-----*/
            List<int> ageLowerBound = new List<int>();
            ageLowerBound.Add(18);
            ageLowerBound.Add(25);
            ageLowerBound.Add(30);
            ageLowerBound.Add(35);
            ageLowerBound.Add(40);
            ageLowerBound.Add(45);
            ageLowerBound.Add(50);
            ageLowerBound.Add(55);
            ageLowerBound.Add(60);
            ageLowerBound.Add(65);
            ageLowerBound.Add(70);

            /*-----Список, в котором хранятся верхние границы возрастов для возрастных групп-----*/
            List<int> ageUpperBound = new List<int>();
            ageUpperBound.Add(24);
            ageUpperBound.Add(29);
            ageUpperBound.Add(34);
            ageUpperBound.Add(39);
            ageUpperBound.Add(44);
            ageUpperBound.Add(49);
            ageUpperBound.Add(54);
            ageUpperBound.Add(59);
            ageUpperBound.Add(64);
            ageUpperBound.Add(69);
            ageUpperBound.Add(100);

            /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
            List<dbObject> dbRecord = new List<dbObject>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                dbRecord.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
            }

            /*-----Список, в котором хранится пол-----*/
            List<byte> dbSex = new List<byte>();
            for (int i = 0; i < dbRecord.Count; i++)
                dbSex.Add(dbRecord[i].getSex());

            /*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
            byte sexMale = dbSex.Min();
            byte sexFemale = dbSex.Max();

            /*-----Уникальные ID мужчин-----*/
            List<int> manUniqueIdList = new List<int>();
            for (int i = 0; i < dbRecord.Count; i++)
            {
                if (dbRecord[i].getSex() == sexMale && dbRecord[i].getYear() == 2012)
                //if (dbRecord[i].getSex() == sexMale)
                {
                    manUniqueIdList.Add(dbRecord[i].getId());
                }
            }
            manUniqueIdList = manUniqueIdList.Distinct().ToList();

            /*-----Уникальные ID женщин-----*/
            List<int> womanUniqueIdList = new List<int>();
            for (int i = 0; i < dbRecord.Count; i++)
            {
                if (dbRecord[i].getSex() == sexFemale && dbRecord[i].getYear() == 2012)
                //if (dbRecord[i].getSex() == sexFemale)
                {
                    womanUniqueIdList.Add(dbRecord[i].getId());
                }
            }
            womanUniqueIdList = womanUniqueIdList.Distinct().ToList();

            /*-----Разделение записей БД на мужские и женские-----*/
            List<dbObject> manList = new List<dbObject>();
            for (int i = 0; i < dbRecord.Count; i++)
                if (dbRecord[i].getSex() == sexMale)
                    manList.Add(dbRecord[i]);

            List<dbObject> womanList = new List<dbObject>();
            for (int i = 0; i < dbRecord.Count; i++)
                if (dbRecord[i].getSex() == sexFemale)
                    womanList.Add(dbRecord[i]);

            /*
             * -----
             * Создания массива списков, где каждый элемент
             * массива - это список объектов, id которых
             * совпадают с уникальными id; например, если уникальный id = 1,
             * то в элемент массива списков записываются все объекты с id = 1.
             * -----
             */
            List<dbObject>[] manIdRecordsArray = new List<dbObject>[manUniqueIdList.Count];
            for (int i = 0; i < manIdRecordsArray.Length; i++)
                manIdRecordsArray[i] = new List<dbObject>();

            for (int i = 0; i < manIdRecordsArray.Length; i++)
                for (int k = 0; k < manList.Count; k++)
                {
                    if (Equals(manUniqueIdList[i], manList[k].getId()))
                    {
                        manIdRecordsArray[i].Add(manList[k]);
                    }
                }

            /*-----Создание аналогичного массива списков для женщин-----*/
            List<dbObject>[] womanIdRecordsArray = new List<dbObject>[womanUniqueIdList.Count];
            for (int i = 0; i < womanIdRecordsArray.Length; i++)
                womanIdRecordsArray[i] = new List<dbObject>();

            for (int i = 0; i < womanIdRecordsArray.Length; i++)
                for (int k = 0; k < womanList.Count; k++)
                {
                    if (Equals(womanUniqueIdList[i], womanList[k].getId()))
                    {
                        womanIdRecordsArray[i].Add(womanList[k]);
                    }
                }

            /*-----Создание пустого списка дозовых историй мужчин; для каждого уникального ID своя дозовая история (по сути, это ячейки, которые надо заполнить)-----*/
            List<RiskCalculator.DoseHistoryRecord[]> manDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
            for (int i = 0; i < manIdRecordsArray.Length; i++)
            {
                manDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[manIdRecordsArray[i].Count]);
            }
            foreach (RiskCalculator.DoseHistoryRecord[] note in manDoseHistoryList)
            {
                for (int i = 0; i < note.Length; i++)
                    note[i] = new RiskCalculator.DoseHistoryRecord();
            }

            /*-----Задание весовых коэффициентов для тканей (в нашем случае учитывается только влияние на лёгкие)-----*/
            double wLung = 0.12;

            /*-----Заполнение дозовых историй мужчин-----*/
            for (int i = 0; i < manIdRecordsArray.Length; i++)
                for (int k = 0; k < manIdRecordsArray[i].Count; k++)
                {
                    manDoseHistoryList[i][k].AgeAtExposure = manIdRecordsArray[i][k].getAgeAtExp();
                    manDoseHistoryList[i][k].AllSolidDoseInmGy = manIdRecordsArray[i][k].getDose() - manIdRecordsArray[i][k].getDoseInt();
                    manDoseHistoryList[i][k].LeukaemiaDoseInmGy = manIdRecordsArray[i][k].getDose() - manIdRecordsArray[i][k].getDoseInt();
                    manDoseHistoryList[i][k].LungDoseInmGy = manIdRecordsArray[i][k].getDoseInt() / wLung;
                }

            /*-----Создание аналогичного списка дозовых историй для женщин-----*/
            List<RiskCalculator.DoseHistoryRecord[]> womanDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
            for (int i = 0; i < womanIdRecordsArray.Length; i++)
            {
                womanDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[womanIdRecordsArray[i].Count]);
            }
            foreach (RiskCalculator.DoseHistoryRecord[] note in womanDoseHistoryList)
            {
                for (int i = 0; i < note.Length; i++)
                    note[i] = new RiskCalculator.DoseHistoryRecord();
            }

            /*-----Заполнение дозовых историй женщин-----*/
            for (int i = 0; i < womanIdRecordsArray.Length; i++)
                for (int k = 0; k < womanIdRecordsArray[i].Count; k++)
                {
                    womanDoseHistoryList[i][k].AgeAtExposure = womanIdRecordsArray[i][k].getAgeAtExp();
                    womanDoseHistoryList[i][k].AllSolidDoseInmGy = womanIdRecordsArray[i][k].getDose() - womanIdRecordsArray[i][k].getDoseInt();
                    womanDoseHistoryList[i][k].LeukaemiaDoseInmGy = womanIdRecordsArray[i][k].getDose() - womanIdRecordsArray[i][k].getDoseInt();
                    womanDoseHistoryList[i][k].LungDoseInmGy = womanIdRecordsArray[i][k].getDoseInt() / wLung;
                }

            /*-----Вычленение только тех членов персонала, что наблюдались включительно по 2012 год-----*/
            List<double>[] manLarExtArray = new List<double>[ageGroups.Count];
            List<double>[] manLarIntArray = new List<double>[ageGroups.Count];

            /*-----Создание аналогичного массива списков LAR для возрастных групп женщин-----*/
            List<double>[] womanLarExtArray = new List<double>[ageGroups.Count];
            List<double>[] womanLarIntArray = new List<double>[ageGroups.Count];

            /*-----Инициализация всех элементов массивов-----*/
            for (int i = 0; i < ageGroups.Count; i++)
            {
                manLarExtArray[i] = new List<double>();
                manLarIntArray[i] = new List<double>();
                womanLarExtArray[i] = new List<double>();
                womanLarIntArray[i] = new List<double>();
            }
            
            for (int i = 0; i < manIdRecordsArray.Length; i++)
                for (int k = 0; k < ageGroups.Count; k++)
                    if (manIdRecordsArray[i][0].getAgeAtExp() >= ageLowerBound[k] && manIdRecordsArray[i][0].getAgeAtExp() <= ageUpperBound[k])
                    {
                        RiskCalculator.DoseHistoryRecord[] record = manDoseHistoryList[i];
                        RiskCalculatorLib.RiskCalculator calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_MALE, manIdRecordsArray[i][0].getAgeAtExp(), ref record, true);
                        manLarExtArray[k].Add(calculator.getLAR(false, true).AllCancers);
                        manLarIntArray[k].Add(calculator.getLAR(false, true).Lung);
                    }

            for (int i = 0; i < womanIdRecordsArray.Length; i++)
                for (int k = 0; k < ageGroups.Count; k++)
                    if (womanIdRecordsArray[i][0].getAgeAtExp() >= ageLowerBound[k] && womanIdRecordsArray[i][0].getAgeAtExp() <= ageUpperBound[k])
                    {
                        RiskCalculator.DoseHistoryRecord[] record = womanDoseHistoryList[i];
                        RiskCalculatorLib.RiskCalculator calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_FEMALE, womanIdRecordsArray[i][0].getAgeAtExp(), ref record, true);
                        womanLarExtArray[k].Add(calculator.getLAR(false, true).AllCancers);
                        womanLarIntArray[k].Add(calculator.getLAR(false, true).Lung);
                    }

            /*
             * Пока это массивы, в которых R хранятся,
             * на самом деле можно хранить тут сразу q
             * 
             */
            /*-----Вычисление знаменателя R, используемого при расчете q, используемого в ИБПО-----*/
            double[] manExtR = new double[ageGroups.Count];
            double[] manIntR = new double[ageGroups.Count];
            double[] womanExtR = new double[ageGroups.Count];
            double[] womanIntR = new double[ageGroups.Count];
            for (int i = 0; i < ageGroups.Count; i++)
            {
                //manExtR[i] = manLarExtArray[i].Sum() / manLarExtArray[i].Count;
                //manIntR[i] = manLarIntArray[i].Sum() / manLarIntArray[i].Count;
                //womanExtR[i] = womanLarExtArray[i].Sum() / womanLarExtArray[i].Count;
                //womanIntR[i] = womanLarIntArray[i].Sum() / womanLarIntArray[i].Count;
                manExtR[i] = 1 - ((manLarExtArray[i].Sum() / manLarExtArray[i].Count) / (4.1 * Math.Pow(10, -2)));//Сейчас здесь вычисляется q, по идее
                manIntR[i] = 1 - ((manLarIntArray[i].Sum() / manLarIntArray[i].Count) / (4.1 * Math.Pow(10, -2)));
                womanExtR[i] = 1 - ((womanLarExtArray[i].Sum() / womanLarExtArray[i].Count) / (4.1 * Math.Pow(10, -2)));
                womanIntR[i] = 1 - ((womanLarIntArray[i].Sum() / womanLarIntArray[i].Count) / (4.1 * Math.Pow(10, -2)));
            }

            testTextBox.Text = "ИБПО! " + Math.Pow(Convert.ToInt32(textBox2.Text), -2);//manLarExtArray[Convert.ToInt32(textBox2.Text)][Convert.ToInt32(textBox1.Text)].ToString();//manIdRecordsArray[Convert.ToInt32(textBox2.Text)][Convert.ToInt32(textBox1.Text)].getId();//manIdRecordsArray.Length + " " + test.Count;//shit[1].Count;//count;//womanDoseHistoryList[0][Convert.ToInt32(textBox1.Text)].LungDoseInmGy;//manIdRecordsArray[0][Convert.ToInt32(textBox1.Text)].getDoseInt();//manUniqueIdList.Count;
            //resultTextBox.Text = "ИБПО! " + manLarIntArray[Convert.ToInt32(textBox2.Text)][Convert.ToInt32(textBox1.Text)].ToString();//manIdRecordsArray[Convert.ToInt32(textBox2.Text)][Convert.ToInt32(textBox1.Text)].getYear();//womanIdRecordsArray.Length;//womanDoseHistoryList[0][Convert.ToInt32(textBox1.Text)].AllSolidDoseInmGy;//womanIdRecordsArray[0][Convert.ToInt32(textBox1.Text)].getDoseInt();//womanUniqueIdList.Count;
            //label1.Text = manLarExtArray[Convert.ToInt32(textBox2.Text)].Count.ToString();//manIdRecordsArray[Convert.ToInt32(textBox2.Text)].Count.ToString();

            ///*-----Вывод в Excel-файл-----*/
            ///*-----Инициализация Excel-файла-----*/
            //Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;
            //excelApp.DisplayAlerts = true;
            //excelApp.StandardFont = "Times-New-Roman";
            //excelApp.StandardFontSize = 12;

            ///*-----Создание рабочей книги с 4 страницами, в которые будет выводиться информация-----*/
            //excelApp.Workbooks.Add(Type.Missing);
            //Excel.Workbook excelWorkbook = excelApp.Workbooks[1];
            //excelApp.SheetsInNewWorkbook = 4;
            //Excel.Worksheet excelWorksheet = null;
            //Excel.Range excelCells = null;

            ///*-----Вывод в столбцы-----*/
            //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
            //excelWorksheet.Name = "Чушь";

            ///*-----Описываем ячейку А1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("A1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "id";

            ///*-----Описываем ячейку B1 на странице-----*/
            //excelCells = excelWorksheet.get_Range("B1");
            //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
            //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
            //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
            //excelCells.Value2 = "последние годы наблюдения";

            ////for (int i = 0; i < manIdRecordsArray.Length; i++)
            ////    for (int k = 0; k < manIdRecordsArray[i].Count; k++)
            //        for (int l = 2; l <= test.Count + 1; l++)
            ////            if (manList[l - 2] == manIdRecordsArray[i][k])
            //            {
            //                excelCells = (Excel.Range)excelWorksheet.Cells[l, "A"];
            //                excelCells.Value2 = test[l-2];//shit[0][i - 2];
            //                excelCells.Borders.ColorIndex = 1;
            //                //excelCells = (Excel.Range)excelWorksheet.Cells[l, "B"];
            //                //excelCells.Value2 = manIdRecordsArray[i][k].getYear();//shit1[0][i - 2];
            //                //excelCells.Borders.ColorIndex = 1;
            //            }


        }

        private void testTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void resultTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void testLabel_Click(object sender, EventArgs e)
        {

        }

        private void resultLabel_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
