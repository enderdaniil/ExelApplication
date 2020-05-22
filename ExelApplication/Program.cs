using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Security;
using Microsoft.VisualBasic.CompilerServices;

namespace ExelApplication
{
    class Program
    {
        interface IExcelWorker
        {
            /// <summary>
            /// Загрузка фйайлов для сравнения
            /// </summary>
            /// <param name="file_1"></param>
            /// <param name="file_2"></param>
            void GetFiles(string file_1, string file_2);

            /// <summary>
            /// новые люди в файле 2
            /// </summary>
            /// <returns></returns>
            List<string> NewPeople();

            /// <summary>
            /// отсутствующие люди в файле 2
            /// </summary>
            /// <returns></returns>
            List<string> MissingPeople();
        }

        class ExcelWorker : IExcelWorker
        {
            class Human
            {
                private string Name;
                private string Surename;
                private string Patronymic;

                public Human(string name, string surename, string patronymic)
                {
                    Name = name;
                    Surename = surename;
                    Patronymic = patronymic;
                }

                public Human()
                {

                }

                void SetName(string name)
                {
                    Name = name;
                }

                void SetSurename(string surename)
                {
                    Surename = surename;
                }

                void SetPatronymic(string patronymic)
                {
                    Patronymic = patronymic;
                }

                public string GetFullName()
                {
                    return (Surename + " " + Name + " " + Patronymic);
                }

                public string GetName()
                {
                    return Name;
                }

                public string GetSurename()
                {
                    return Surename;
                }

                public string GetPatronymic()
                {
                    return Patronymic;
                }

                public static bool operator !=(Human firstHuman, Human secondHuman)
                {
                    if ((firstHuman.Name == secondHuman.Name) && (firstHuman.Surename == secondHuman.Surename) && (firstHuman.Patronymic == secondHuman.Patronymic))
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }

                public static bool operator ==(Human firstHuman, Human secondHuman)
                {
                    if ((firstHuman.Name == secondHuman.Name) && (firstHuman.Surename == secondHuman.Surename) && (firstHuman.Patronymic == secondHuman.Patronymic))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            private Microsoft.Office.Interop.Excel.Worksheet workSheet_1;
            private Microsoft.Office.Interop.Excel.Worksheet workSheet_2;

            private Excel.Application excelApp_1;
            private Excel.Application excelApp_2;

            private Excel.Workbook workBook_1;
            private Excel.Workbook workBook_2;

            public void GetFiles(string file_1, string file_2)
            {
                // Получить объект приложения Excel.
                excelApp_1 = new Excel.ApplicationClass();

                // Сделать Excel невидимым (необязательно).
                excelApp_1.Visible = false;

                // Откройте рабочую книгу только для чтения.
                workBook_1 = excelApp_1.Workbooks.Open(
                    file_1,
                    Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                workSheet_1 = (Microsoft.Office.Interop.Excel.Worksheet)workBook_1.Sheets[1];

                excelApp_2 = new Excel.ApplicationClass();

                // Сделать Excel невидимым (необязательно).
                excelApp_1.Visible = false;

                workBook_2 = excelApp_2.Workbooks.Open(
                    file_2,
                    Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                workSheet_2 = (Microsoft.Office.Interop.Excel.Worksheet)workBook_2.Sheets[1];
            }

            public List<string> MissingPeople()
            {
                List<Human> firstSheetPeople = new List<Human>();
                List<Human> secondSheetPeople = new List<Human>();
                int lastRow_1 = workSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                int lastRow_2 = workSheet_2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                int lengthOfFirst = workSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                int lengthOfSecond = workSheet_2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

                int firstSurenameNumber = 0;
                for (int i = 1; i < lengthOfFirst + 1; i++)
                {
                    if (workSheet_1.Cells[1, i].ToString() == "Фамилия")
                    {
                        firstSurenameNumber = i;
                        break;
                    }
                }

                int firstNameNumber = 1;
                for (int i = 1; i < lengthOfFirst + 1; i++)
                {
                    if (workSheet_1.Cells[1, i].ToString() == "Имя")
                    {
                        firstNameNumber = i;
                        break;
                    }
                }

                int firstPatronymicNumber = 2;
                for (int i = 1; i < lengthOfFirst + 1; i++)
                {
                    if (workSheet_1.Cells[1, i].ToString() == "Отчество")
                    {
                        firstPatronymicNumber = i;
                        break;
                    }
                }


                int secondSurenameNumber = 0;
                for (int i = 1; i < lengthOfSecond + 1; i++)
                {
                    if (workSheet_1.Cells[1, i].ToString() == "Фамилия")
                    {
                        secondSurenameNumber = i;
                        break;
                    }
                }

                int secondNameNumber = 1;
                for (int i = 1; i < lengthOfSecond + 1; i++)
                {
                    if (workSheet_1.Cells[1, i].ToString() == "Имя")
                    {
                        secondNameNumber = i;
                        break;
                    }
                }

                int secondPatronymicNumber = 2;
                for (int i = 1; i < lengthOfSecond + 1; i++)
                {
                    if (workSheet_1.Cells[1, i].ToString() == "Отчество")
                    {
                        secondPatronymicNumber = i;
                        break;
                    }
                }

                for (int i = 2; i < lastRow_1 + 1; i++)
                {
                    string surename = workSheet_1.Cells[i, firstSurenameNumber].ToString();
                    string name = workSheet_1.Cells[i, firstNameNumber].ToString();
                    string patronymic = workSheet_1.Cells[i, firstPatronymicNumber].ToString();
                    Human currentHuman = new Human(name, surename, patronymic);
                    firstSheetPeople.Add(currentHuman);
                }

                for (int i = 2; i < lastRow_2 + 1; i++)
                {
                    string surename = workSheet_2.Cells[i, secondSurenameNumber].ToString();
                    string name = workSheet_2.Cells[i, secondNameNumber].ToString();
                    string patronymic = workSheet_2.Cells[i, secondPatronymicNumber].ToString();
                    Human currentHuman = new Human(name, surename, patronymic);
                    secondSheetPeople.Add(currentHuman);
                }

                List<Human> missingPeople = new List<Human>();

                bool exist = false;
                for (int i = 0; i < lastRow_1; i++)
                {
                    exist = false;
                    for (int j = 0; j < lastRow_2; j++)
                    {
                        if (firstSheetPeople[i] == secondSheetPeople[j])
                        {
                            exist = true;
                        }
                    }

                    if (!exist)
                    {
                        missingPeople.Add(firstSheetPeople[i]);
                    }
                }

                int amountOfMissing = missingPeople.Count;

                List<string> returnListOfPeople = new List<string>();

                for (int i = 0; i < amountOfMissing; i++)
                {
                    returnListOfPeople.Add(missingPeople[i].GetFullName());
                }

                return returnListOfPeople;
            }

            public List<string> NewPeople()
            {
                {
                    List<Human> firstSheetPeople = new List<Human>();
                    List<Human> secondSheetPeople = new List<Human>();
                    int lastRow_1 = workSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                    int lastRow_2 = workSheet_2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                    
                    int lengthOfFirst = workSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    int lengthOfSecond = workSheet_2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

                    int firstSurenameNumber = 0;
                    for (int i = 1; i < lengthOfFirst + 1; i++)
                    {
                        if (workSheet_1.Cells[1, i].ToString() == "Фамилия")
                        {
                            firstSurenameNumber = i;
                            break;
                        }
                    }

                    int firstNameNumber = 1;
                    for (int i = 1; i < lengthOfFirst + 1; i++)
                    {
                        if (workSheet_1.Cells[1, i].ToString() == "Имя")
                        {
                            firstNameNumber = i;
                            break;
                        }
                    }

                    int firstPatronymicNumber = 2;
                    for (int i = 1; i < lengthOfFirst + 1; i++)
                    {
                        if (workSheet_1.Cells[1, i].ToString() == "Отчество")
                        {
                            firstPatronymicNumber = i;
                            break;
                        }
                    }

                    
                    int secondSurenameNumber = 0;
                    for (int i = 1; i < lengthOfSecond + 1; i++)
                    {
                        if (workSheet_1.Cells[1, i].ToString() == "Фамилия")
                        {
                            secondSurenameNumber = i;
                            break;
                        }
                    }

                    int secondNameNumber = 1;
                    for (int i = 1; i < lengthOfSecond + 1; i++)
                    {
                        if (workSheet_1.Cells[1, i].ToString() == "Имя")
                        {
                            secondNameNumber = i;
                            break;
                        }
                    }

                    int secondPatronymicNumber = 2;
                    for (int i = 1; i < lengthOfSecond + 1; i++)
                    {
                        if (workSheet_1.Cells[1, i].ToString() == "Отчество")
                        {
                            secondPatronymicNumber = i;
                            break;
                        }
                    }

                    for (int i = 2; i < lastRow_1 + 1; i++)
                    {
                        string surename = workSheet_1.Cells[i, firstSurenameNumber].ToString();
                        string name = workSheet_1.Cells[i, firstNameNumber].ToString();
                        string patronymic = workSheet_1.Cells[i, firstPatronymicNumber].ToString();
                        Human currentHuman = new Human(name, surename, patronymic);
                        firstSheetPeople.Add(currentHuman);
                    }

                    for (int i = 2; i < lastRow_2 + 1; i++)
                    {
                        string surename = workSheet_2.Cells[i, secondSurenameNumber].ToString();
                        string name = workSheet_2.Cells[i, secondNameNumber].ToString();
                        string patronymic = workSheet_2.Cells[i, secondPatronymicNumber].ToString();
                        Human currentHuman = new Human(name, surename, patronymic);
                        secondSheetPeople.Add(currentHuman);
                    }

                    List<Human> missingPeople = new List<Human>();

                    bool exist = false;
                    for (int i = 0; i < lastRow_2; i++)
                    {
                        exist = false;
                        for (int j = 0; j < lastRow_1; j++)
                        {
                            if (firstSheetPeople[i] == secondSheetPeople[j])
                            {
                                exist = true;
                            }
                        }

                        if (!exist)
                        {
                            missingPeople.Add(firstSheetPeople[i]);
                        }
                    }

                    int amountOfMissing = missingPeople.Count;

                    List<string> returnListOfPeople = new List<string>();

                    for (int i = 0; i < amountOfMissing; i++)
                    {
                        returnListOfPeople.Add(missingPeople[i].GetFullName());
                    }

                    return returnListOfPeople;
                }
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }
}
