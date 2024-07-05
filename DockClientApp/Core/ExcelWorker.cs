﻿using Aspose.Cells;
using DockClientApp.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace DockClientApp.Core
{
    public class ExcelWorker
    {
        private List<string> _nameOfExcel;
        private List<string> _nameOfColumn;

        private List<Document> _expList;

        private List<Document> _listOfDocument;

        private List<Document> _estDocumment;
        private List<Document> _famDocuments;
        private List<Document> _iffDocuments;

        public ExcelWorker()
        {
            _nameOfColumn = new List<string>()
            {
                "Название факультета",
                "Год",
                "Направление",
                "Авторы",
                "Название публикации",
                "Дата публикации",
                "Место",
                "Руководитель эксп. компании",
                "Эксперты",
                "Срок полномочий"
            };

            _nameOfExcel = new List<string>()
            {
                "2_Состав экспертных групп_2020-2023",
                "ЕСТ-конференции",
                "ЕСТ-статьи",
                "ЕСТ-студ",
                "ИФФ-конференции",
                "ИФФ-статьи",
                "ИФФ-студ",
                "ФАМИКОН-конференции",
                "ФАМИКОН-статьи",
                "ФАМИКОН-студ"
            };

            _listOfDocument = new();
            _expList = new();
            _estDocumment = new();
            _famDocuments = new();
            _iffDocuments = new();
        }


        public void FormNewDoc(List<Document> docs, string direction)
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            for (int row = 0; row < docs.Count + 1; row++)
            {
                for (int col = 0; col <= 9 ; col++)
                {
                    Cell cell;

                    if (row == 0)
                    {
                        cell = worksheet.Cells[row, col];
                        Style style = cell.GetStyle();
                        style.Pattern = BackgroundType.Solid;
                        style.ForegroundColor = Color.Green;
                        style.Font.Color = Color.White;
                        style.Font.IsBold = true;
                        style.HorizontalAlignment = TextAlignmentType.Center;
                        cell.SetStyle(style);
                        switch (col)
                        {
                            case 0:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 1:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 2:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 3:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 4:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 5:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 6:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 7:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 8:
                                cell.PutValue(_nameOfColumn[col]); break;
                            case 9:
                                cell.PutValue(_nameOfColumn[col]); break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        cell = worksheet.Cells[row, col];

                        var doc = docs[row-1];

                        if(doc.DateOfPublication == null)
                        {
                            if ((doc.Place.ToLower().Contains("янв.") || doc.Place.ToLower().Contains("jan.")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Январь";
                            }
                            else if ((doc.Place.ToLower().Contains("февр") || doc.Place.ToLower().Contains("feb")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Февраль";
                            }
                            else if ((doc.Place.ToLower().Contains("март") || doc.Place.ToLower().Contains("mar")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Март";
                            }
                            else if ((doc.Place.ToLower().Contains("апр") || doc.Place.ToLower().Contains("apr")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Апрель";
                            }
                            else if ((doc.Place.ToLower().Contains("май") || doc.Place.ToLower().Contains("маz") || doc.Place.ToLower().Contains("may")) && !doc.Place.ToLower().Contains("майкоп") && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Май";
                            }
                            else if ((doc.Place.ToLower().Contains("июн") || doc.Place.ToLower().Contains("jun")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Июнь";
                            }
                            else if ((doc.Place.ToLower().Contains("июл") || doc.Place.ToLower().Contains("jul")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Июль";
                            }
                            else if ((doc.Place.ToLower().Contains("авг") || doc.Place.ToLower().Contains("aug")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Август";
                            }
                            else if ((doc.Place.ToLower().Contains("сент") || doc.Place.ToLower().Contains("sep")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Сентябрь";
                            }
                            else if ((doc.Place.ToLower().Contains("окт") || doc.Place.ToLower().Contains("oct")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Октябрь";
                            }
                            else if ((doc.Place.ToLower().Contains("нояб") || doc.Place.ToLower().Contains("nov")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Ноябрь";
                            }
                            else if ((doc.Place.ToLower().Contains("дек") || doc.Place.ToLower().Contains("dec")) && doc.DateOfPublication == null)
                            {
                                doc.DateOfPublication = "Декабрь";
                            }
                        }

                        switch (col)
                        {
                            case 0:
                                cell.PutValue(doc.NameOfDirection); break;
                            case 1:
                                cell.PutValue(doc.Year); break;
                            case 2:
                                cell.PutValue(doc.WorkDirection); break;
                            case 3:
                                cell.PutValue(doc.Authors); break;
                            case 4:
                                cell.PutValue(doc.NameOfPublication); break;
                            case 5:
                                cell.PutValue(doc.DateOfPublication); break;
                            case 6:
                                cell.PutValue(doc.Place); break;
                            case 7:
                                cell.PutValue($"{doc.MainFio} - {doc.Post}"); break;
                            case 8:
                                cell.PutValue(doc.Group); break;
                            case 9:
                                cell.PutValue(doc.Period); break;
                            default:
                                break;
                        }
                    }
                }
            }

            workbook.Save($"{direction}Отформатированные данные.xlsx");
        }

        public List<Document> ReadDataFromExcel(string path)
        {
            _listOfDocument.Clear();
            _expList.Clear();
            _estDocumment.Clear();
            _famDocuments.Clear();
            _iffDocuments.Clear();

            try
            {
                Workbook workbook;

                foreach (string name in _nameOfExcel)
                {
                    workbook = new Workbook($"{path}/{name}.xlsx");
                    WorksheetCollection collection = workbook.Worksheets;

                    for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
                    {
                        Worksheet worksheet = collection[worksheetIndex];

                        List<Document> result = new();

                        if (name.Contains("ЕСТ"))
                        {
                            result = SearchCurrentFile(name.Replace("ЕСТ-", ""), worksheet, "ЕСТ");

                            result.ForEach(_estDocumment.Add);

                            continue;
                        }
                        else if (name.Contains("ФАМИКОН"))
                        {
                            result = SearchCurrentFile(name.Replace("ФАМИКОН-", ""), worksheet, "ФАМИКОН");

                            result.ForEach(_famDocuments.Add);

                            continue;
                        }
                        else if (name.Contains("ИФФ"))
                        {
                            result = SearchCurrentFile(name.Replace("ИФФ-", ""), worksheet, "ИФФ");

                            result.ForEach(_iffDocuments.Add);

                            continue;
                        }
                        else if (name == "2_Состав экспертных групп_2020-2023")
                        {
                            int totalColumns = worksheet.Cells.MaxColumn + 1;

                            for (int col = 0; col < totalColumns; col++)
                            {
                                string cellValue = worksheet.Cells[0, col].StringValue;

                                if (cellValue == "Направление работы")
                                {
                                    for (int row = 2; row < worksheet.Cells.MaxRow + 1; row++)
                                    {
                                        string cellValueOfWorkDirection = worksheet.Cells[row, col].StringValue;
                                        if (cellValueOfWorkDirection != string.Empty)
                                        {
                                            Document document = new();
                                            document.WorkDirection = cellValueOfWorkDirection;
                                            for (int rowOfThisColumn = row; rowOfThisColumn <= worksheet.Cells.MaxDataRow + 1; rowOfThisColumn++)
                                            {
                                                Cell currentCell = worksheet.Cells[rowOfThisColumn, col];
                                                if (cellValueOfWorkDirection == currentCell.StringValue)
                                                {
                                                    for (int colOfThisRow = 0; colOfThisRow < totalColumns; colOfThisRow++)
                                                    {
                                                        switch (colOfThisRow)
                                                        {
                                                            case 1:
                                                                document.Period = worksheet.Cells[rowOfThisColumn, colOfThisRow].StringValue;
                                                                break;
                                                            case 3:
                                                                document.Post = worksheet.Cells[rowOfThisColumn, colOfThisRow].StringValue;
                                                                break;
                                                            case 4:
                                                                document.MainFio = worksheet.Cells[rowOfThisColumn, colOfThisRow].StringValue;
                                                                break;
                                                            default:
                                                                break;
                                                        }
                                                    }

                                                    continue;
                                                }

                                                else if (currentCell.StringValue == string.Empty)
                                                {

                                                    document.Group += $"{worksheet.Cells[rowOfThisColumn, 4].StringValue} - {worksheet.Cells[rowOfThisColumn, 3].StringValue}; ";

                                                    continue;
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }

                                            _expList.Add(document);
                                        }
                                    }
                                }
                            }
                        }
                    }

                }

                foreach (var document in _expList)
                {
                    if (document.WorkDirection.Contains("Инженерно-физический факультет"))
                    {
                        foreach (var iff in _iffDocuments)
                        {
                            if ((document.Period.Contains("2020") && iff.Year == "2021") || (document.Period.Contains("2023") && iff.Year == "2022"))
                            {
                                _listOfDocument.Add(new Document()
                                {
                                    Post = document.Post,
                                    MainFio = document.MainFio,
                                    Group = document.Group,
                                    Period = document.Period,
                                    WorkDirection = document.WorkDirection,
                                    DateOfPublication = iff.DateOfPublication,
                                    NameOfDirection = iff.NameOfDirection,
                                    NameOfPublication = iff.NameOfPublication,
                                    Place = iff.Place,
                                    Authors = iff.Authors,
                                    Year = iff.Year,
                                });
                            }
                        }

                    }
                    else if (document.WorkDirection.Contains("Факультет естествознания"))
                    {
                        foreach (var est in _estDocumment)
                        {
                            if ((document.Period.Contains("2020") && est.Year == "2021") || (document.Period.Contains("2023") && est.Year == "2022"))
                            {
                                _listOfDocument.Add(new Document()
                                {
                                    Post = document.Post,
                                    MainFio = document.MainFio,
                                    Group = document.Group,
                                    Period = document.Period,
                                    WorkDirection = document.WorkDirection,
                                    DateOfPublication = est.DateOfPublication,
                                    NameOfDirection = est.NameOfDirection,
                                    NameOfPublication = est.NameOfPublication,
                                    Place = est.Place,
                                    Authors = est.Authors,
                                    Year = est.Year,
                                });
                            }
                        }
                    }
                    else if (document.WorkDirection.Contains("Факультет математики и компьютерных наук"))
                    {
                        foreach (var fam in _famDocuments)
                        {
                            if ((document.Period.Contains("2020") && fam.Year == "2021") || (document.Period.Contains("2023") && fam.Year == "2022"))
                            {
                                _listOfDocument.Add(new Document()
                                {
                                    Post = document.Post,
                                    MainFio = document.MainFio,
                                    Group = document.Group,
                                    Period = document.Period,
                                    WorkDirection = document.WorkDirection,
                                    DateOfPublication = fam.DateOfPublication,
                                    NameOfDirection = fam.NameOfDirection,
                                    NameOfPublication = fam.NameOfPublication,
                                    Place = fam.Place,
                                    Authors = fam.Authors,
                                    Year = fam.Year,
                                });
                            }

                        }
                    }
                    else if (document.WorkDirection.Contains("НОК «Институт живых систем и инженерии здоровья»"))
                    {
                        foreach (var est in _estDocumment)
                        {
                            if (document.Period.Contains("2023") && est.Year == "2023")
                            {
                                var exp = document.Group.Split("; ");
                                string newGroup = string.Empty;

                                for (int i = 0; i < 3; i++)
                                {
                                    newGroup += $"{exp[i]}; ";
                                }

                                _listOfDocument.Add(new Document()
                                {
                                    Post = document.Post,
                                    MainFio = document.MainFio,
                                    Group = newGroup,
                                    Period = document.Period,
                                    WorkDirection = document.WorkDirection,
                                    DateOfPublication = est.DateOfPublication,
                                    NameOfDirection = est.NameOfDirection,
                                    NameOfPublication = est.NameOfPublication,
                                    Place = est.Place,
                                    Authors = est.Authors,
                                    Year = est.Year,
                                });
                            }
                        }
                    }
                    else if (document.WorkDirection.Contains("НОК «Институт точных наук и цифровых технологий»"))
                    {
                        foreach (var fam in _famDocuments)
                        {
                            if (document.Period.Contains("2023") && fam.Year == "2023")
                            {
                                var exp = document.Group.Split("; ");
                                string newGroup = string.Empty;

                                for (int i = 0; i < 5; i++)
                                {
                                    newGroup += $"{exp[i]}; ";
                                }

                                _listOfDocument.Add(new Document()
                                {
                                    Post = document.Post,
                                    MainFio = document.MainFio,
                                    Group = newGroup,
                                    Period = document.Period,
                                    WorkDirection = document.WorkDirection,
                                    DateOfPublication = fam.DateOfPublication,
                                    NameOfDirection = fam.NameOfDirection,
                                    NameOfPublication = fam.NameOfPublication,
                                    Place = fam.Place,
                                    Authors = fam.Authors,
                                    Year = fam.Year,
                                });
                            }
                        }
                        foreach (var iff in _iffDocuments)
                        {
                            if (document.Period.Contains("2023") && iff.Year == "2023")
                            {
                                var exp = document.Group.Split("; ");
                                string newGroup = string.Empty;

                                for (int i = 5; i < 10; i++)
                                {
                                    newGroup += $"{exp[i]}; ";
                                }

                                if (string.IsNullOrEmpty(iff.NameOfPublication)) { }

                                _listOfDocument.Add(new Document()
                                {
                                    Post = document.Post,
                                    MainFio = document.MainFio,
                                    Group = newGroup,
                                    Period = document.Period,
                                    WorkDirection = document.WorkDirection,
                                    DateOfPublication = iff.DateOfPublication,
                                    NameOfDirection = iff.NameOfDirection,
                                    NameOfPublication = iff.NameOfPublication,
                                    Place = iff.Place,
                                    Authors = iff.Authors,
                                    Year = iff.Year,
                                });
                            }
                        }
                    }
                }


                return _listOfDocument;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private List<Document> SearchCurrentFile(string additionalName, Worksheet worksheet, string name)
        {
            Dictionary<int, string> columnList;

            List<Document> middleList = new();
            Document document;

            int totalColumns = worksheet.Cells.MaxColumn + 1;
            int totalRows = worksheet.Cells.MaxRow + 1;

            switch (additionalName)
            {
                case "конференции":
                    columnList = new();

                    for (int col = 0; col < totalColumns; col++)
                    {

                        string cellValue = worksheet.Cells[0, col].StringValue;

                        if (cellValue == "Название доклада")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Ф.И.О.")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Дата начала проведения")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Дата окончания проведения")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Город проведения")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else
                            continue;
                    }

                    for (int row = 1; row < totalRows; row++)
                    {
                        document = new();

                        document.Year = worksheet.Name;

                        foreach (var col in columnList)
                        {
                            switch (col.Value)
                            {
                                case "Название доклада":
                                    document.NameOfPublication = worksheet.Cells[row, col.Key].StringValue;
                                    break;
                                case "Ф.И.О.":
                                    document.Authors = worksheet.Cells[row, col.Key].StringValue;
                                    break;
                                case "Дата начала проведения":
                                    document.DateOfPublication += $"{worksheet.Cells[row, col.Key].StringValue} - ";
                                    break;
                                case "Дата окончания проведения":
                                    document.DateOfPublication += worksheet.Cells[row, col.Key].StringValue;
                                    break;
                                case "Город проведения":
                                    document.Place = worksheet.Cells[row, col.Key].StringValue;
                                    break;
                                default:
                                    break;
                            }
                        }

                        document.NameOfDirection = name;

                        if (!string.IsNullOrEmpty(document.Authors.Split(";")[0]))
                        {
                            middleList.Add(document);
                        }
                    }

                    break;
                case "статьи":
                    columnList = new();

                    for (int col = 0; col < totalColumns; col++)
                    {
                        string cellValue = worksheet.Cells[0, col].StringValue;

                        if (cellValue == "Название публикации")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Ф.И.О.")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Соавторы")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Название, серия и № журнала")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Место расположения, страницы")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else
                            continue;
                    }

                    for (int row = 1; row < totalRows; row++)
                    {
                        document = new();

                        document.Year = worksheet.Name;

                        foreach (var col in columnList)
                        {
                            switch (col.Value)
                            {
                                case "Название публикации":
                                    document.NameOfPublication = worksheet.Cells[row, col.Key].StringValue;
                                    break;
                                case "Ф.И.О.":
                                    document.Authors += $"{worksheet.Cells[row, col.Key].StringValue}; ";
                                    break;
                                case "Соавторы":
                                    document.Authors += $"Соавторы: {worksheet.Cells[row, col.Key].StringValue};";
                                    break;
                                case "Название, серия и № журнала":
                                    document.Place += $"{worksheet.Cells[row, col.Key].StringValue} - ";
                                    break;
                                case "Место расположения, страницы":
                                    document.Place += $"стр. {worksheet.Cells[row, col.Key].StringValue}";
                                    break;
                                default:
                                    break;
                            }
                        }

                        document.NameOfDirection = name;

                        if (!string.IsNullOrEmpty(document.Authors.Split(";")[0]))
                        {
                            middleList.Add(document);
                        }
                    }

                    break;
                case "студ":
                    columnList = new();

                    for (int col = 0; col < totalColumns; col++)
                    {
                        string cellValue = worksheet.Cells[0, col].StringValue;

                        if (cellValue == "Название статьи")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Ф.И.О.")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "ФИО студента")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Название журнала, сборника")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else if (cellValue == "Место расположение , страницы")
                        {
                            columnList.Add(col, cellValue);
                        }
                        else
                            continue;
                    }

                    for (int row = 1; row < totalRows; row++)
                    {
                        document = new();

                        document.Year = worksheet.Name;

                        foreach (var col in columnList)
                        {
                            switch (col.Value)
                            {
                                case "Название статьи":
                                    document.NameOfPublication = worksheet.Cells[row, col.Key].StringValue;
                                    break;
                                case "Ф.И.О.":
                                    document.Authors += $"{worksheet.Cells[row, col.Key].StringValue}; ";
                                    break;
                                case "ФИО студента":
                                    document.Authors += $"Студент: {worksheet.Cells[row, col.Key].StringValue}; ";
                                    break;
                                case "Название журнала, сборника":
                                    document.Place += $"{worksheet.Cells[row, col.Key].StringValue} - ";
                                    break;
                                case "Место расположение , страницы":
                                    document.Place += $"стр. {worksheet.Cells[row, col.Key].StringValue}";
                                    break;
                                default:
                                    break;
                            }
                        }

                        document.NameOfDirection = name;

                        if (!string.IsNullOrEmpty(document.Authors.Split(";")[0]))
                        {
                            middleList.Add(document);
                        }
                    }

                    break;
                default:
                    break;
            }

            return middleList;
        }
    }
}
