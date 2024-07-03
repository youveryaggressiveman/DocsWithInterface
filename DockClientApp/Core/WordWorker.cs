using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Windows;

namespace DockClientApp.Core
{
    public class WordWorker
    {
        public static List<Model.Document> S_CreatedDoc = new();

        public static int S_SdditionName = 1;

        private List<DirectoryInfo> _foldersInfo;
        private List<FileInfo> _filesInfo;
        private DirectoryInfo _direct;

        public WordWorker(string[] filesName, string direct, string[] foldersName)
        {
            _filesInfo = new List<FileInfo>();
            _foldersInfo = new List<DirectoryInfo>();

            if (Directory.Exists(direct))
            {
                _direct = new DirectoryInfo(direct);
            }
            else
            {
                Directory.CreateDirectory(direct);

                _direct = new DirectoryInfo(direct);
            }

            foreach (var folder in foldersName)
            {
                if (Directory.Exists($"{direct}{folder}"))
                {
                    _foldersInfo.Add(new DirectoryInfo($"{direct}{folder}"));
                }
                else
                {
                    Directory.CreateDirectory($"{direct}{folder}");

                    _foldersInfo.Add(new DirectoryInfo($"{direct}{folder}"));
                }
            }

            foreach (var file in filesName)
            {
                if (File.Exists(file))
                {
                    _filesInfo.Add(new FileInfo(file));
                }
                else
                {
                    throw new ArgumentException("File not found");
                }
            }

        }

        public void Proccess(List<Model.Document> documents, string type, int countOfRepeat, CancellationToken token)
        {
            if (documents.Count == 0)
            {
                return;
            }

            foreach (var document in documents)
            {
                if (token.IsCancellationRequested)
                {
                    token.ThrowIfCancellationRequested();
                }

                CreateDocument(type, countOfRepeat, document);
            }

        }

        private void CreateDocument(string fileName, int countOfIteration, Model.Document document)
        {
            Word.Application app = null;

            try
            {
                Dictionary<string, string> listOfData = new()
                {
                    { "<POST>", document.Post},
                    { "<MAIN_FIO>", document.MainFio},
                    { "<GROUP>", document.Group},
                    { "<NAME_OF_PUBLICATION>", document.NameOfPublication },
                    { "<PLACE>", document.Place},
                    { "<AUTHORS>", document.Authors }
                };

                FileInfo currentFile;
                Object file = new();
                Object missing = Type.Missing;

                foreach (FileInfo fil in _filesInfo)
                {
                    if (fil.Name == fileName.Replace("Template/", ""))
                    {
                        currentFile = fil;
                        file = currentFile.FullName;

                        break;
                    }
                }

                app = new Word.Application();

                app.Documents.Open(file);

                foreach (var data in listOfData)
                {
                    if (data.Value.Length > 255)
                    {
                        AddBigData(data.Key, data.Value, app);
                    }
                    else
                    {
                        Word.Find find = app.Selection.Find;
                        find.Text = data.Key;
                        find.Replacement.Text = data.Value;

                        Object wrap = Word.WdFindWrap.wdFindContinue;
                        Object replace = Word.WdReplace.wdReplaceAll;

                        find.Execute(FindText: Type.Missing,
                            MatchCase: false,
                            MatchWholeWord: false,
                            MatchWildcards: false,
                            MatchSoundsLike: missing,
                            MatchAllWordForms: false,
                            Forward: true,
                            Wrap: wrap,
                            Format: false,
                            ReplaceWith: missing, Replace: replace);
                    }
                }

                var expert = document.Group.Split(";");

                int number = 1;

                for (int count = 0; count < countOfIteration; count++)
                {
                    var data = expert[count].Split(" - ");

                    app.ActiveDocument.Content.Find.Execute($"<FIO_EXP{number}>", ReplaceWith: data[0]);
                    app.ActiveDocument.Content.Find.Execute($"<POST_EXP{number}>", ReplaceWith: data[1]);

                    number++;
                }
                string newAuthors = Regex.Replace(document.Authors.Split(";")[0], @"[\\/:*?""<>|+\s]", "");
                string newNameOfDirection = Regex.Replace(document.NameOfDirection, @"[\\/:*?""<>|+\s.]", "");
                string newNameOfPublication = Regex.Replace(document.NameOfPublication, @"[\\/:*?""<>|+.]", "");

                foreach (var folder in _foldersInfo)
                {
                    if (folder.Name.Contains(document.NameOfDirection))
                    {
                        var currentFolder = Directory.CreateDirectory($@"{folder.FullName}/{document.Year}_{newNameOfDirection}_{newAuthors}_{S_SdditionName}");

                        Object newFileName = Path.Combine(currentFolder.FullName, "Заключение экспертной группы.rtf");
                        app.ActiveDocument.SaveAs2(newFileName);
                        app.ActiveDocument.Close();
                        app.Quit();

                        break;
                    }
                }

                S_CreatedDoc.Add(document);
                S_SdditionName++;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                try
                {
                    app?.Quit();
                }
                catch
                {

                }
                
            }
        }

        private void AddBigData(string findText, string replaceWith, Word.Application app)
        {
            while (replaceWith.Length > 255)
            {
                var replacePart = replaceWith.Substring(0, 200);
                replacePart += findText;
                replaceWith = replaceWith.Substring(200);

                app.ActiveDocument.Content.Find.Execute(FindText: findText, ReplaceWith: replacePart, Replace: Word.WdReplace.wdReplaceAll);
            }

            app.ActiveDocument.Content.Find.Execute(FindText: findText, ReplaceWith: replaceWith, Replace: Word.WdReplace.wdReplaceAll);
        }
    }
}
