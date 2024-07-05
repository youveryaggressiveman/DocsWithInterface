using Accessibility;
using Aspose.Cells.Drawing;
using DockClientApp.Command;
using DockClientApp.Core;
using DockClientApp.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Media3D;

namespace DockClientApp.ViewModel
{
    public class MainWindowViewModel : BaseViewModel
    {
        private ObservableCollection<Document> _listOfDocument;

        private bool _isFinished = true;
        private bool _isStarted = false;
        private bool _isReadyToForm = false;

        private string[] _templates;
        private string _createString;
        private string _path;
        private int _count;

        private Task _threadFor1;
        private Task _threadFor2;
        private Task _threadFor3;
        private Task _threadFor4;
        private Task _threadFor5;

        CancellationTokenSource _cancelTokenSource;

        System.Timers.Timer _timer;

        private ExcelWorker _excel;
        private WordWorker _word;

        public bool IsReadyToForm
        {
            get => _isReadyToForm;
            set
            {
                _isReadyToForm = value;
                OnPropertyChanged(nameof(IsReadyToForm));
            }
        }

        public bool IsStarted
        {
            get => _isStarted;
            set
            {
                _isStarted = value;
                OnPropertyChanged(nameof(IsStarted));
            }
        }
        public bool IsFinshed
        {
            get => _isFinished;
            set
            {
                _isFinished = value;
                OnPropertyChanged(nameof(IsFinshed));
            }
        }
        public int Count
        {
            get => _count;
            set
            {
                _count = value;
                OnPropertyChanged(nameof(Count));
            }
        }

        public ObservableCollection<Document> ListOfDocument
        {
            get => _listOfDocument;
            set
            {
                _listOfDocument = value;
                OnPropertyChanged(nameof(ListOfDocument));
            }
        }

        public string CreateString
        {
            get => _createString;
            set
            {
                _createString = value;
                OnPropertyChanged(nameof(CreateString));
            }
        }

        public string Path
        {
            get => _path;
            set
            {
                _path = value;
                OnPropertyChanged(nameof(Path));
            }
        }

        public ICommand LoadFiles { get; }
        public ICommand CreateDoc { get; }
        public ICommand StopDoc { get; }
        public ICommand FormReport { get; }
        public MainWindowViewModel()
        {
            _templates = new string[] { "Assets/Template/TemplateFor2.rtf", "Assets/Template/TemplateFor1.rtf", "Assets/Template/TemplateFor3.rtf", "Assets/Template/TemplateFor4.rtf", "Assets/Template/TemplateFor5.rtf" };

            ListOfDocument = new ObservableCollection<Document>();

            _timer = new System.Timers.Timer(5000);

            _excel = new ExcelWorker();
            _word = new WordWorker(_templates, @"C:\ExpFiles\", new string[] {@"ЕСТ\", @"ИФФ\", @"ФАМИКОН\"});

            CreateString = string.Empty;
            Count = ListOfDocument.Count;

            FormReport = new DelegateCommand(Form);
            StopDoc = new DelegateCommand(Stop);
            CreateDoc = new DelegateCommand(Create);
            LoadFiles = new DelegateCommand(Load);
        }

        private void Form(object obj)
        {
            try
            {
                Filter(ListOfDocument);

                _excel.FormNewDoc(ListOfDocument.ToList(), @"C:\ExpFiles\");

                MessageBox.Show("Таблица успешна сформирована!", "Успех");
            }
            catch
            {
                MessageBox.Show("При формировании таблицы произошла ошибка. Повторите попытку позже", "Ошибка");
            }
        }

        private void Stop(object obj)
        {
            if (ListOfDocument.Count == 0 || _cancelTokenSource == null)
            {
                MessageBox.Show("Для начала работы необходимо загрузить данные", "Ошибка");

                return;
            }

            try
            {
                _timer.Stop();

                _cancelTokenSource.Cancel();
                Thread.Sleep(1000);
                _cancelTokenSource.Dispose();

                try
                {
                    Task.WaitAll(_threadFor1, _threadFor2, _threadFor3, _threadFor4, _threadFor5);
                }
                catch 
                {
                    ProcessKill();
                }
            }
            catch
            {
            }
            finally
            {
                var result = WordWorker.S_CreatedDoc;
                CreateString += $"{result.Count}/{Count} документов создано\n\nСоздание документов было прервано\n\n";

                IsFinshed = true;
                IsStarted = false;
            }
        }
        private List<List<Document>> Filter(ObservableCollection<Document> allDoc)
        {
            List<List<Document>> filtredList = new()
            {
                new List<Document>(),
                new List<Document>(),
                new List<Document>(),
                new List<Document>(),
                new List<Document>()
            };

            foreach (var doc in allDoc)
            {
                if (WordWorker.S_CreatedDoc.Count != 0)
                {
                    foreach (var currentDoc in WordWorker.S_CreatedDoc)
                    {
                        if (Object.Equals(currentDoc, doc))
                        {
                            goto LoopEnd;
                        }
                    }
                }

                
                var oldExperts = doc.Group.Split("; ");
                doc.Group = string.Empty;

                foreach (var expert in oldExperts)
                {
                    var data = expert.Split(" - ");

                    if (!doc.Authors.Contains(data[0]))
                    {
                        doc.Group += $"{expert}; ";
                    }
                }

                var experts = doc.Group.Split("; ");

                if (experts.Length == 2)
                {
                    filtredList[0].Add(doc);
                }
                else if (experts.Length == 3)
                {
                    filtredList[1].Add(doc);
                }
                else if (experts.Length == 4)
                {
                    filtredList[2].Add(doc);
                }
                else if (experts.Length == 5)
                {
                    filtredList[3].Add(doc);
                }
                else if (experts.Length == 6)
                {
                    filtredList[4].Add(doc);
                }
            LoopEnd: continue;

            }

            ListOfDocument = allDoc;

            return filtredList;
        }

        private void ProcessKill()
        {
            Process[] processes = Process.GetProcessesByName("WINWORD");

            foreach (Process process in processes)
            {
                if (string.IsNullOrEmpty(process.MainWindowTitle))
                {
                    process.Kill();
                }
            }

        }

        private void Create(object obj)
        {
            if(ListOfDocument.Count == 0)
            {
                MessageBox.Show("Для начала работы необходимо загрузить данные", "Ошибка");

                return;
            }

            CreateString += "Создание документов началось...\n\n";

            IsFinshed = false;
            IsStarted = true;

            try
            {
                _timer = new System.Timers.Timer(5000);

                var filtredList = Filter(ListOfDocument);

                _cancelTokenSource = new CancellationTokenSource();
                CancellationToken token = _cancelTokenSource.Token;

                _threadFor1 = new Task(() => _word.Proccess(filtredList[0], "TemplateFor1.rtf", 1, token), token);
                _threadFor2 = new Task(() => _word.Proccess(filtredList[1], "TemplateFor2.rtf", 2, token), token);
                _threadFor3 = new Task(() => _word.Proccess(filtredList[2], "TemplateFor3.rtf", 3, token), token);
                _threadFor4 = new Task(() => _word.Proccess(filtredList[3], "TemplateFor4.rtf", 4, token), token);
                _threadFor5 = new Task(() => _word.Proccess(filtredList[4], "TemplateFor5.rtf", 5, token), token);

                _threadFor1.Start();
                _threadFor2.Start();
                _threadFor3.Start();
                _threadFor4.Start();
                _threadFor5.Start();

                _timer.Elapsed += (sender, e) =>
                {
                    App.Current.Dispatcher.Invoke(() =>
                    {
                        var result = WordWorker.S_CreatedDoc;
                        if (result.Count != Count)
                        {
                            CreateString += $"{result.Count}/{Count} документов создано\n";
                        }
                        else
                        {
                            CreateString += $"\nВсе документы были успешно созданы\n";

                            IsFinshed = true;
                            IsStarted = false;

                            ProcessKill();

                            _timer.Stop();

                        }
                    });

                };

                _timer.Start();
            }
            catch
            {

                MessageBox.Show("При создании документа произошла ошибка", "Ошибка");
            }
        }

        private void Load(object obj)
        {
            if (string.IsNullOrEmpty(Path))
            {
                MessageBox.Show("Укажите путь до файлов", "Напоминание");
            }
            try
            {
                if(ListOfDocument.Count != 0)
                {
                    var result = MessageBox.Show("Вы точно хотите обноваить список данных?", "Вопрос", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        WordWorker.S_CreatedDoc = new List<Document>();
                        WordWorker.S_SdditionName = 1;
                    }
                    else
                    {
                        return;
                    }
                }

                ListOfDocument.Clear();

                var listOfDocument = _excel.ReadDataFromExcel(Path);

                listOfDocument.ForEach(ListOfDocument.Add);

                MessageBox.Show("Все данные из файлов успешно получены", "Успех");

                CreateString += $"Новые данные для создания документов получены\n\n";
            }
            catch
            {
                MessageBox.Show("По данному пути не было найденно нужных файлов", "Ошибка");
            }
            finally
            {
                Count = ListOfDocument.Count;

                if(Count != 0)
                {
                    IsReadyToForm = true;
                }
                else
                {
                    IsReadyToForm = false;
                }
            }

        }
    }
}
