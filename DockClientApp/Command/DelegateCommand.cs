using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DockClientApp.Command
{
    public class DelegateCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public DelegateCommand(Action<object> execute, Func<object, bool> canExecute) =>
            (_execute, _canExecute) = (execute, canExecute);

        public DelegateCommand(Action<object> execute) : this(execute, null) { }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter)
        {
            if (_canExecute != null)
            {
                return _canExecute(parameter);
            }

            return true;
        }

        public void Execute(object? parameter)
        {
            _execute?.Invoke(parameter);
        }
    }
}
