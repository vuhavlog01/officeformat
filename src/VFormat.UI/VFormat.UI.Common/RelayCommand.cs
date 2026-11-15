using System;
using System.Windows.Input;

namespace VFormat.UI.Common;

public class RelayCommand : ICommand
{
	private readonly Action<object> _execute;

	private readonly Predicate<object> _canExecute;

	public event EventHandler CanExecuteChanged
	{
		add
		{
			CommandManager.RequerySuggested += value;
		}
		remove
		{
			CommandManager.RequerySuggested -= value;
		}
	}

	public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
	{
		if (execute == null)
		{
			throw new ArgumentNullException("execute");
		}
		_execute = execute;
		_canExecute = canExecute;
	}

	public RelayCommand(Action execute, Func<bool> canExecute = null)
	{
		if (execute == null)
		{
			throw new ArgumentNullException("execute");
		}
		_execute = (object p) =>
		{
			execute();
		};
		_canExecute = ((canExecute != null) ? ((Predicate<object>)((object p) => canExecute())) : null);
	}

	public bool CanExecute(object parameter)
	{
		if (_canExecute != null)
		{
			return _canExecute(parameter);
		}
		return true;
	}

	public void Execute(object parameter)
	{
		_execute(parameter);
	}

	public void RaiseCanExecuteChanged()
	{
		CommandManager.InvalidateRequerySuggested();
	}
}
