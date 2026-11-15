using System;
using Microsoft.Office.Interop.Word;

namespace VFormat.WordEngine.Interop;

public class ScreenFreezeScope : IDisposable
{
	private readonly Application _app;

	private readonly bool _originalScreenUpdating;

	private readonly WdAlertLevel _originalDisplayAlerts;

	private bool _disposed;

	public ScreenFreezeScope(Application app)
	{
		_app = app;
		if (_app != null)
		{
			try
			{
				_originalScreenUpdating = _app.ScreenUpdating;
				_originalDisplayAlerts = _app.DisplayAlerts;
				_app.ScreenUpdating = false;
				_app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
			}
			catch
			{
			}
		}
	}

	public void Dispose()
	{
		if (_disposed)
		{
			return;
		}
		_disposed = true;
		if (_app != null)
		{
			try
			{
				_app.ScreenUpdating = _originalScreenUpdating;
			}
			catch
			{
			}
			try
			{
				_app.DisplayAlerts = _originalDisplayAlerts;
			}
			catch
			{
			}
		}
	}
}
