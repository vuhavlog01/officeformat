using System;
using Microsoft.Office.Interop.Word;

namespace VFormat.WordEngine.Interop;

public class SafeUndoScope : IDisposable
{
	private readonly Application _app;

	private UndoRecord _undoRecord;

	private bool _isStarted;

	private bool _disposed;

	public SafeUndoScope(Application app, string actionName = "OfficeFormat Chuẩn Hóa Thể Thức")
	{
		_app = app;
		if (_app == null)
		{
			return;
		}
		try
		{
			_undoRecord = _app.UndoRecord;
			if (_undoRecord != null && !_undoRecord.IsRecordingCustomRecord)
			{
				_undoRecord.StartCustomRecord(actionName);
				_isStarted = true;
			}
		}
		catch
		{
		}
	}

	public void Dispose()
	{
		if (_disposed)
		{
			return;
		}
		_disposed = true;
		if (!_isStarted || _undoRecord == null)
		{
			_undoRecord = null;
			return;
		}
		try
		{
			if (_undoRecord.IsRecordingCustomRecord)
			{
				_undoRecord.EndCustomRecord();
			}
		}
		catch
		{
		}
		finally
		{
			_undoRecord = null;
		}
	}
}
