using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace VFormat.UI.Common;

public sealed class ProgressDialogHost : IDisposable
{
	private struct NativeRect
	{
		public int Left;

		public int Top;

		public int Right;

		public int Bottom;
	}

	private readonly ManualResetEvent _ready = new ManualResetEvent(initialState: false);

	private Thread _thread;

	private Dispatcher _dispatcher;

	private ProgressDialog _dialog;

	private volatile bool _cancelRequested;

	private volatile bool _disposed;

	private readonly object _sync = new object();

	private int _pendingPercent;

	private string _pendingMessage;

	private bool _updateQueued;

	public bool IsCancelRequested => _cancelRequested;

	private ProgressDialogHost()
	{
	}

	public static ProgressDialogHost Start(string title, string statusMessage, IntPtr ownerHwnd, bool allowCancel)
	{
		ProgressDialogHost host = new ProgressDialogHost();
		host._thread = new Thread(() =>
		{
			host.Run(title, statusMessage, ownerHwnd, allowCancel);
		});
		host._thread.SetApartmentState(ApartmentState.STA);
		host._thread.IsBackground = true;
		host._thread.Name = "OfficeFormat Progress";
		host._thread.Start();
		host._ready.WaitOne(5000);
		return host;
	}

	private void Run(string title, string statusMessage, IntPtr ownerHwnd, bool allowCancel)
	{
		try
		{
			_dispatcher = Dispatcher.CurrentDispatcher;
			_dialog = new ProgressDialog(title, statusMessage);
			_dialog.EnableCancel(allowCancel);
			_dialog.CancelRequested += () =>
			{
				_cancelRequested = true;
			};
			CenterOver(_dialog, ownerHwnd);
			_dialog.Closed += (object s, EventArgs e) =>
			{
				try
				{
					_dispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
				}
				catch
				{
				}
			};
			_dialog.Show();
		}
		catch
		{
			_dialog = null;
		}
		finally
		{
			_ready.Set();
		}
		if (_dialog != null)
		{
			try
			{
				Dispatcher.Run();
			}
			catch
			{
			}
		}
	}

	public void Update(int percent, string message)
	{
		if (_disposed || _dispatcher == null || _dialog == null)
		{
			return;
		}
		lock (_sync)
		{
			_pendingPercent = percent;
			if (!string.IsNullOrEmpty(message))
			{
				_pendingMessage = message;
			}
			if (_updateQueued)
			{
				return;
			}
			_updateQueued = true;
		}
		try
		{
			_dispatcher.BeginInvoke(DispatcherPriority.Normal, (Action)(() =>
			{
				int pendingPercent;
				string pendingMessage;
				lock (_sync)
				{
					pendingPercent = _pendingPercent;
					pendingMessage = _pendingMessage;
					_updateQueued = false;
				}
				try
				{
					if (_dialog != null && _dialog.IsVisible)
					{
						_dialog.UpdateProgress(pendingPercent, pendingMessage);
					}
				}
				catch
				{
				}
			}));
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
		try
		{
			if (_dispatcher != null && _dialog != null)
			{
				_dispatcher.BeginInvoke((Action)(() =>
				{
					try
					{
						_dialog.Close();
					}
					catch
					{
					}
				}));
			}
		}
		catch
		{
		}
		try
		{
			if (_thread != null)
			{
				_thread.Join(3000);
			}
		}
		catch
		{
		}
	}

	[DllImport("user32.dll")]
	private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect rect);

	private static void CenterOver(Window dialog, IntPtr ownerHwnd)
	{
		try
		{
			if (ownerHwnd == IntPtr.Zero || !GetWindowRect(ownerHwnd, out var r) || r.Right <= r.Left)
			{
				return;
			}
			dialog.SourceInitialized += (object s, EventArgs e) =>
			{
				try
				{
					PresentationSource presentationSource = PresentationSource.FromVisual(dialog);
					double num = 1.0;
					double num2 = 1.0;
					if (presentationSource != null && presentationSource.CompositionTarget != null)
					{
						num = presentationSource.CompositionTarget.TransformFromDevice.M11;
						num2 = presentationSource.CompositionTarget.TransformFromDevice.M22;
					}
					double num3 = ((dialog.ActualWidth > 0.0) ? dialog.ActualWidth : dialog.Width);
					double num4 = ((dialog.ActualHeight > 0.0) ? dialog.ActualHeight : 170.0);
					dialog.Left = (double)(r.Left + r.Right) / 2.0 * num - num3 / 2.0;
					dialog.Top = (double)(r.Top + r.Bottom) / 2.0 * num2 - num4 / 2.0;
				}
				catch
				{
				}
			};
			dialog.WindowStartupLocation = WindowStartupLocation.Manual;
		}
		catch
		{
		}
	}
}
