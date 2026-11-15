using System;
using System.Runtime.InteropServices;

namespace VFormat.WordEngine.Interop;

public static class OperationProgress
{
	private sealed class Context
	{
		public Action<int, string> Report;

		public Func<bool> IsCancelRequested;

		public int Steps;

		public int Expected;

		public int LastTick;

		public int LastPercent;

		public string LastMessage;

		public bool Cancelled;

		public bool HasAbsolute;

		public Context Previous;
	}

	private struct NativeMessage
	{
		public IntPtr hwnd;

		public uint message;

		public IntPtr wParam;

		public IntPtr lParam;

		public uint time;

		public int ptX;

		public int ptY;
	}

	private sealed class Scope : IDisposable
	{
		private Context _ctx;

		public Scope(Context ctx)
		{
			_ctx = ctx;
		}

		public void Dispose()
		{
			if (_ctx != null)
			{
				if (_current == _ctx)
				{
					_current = _ctx.Previous;
				}
				_ctx = null;
			}
		}
	}

	private const int ReportIntervalMs = 100;

	private const uint PM_NOREMOVE = 0u;

	[ThreadStatic]
	private static Context _current;

	public static bool IsActive => _current != null;

	public static bool WasCancelled
	{
		get
		{
			if (_current != null)
			{
				return _current.Cancelled;
			}
			return false;
		}
	}

	public static IDisposable Begin(Action<int, string> report, Func<bool> isCancelRequested)
	{
		Context context = new Context();
		context.Report = report;
		context.IsCancelRequested = isCancelRequested;
		context.LastTick = Environment.TickCount - 100;
		context.Previous = _current;
		return new Scope(_current = context);
	}

	public static void Expect(int steps)
	{
		Context current = _current;
		if (current != null)
		{
			current.Expected = Math.Max(0, steps);
			current.Steps = 0;
		}
	}

	public static void Step(string message = null)
	{
		Context current = _current;
		if (current != null)
		{
			current.Steps++;
			int percent;
			if (current.Expected > 0)
			{
				percent = Math.Min(99, current.Steps * 100 / current.Expected);
			}
			else
			{
				percent = ((!current.HasAbsolute) ? ((int)(95.0 * (double)current.Steps / ((double)current.Steps + 40.0))) : current.LastPercent);
			}
			Pump(current, percent, message);
		}
	}

	public static void Report(int current, int total, string message)
	{
		Context current2 = _current;
		if (current2 != null)
		{
			current2.HasAbsolute = true;
			int percent = ((total > 0) ? ((int)Math.Round(100.0 * (double)current / (double)total)) : current);
			Pump(current2, percent, message);
		}
	}

	public static void ThrowIfCancelled()
	{
		Context current = _current;
		if (current != null)
		{
			CheckCancel(current);
		}
	}

	private static void Pump(Context ctx, int percent, string message)
	{
		CheckCancel(ctx);
		if (percent < 0)
		{
			percent = 0;
		}
		if (percent > 100)
		{
			percent = 100;
		}
		if (percent > ctx.LastPercent)
		{
			ctx.LastPercent = percent;
		}
		if (!string.IsNullOrEmpty(message))
		{
			ctx.LastMessage = message;
		}
		int tickCount = Environment.TickCount;
		if (tickCount - ctx.LastTick < 100)
		{
			return;
		}
		ctx.LastTick = tickCount;
		KeepThreadResponsive();
		if (ctx.Report == null)
		{
			return;
		}
		try
		{
			ctx.Report(ctx.LastPercent, ctx.LastMessage);
		}
		catch
		{
		}
	}

	private static void CheckCancel(Context ctx)
	{
		if (!ctx.Cancelled && ctx.IsCancelRequested != null)
		{
			bool flag = false;
			try
			{
				flag = ctx.IsCancelRequested();
			}
			catch
			{
			}
			if (flag)
			{
				ctx.Cancelled = true;
			}
		}
		if (ctx.Cancelled)
		{
			throw new OperationCancelledByUserException();
		}
	}

	[DllImport("user32.dll")]
	private static extern bool PeekMessage(out NativeMessage msg, IntPtr hWnd, uint filterMin, uint filterMax, uint removeMsg);

	private static void KeepThreadResponsive()
	{
		try
		{
			PeekMessage(out var _, IntPtr.Zero, 0u, 0u, 0u);
		}
		catch
		{
		}
	}
}
