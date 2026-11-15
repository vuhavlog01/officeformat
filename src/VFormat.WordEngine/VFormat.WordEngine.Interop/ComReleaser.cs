using System.Runtime.InteropServices;

namespace VFormat.WordEngine.Interop;

public static class ComReleaser
{
	public static void Release<T>(ref T comObj) where T : class
	{
		if (comObj == null)
		{
			return;
		}
		try
		{
			if (Marshal.IsComObject(comObj))
			{
				Marshal.ReleaseComObject(comObj);
			}
		}
		catch
		{
		}
		finally
		{
			comObj = null;
		}
	}

	public static void FinalRelease<T>(ref T comObj) where T : class
	{
		if (comObj == null)
		{
			return;
		}
		try
		{
			if (Marshal.IsComObject(comObj))
			{
				Marshal.FinalReleaseComObject(comObj);
			}
		}
		catch
		{
		}
		finally
		{
			comObj = null;
		}
	}

	public static void ReleaseAll(params object[] comObjs)
	{
		if (comObjs == null)
		{
			return;
		}
		foreach (object obj in comObjs)
		{
			if (obj == null)
			{
				continue;
			}
			try
			{
				if (Marshal.IsComObject(obj))
				{
					Marshal.ReleaseComObject(obj);
				}
			}
			catch
			{
			}
		}
	}
}
