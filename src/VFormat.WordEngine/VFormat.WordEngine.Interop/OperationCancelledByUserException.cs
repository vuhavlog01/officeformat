using System;

namespace VFormat.WordEngine.Interop;

public class OperationCancelledByUserException : OperationCanceledException
{
	public OperationCancelledByUserException()
		: base("Đã dừng theo yêu cầu của người dùng.")
	{
	}
}
