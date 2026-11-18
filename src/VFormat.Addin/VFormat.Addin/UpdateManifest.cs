using System;

namespace VFormat.Addin;

internal sealed class UpdateManifest
{
	public Version Version { get; set; }

	public string VersionText { get; set; }

	public string UpdateUrl { get; set; }

	public string Sha256 { get; set; }
}
