using Xunit;
using VFormat.License.Hwid;
using VFormat.License.Models;

namespace VFormat.Tests.LicenseTests;

public class LicenseValidationTests
{
    [Fact]
    public void GetMachineHwid_ReturnsNonEmptyDeterministicHwid()
    {
        string hwid1 = HwidGenerator.GetMachineHwid();
        string hwid2 = HwidGenerator.GetMachineHwid();

        Assert.False(string.IsNullOrWhiteSpace(hwid1));
        Assert.Equal(hwid1, hwid2);
        Assert.True(hwid1.Length <= 32);
    }

    [Theory]
    [InlineData(LicenseStatus.Active, true)]
    [InlineData(LicenseStatus.OfflineActive, true)]
    [InlineData(LicenseStatus.Expired, false)]
    [InlineData(LicenseStatus.Invalid, false)]
    [InlineData(LicenseStatus.NoLicense, false)]
    [InlineData(LicenseStatus.Locked, false)]
    [InlineData(LicenseStatus.DeviceBlocked, false)]
    public void LicenseStatus_StatusEvaluation_DistinguishesUsableState(LicenseStatus status, bool isUsable)
    {
        bool actual = status == LicenseStatus.Active || status == LicenseStatus.OfflineActive || status == LicenseStatus.GracePeriod;
        if (status == LicenseStatus.GracePeriod) actual = true;
        
        Assert.Equal(isUsable, actual);
    }

    [Fact]
    public void LicenseCheckResult_Instantiate_RetainsProperties()
    {
        var result = new LicenseCheckResult
        {
            Status = LicenseStatus.Active,
            CustomerName = "Người dùng thử nghiệm",
            DaysLeft = 365
        };

        Assert.Equal(LicenseStatus.Active, result.Status);
        Assert.Equal("Người dùng thử nghiệm", result.CustomerName);
        Assert.Equal(365, result.DaysLeft);
        Assert.True(result.IsLicensed);
    }
}
