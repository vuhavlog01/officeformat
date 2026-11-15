using System;
using Xunit;
using VFormat.Core.Common;

namespace VFormat.Tests.CommonTests;

public class OrphanWordHelperTests
{
    [Fact]
    public void FixOrphanWordsInText_Image1_BindsVongDoi()
    {
        string input = "- 100% văn bản đủ điều kiện được xử lý, ký số, phát hành, trao đổi và lưu trữ trên môi trường số; tối thiểu 95% hồ sơ công việc được quản lý điện tử theo vòng đời.\r";
        string result = OrphanWordHelper.FixOrphanWordsInText(input);

        Assert.Contains("vòng\u00a0đời.\r", result);
    }

    [Fact]
    public void FixOrphanWordsInText_Image2_BindsViecLam()
    {
        string input = "- Tổ chức sử dụng thống nhất, nền nếp Hệ thống thông tin điều hành tác nghiệp của Đảng cho toàn bộ quy trình tiếp nhận, luân chuyển văn bản, xử lý công việc và phát hành văn bản của Đảng ủy xã; thực hiện phân quyền chặt chẽ theo vị trí việc làm.\r\n";
        string result = OrphanWordHelper.FixOrphanWordsInText(input);

        Assert.Contains("việc\u00a0làm.\r\n", result);
    }

    [Fact]
    public void FixOrphanWordsInText_Image3_BindsThanhUy()
    {
        string input = "- Triển khai thực hiện nghiêm túc 04 thủ tục hành chính của Đảng trên môi trường điện tử theo Quy định số 338-QĐ/TW của Ban Bí thư và chỉ đạo của Thành ủy.";
        string result = OrphanWordHelper.FixOrphanWordsInText(input);

        Assert.EndsWith("Thành\u00a0ủy.", result);
    }

    [Fact]
    public void FixOrphanWordsInText_MultipleSpacesBeforeLastWord_NormalizesToSingleNBSP()
    {
        string input = "Thực hiện theo vòng    đời.\r";
        string result = OrphanWordHelper.FixOrphanWordsInText(input);

        Assert.Contains("vòng\u00a0đời.\r", result);
        Assert.DoesNotContain("vòng  ", result);
    }

    [Fact]
    public void FixOrphanWordsInText_SoftLineBreak_BindsBothSegments()
    {
        string input = "- Bước 1 việc làm\v- Bước 2 thực hiện\r";
        string result = OrphanWordHelper.FixOrphanWordsInText(input);

        Assert.Contains("việc\u00a0làm\v", result);
        Assert.Contains("thực\u00a0hiện\r", result);
    }

    [Theory]
    [InlineData("QUYETDINH")]
    [InlineData("Hello")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void FixOrphanWordsInText_SingleWordOrEmpty_DoesNotChange(string input)
    {
        string result = OrphanWordHelper.FixOrphanWordsInText(input);
        Assert.Equal(input, result);
    }
}
