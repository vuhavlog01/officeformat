using System.Globalization;
using Xunit;
using VFormat.UI.Common;
using VFormat.Core.Enums;

namespace VFormat.Tests.UITests;

public class ViewModelTests
{
    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Assertions", "xUnit2000:Constants should be on the left")]
    public void PageMarginViewModel_InitialDefaults_FollowsND30Standard()
    {
        var vm = new PageMarginViewModel();

        Assert.Equal(0, vm.SelectedRegulationIndex);
        Assert.Equal(RulePackType.ND30_2020, vm.SelectedRulePack);
        Assert.Equal(20f, vm.TopMarginMm);
        Assert.Equal(20f, vm.BottomMarginMm);
        Assert.Equal(30f, vm.LeftMarginMm);
        Assert.Equal(20f, vm.RightMarginMm);
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Assertions", "xUnit2000:Constants should be on the left")]
    public void PageMarginViewModel_RegulationChange_UpdatesRulePackAndMargins()
    {
        var vm = new PageMarginViewModel();

        // Switch to HD05 (Party)
        vm.SelectedRegulationIndex = 1;
        Assert.Equal(RulePackType.HD05_VPTW, vm.SelectedRulePack);
        Assert.Equal(20f, vm.TopMarginMm);
        Assert.Equal(20f, vm.BottomMarginMm);
        Assert.Equal(30f, vm.LeftMarginMm);
        Assert.Equal(15f, vm.RightMarginMm);

        // Switch to Military (TT68)
        vm.SelectedRegulationIndex = 2;
        Assert.Equal(RulePackType.TT68_2023_BQP, vm.SelectedRulePack);
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Assertions", "xUnit2000:Constants should be on the left")]
    public void PageMarginViewModel_MarginPropertyChanges_TriggersEvent()
    {
        var vm = new PageMarginViewModel();
        bool topChanged = false;
        vm.PropertyChanged += (s, e) =>
        {
            if (e.PropertyName == nameof(vm.TopMarginMm)) topChanged = true;
        };

        vm.TopMarginMm = 25f;
        Assert.True(topChanged);
        Assert.Equal(25f, vm.TopMarginMm);
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Assertions", "xUnit2000:Constants should be on the left")]
    public void PageNumberViewModel_DefaultsAndCustom_Behavior()
    {
        var vm = new PageNumberViewModel();

        Assert.True(vm.IsFromStart);
        Assert.False(vm.IsCustomStart);
        Assert.True(vm.HideFirstPage);
        Assert.Equal(1, vm.CustomStartNumber);

        // Toggle custom start
        vm.IsCustomStart = true;
        vm.CustomStartNumber = 5;
        Assert.Equal(5, vm.StartNumber);
        Assert.False(vm.FromStart);
    }

    [Theory]
    [InlineData(2.5f, "2.5")]
    [InlineData(0f, "0")]
    [InlineData(14.0f, "14")]
    public void FlexibleFloatConverter_ConvertFloatToString_UsesInvariantCulture(float value, string expected)
    {
        var converter = new FlexibleFloatConverter();
        object result = converter.Convert(value, typeof(string), null, CultureInfo.InvariantCulture);

        Assert.Equal(expected, result?.ToString());
    }

    [Theory]
    [InlineData("2.5", 2.5f)]
    [InlineData("2,5", 2.5f)]
    [InlineData("14", 14f)]
    [InlineData("", 0f)]
    public void FlexibleFloatConverter_ConvertBackStringToFloat_HandlesDotAndComma(string input, float expected)
    {
        var converter = new FlexibleFloatConverter();
        object result = converter.ConvertBack(input, typeof(float), null, CultureInfo.InvariantCulture);

        Assert.Equal(expected, (float)result);
    }
}
