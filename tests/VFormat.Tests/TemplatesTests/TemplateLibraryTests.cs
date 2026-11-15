using Xunit;
using VFormat.Core.Templates;

namespace VFormat.Tests.TemplatesTests;

public class TemplateLibraryTests
{
    [Fact]
    public void TemplateLibraryService_Instance_IsAccessible()
    {
        var service = TemplateLibraryService.Instance;
        Assert.NotNull(service);

        var categories = service.GetCategories();
        Assert.NotNull(categories);

        var allTemplates = service.GetAllTemplates();
        Assert.NotNull(allTemplates);
    }

    [Fact]
    public void TemplateLibraryService_SearchEmpty_ReturnsEmptyOrAll()
    {
        var service = TemplateLibraryService.Instance;
        var results = service.SearchAll("tờ trình");
        Assert.NotNull(results);
    }
}
