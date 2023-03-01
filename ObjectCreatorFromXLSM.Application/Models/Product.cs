namespace ObjectCreatorFromXLSM.Application.Models;

public class Product {
    public string NamePrefix { get; set; }
    public string Name { get; set; }
    public string Dimension { get; set; }
    public string ProductLink { get; set; }
    public string Id { get; set; }
    public string Size { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
    public float Depth { get; set; }
    public string? VideoLink { get; set; }
    public List<ProductDescription> Descriptions { get; set; }
    public List<string> AdditionalInformations { get; set; }
    public List<ProductTableIntroduction> TableIntroduction { get; set; }
    public List<ProductAvailableColors> AvailableColors { get; set; }
    public List<string> Tags { get; set; }
}