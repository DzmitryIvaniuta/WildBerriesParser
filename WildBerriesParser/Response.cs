using System.Text.Json.Serialization;

namespace WildBerriesParser;

public class Response
{
    [JsonPropertyName("data")] public Data Data { get; set; }
}

public class Data
{
    [JsonPropertyName("products")] public Product[] Products { get; set; }
}

public class Product
{
    [JsonPropertyName("id")] public int Id { get; set; }
    [JsonPropertyName("name")] public string Title { get; set; }
    [JsonPropertyName("brand")] public string Brand { get; set; }
    [JsonPropertyName("priceU")] public int Price { get; set; }
    [JsonPropertyName("feedbacks")] public int Feedbacks { get; set; }
}