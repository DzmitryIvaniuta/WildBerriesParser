using System.Net;
using System.Text;
using System.Text.Json;

namespace WildBerriesParser;

public static class Program
{
    private const string UrlBase = "https://search.wb.ru/exactmatch/ru/common/v4/search?curr=rub&dest=-1029256,-102269,-2162196,-1257786&lang=ru&locale=ru&resultset=catalog&sort=popular&suppressSpellcheck=false";

    private static async Task Main()
    {
        try
        {
            Console.WriteLine("Parsing started...");
            
            var filePath = Path.Combine(Environment.CurrentDirectory, $"Test-{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            var keyPath = Path.Combine(Environment.CurrentDirectory, "Keys.txt");
            var keys = await GetSearchKeysAsync(keyPath);
            
            using ExcelHelper helper = new();
            foreach (var key in keys)
            {
                Console.WriteLine($"Parsing by key: {key}");
                
                var url = BuildUrl(key);
                var result = await GetDataAsync(url);
                if (result?.Data.Products == null)
                    continue;

                Console.WriteLine($"The key {key} found {result?.Data.Products.Length} of products");
                
                if (helper.Open(filePath, key))
                {
                    SetData(helper, result.Data.Products);
                }

                helper.ColumnsAutoFit();
                helper.Save();
            }

            Console.WriteLine($"Parsing is finished, the data are saved to {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

        Console.WriteLine("Press any button to end");
        Console.ReadKey();
    }

    private static void SetData(ExcelHelper helper, Product[] products)
    {
        helper.Set(column: "A", row: 1, data: "Title");
        helper.Set(column: "B", row: 1, data: "Brand");
        helper.Set(column: "C", row: 1, data: "Id");
        helper.Set(column: "D", row: 1, data: "Feedbacks");
        helper.Set(column: "E", row: 1, data: "Price");

        for (int i = 0; i < products.Length; i++)
        {
            helper.Set(column: "A", row: i + 2, data: products[i].Title);
            helper.Set(column: "B", row: i + 2, data: products[i].Brand);
            helper.Set(column: "C", row: i + 2, data: products[i].Id);
            helper.Set(column: "D", row: i + 2, data: products[i].Feedbacks);
            helper.Set(column: "E", row: i + 2, data: products[i].Price.ToString()[..^2]);
        }
    }

    private static async Task<List<string>> GetSearchKeysAsync(string path)
    {
        var keys = new List<string>();

        if (!File.Exists(path))
        {
            Console.WriteLine($"{path}: File not found");

            return null!;
        }

        using StreamReader sr = new(path);
        while (!sr.EndOfStream)
        {
            keys.Add(await sr.ReadLineAsync() ?? string.Empty);
        }

        return keys;
    }

    private static string BuildUrl(string key)
    {
        var sb = new StringBuilder();
        sb.Append(UrlBase)
            .Append($"&query={key}");

        return sb.ToString();
    }

    private static async Task<Response> GetDataAsync(string url)
    {
        try
        {
            using HttpClientHandler handler = new()
            {
                AllowAutoRedirect = false,
                AutomaticDecompression =
                    DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.None,
                CookieContainer = new CookieContainer()
            };
            using HttpClient client = new(handler);
            client.DefaultRequestHeaders.Add("Connection", "keep-alive");
            client.DefaultRequestHeaders.Add("Accept", "*/*");

            using var response = await client.GetAsync(url);
            var json = await response.Content.ReadAsStringAsync();
            if (!string.IsNullOrEmpty(json))
            {
                var result = JsonSerializer.Deserialize<Response>(json);

                return result;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

        return null;
    }
}