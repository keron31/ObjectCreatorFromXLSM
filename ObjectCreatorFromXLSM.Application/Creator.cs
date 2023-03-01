using System.Text;
using ObjectCreatorFromXLSM.Application.Models;
using ObjectCreatorFromXLSM.Application.Services;

namespace ObjectCreatorFromXLSM.Application;
public class Creator
{
    private readonly ReaderFromFile _readerFromFile;

    public Creator()
    {
        _readerFromFile = new ReaderFromFile();
    }

    public void CreateObject()
    {
        List<Product> products = new();

        foreach ( string filePath in Directory.GetFiles(Variables.FOLDER_PATH, "*.xlsm") )
        {
            products.Add( _readerFromFile.ReadFromExcel(filePath));
        }

        DisplayObjects(products);
    }

    private static void DisplayObjects(List<Product> products)
    {
        foreach (var product in products)
        {
            Console.WriteLine(product.Name);
            Console.WriteLine(product.NamePrefix);
            Console.WriteLine(product.Dimension);
            Console.WriteLine(product.ProductLink);
            Console.WriteLine(product.Id);
            Console.WriteLine(product.Size);
            Console.WriteLine(product.Width);
            Console.WriteLine(product.Height);
            Console.WriteLine(product.Depth);
            Console.WriteLine(product.VideoLink);
            foreach (var description in product.Descriptions)
            {
                Console.WriteLine(description.Description);
                Console.WriteLine(description.IsTitle);
            }
            foreach (var additionalInformation in product.AdditionalInformations)
            {
                Console.WriteLine(additionalInformation);
            }
            foreach (var tableIntroduction in product.TableIntroduction)
            {
                Console.WriteLine(tableIntroduction.Title);
                Console.WriteLine(tableIntroduction.Description);
            }
            foreach (var availableColor in product.AvailableColors)
            {
                Console.WriteLine(availableColor.Color);
                Console.WriteLine(availableColor.Description);
            }
            foreach (var tag in product.Tags)
            {
                Console.WriteLine(tag);
            }
        }
    }
}
