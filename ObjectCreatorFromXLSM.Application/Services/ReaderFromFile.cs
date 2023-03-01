using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ObjectCreatorFromXLSM.Application.Models;

namespace ObjectCreatorFromXLSM.Application.Services;

public class ReaderFromFile
{
    public Product ReadFromExcel(string filePath)
    {
        Product product = new();
        List<ProductDescription> descriptions = new();
        List<string> additionalInformations = new();
        List<ProductTableIntroduction> tableIntroduction = new();
        List<ProductAvailableColors> availableColors = new();
        List<string> tags = new();

        StringBuilder sb = new();

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheetPart = null;

            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == Variables.SHEET_NAME);

            if (sheet != null)
            {
                worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            }

            if (worksheetPart != null)
            {
                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

                foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                {
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        var readAllCells = false;
                        switch (GetCellValue(sharedStringTable, cell))
                        {
                            case Variables.PRODUCT_NAME:
                                // pobieranie komórki po prawej
                                Cell namePrefixCell = cell.NextSibling<Cell>();
                                product.NamePrefix = GetCellValue(sharedStringTable, namePrefixCell);

                                Cell nameCell = namePrefixCell.NextSibling<Cell>();
                                product.Name = GetCellValue(sharedStringTable, nameCell);
                                readAllCells = true;
                                break;
                            case Variables.SIZE:
                                // pobieranie poprzedniej komórki zeby sprawdzic czy jest + przy rozmiarze
                                Cell sizeAvailableCell = cell.PreviousSibling<Cell>();
                                string sizeAvailable = GetCellValue(sharedStringTable, sizeAvailableCell);
                                if (sizeAvailable == "+")
                                {
                                    Cell dimensionCell = cell.NextSibling<Cell>();
                                    product.Size = GetCellValue(sharedStringTable, dimensionCell);

                                    Cell productLinkCell = dimensionCell.NextSibling<Cell>();
                                    product.ProductLink = GetCellValue(sharedStringTable, productLinkCell);

                                    Cell idCell = productLinkCell.NextSibling<Cell>();
                                    product.Id = GetCellValue(sharedStringTable, idCell);

                                    Cell sizeCell = idCell.NextSibling<Cell>();
                                    product.Size = GetCellValue(sharedStringTable, sizeCell);

                                    Cell widthCell = sizeCell.NextSibling<Cell>();
                                    product.Width = float.Parse(GetCellValue(sharedStringTable, widthCell));

                                    Cell heightCell = widthCell.NextSibling<Cell>();
                                    product.Height = float.Parse(GetCellValue(sharedStringTable, heightCell));

                                    Cell depthCell = heightCell.NextSibling<Cell>();
                                    product.Depth = float.Parse(GetCellValue(sharedStringTable, depthCell));
                                }
                                readAllCells = true;
                                break;
                            case Variables.VIDEO_LINK:
                                Cell videoTypeLinkCell = cell.NextSibling<Cell>();
                                string videoTypeLink = GetCellValue(sharedStringTable, videoTypeLinkCell);
                                if (videoTypeLink == "yt")
                                {
                                    Cell videoLinkCell = videoTypeLinkCell.NextSibling<Cell>();
                                    product.VideoLink = "https://www.youtube.com/watch?v=" + GetCellValue(sharedStringTable, videoLinkCell);
                                    readAllCells = true;
                                    break;
                                }
                                else if (videoTypeLink == "w")
                                {
                                    Cell videoLinkCell = videoTypeLinkCell.NextSibling<Cell>();
                                    product.VideoLink = GetCellValue(sharedStringTable, videoLinkCell);
                                    readAllCells = true;
                                    break;
                                }
                                else
                                {
                                    readAllCells = true;
                                    break;
                                }
                            case Variables.DESCRIPTIONS:
                                Cell previousDescriptionCell = cell.NextSibling<Cell>();
                                Cell descriptionCell = previousDescriptionCell.NextSibling<Cell>();
                                string description = GetCellValue(sharedStringTable, descriptionCell);

                                Cell previousIsTitleCell = cell.NextSibling<Cell>();
                                Cell isTitleCell = previousIsTitleCell.NextSibling<Cell>();
                                string isTitle = GetCellValue(sharedStringTable, isTitleCell);

                                if (description != "")
                                {
                                    if (isTitle == "h2")
                                    {
                                        descriptions.Add(new ProductDescription
                                        {
                                            Description = description,
                                            IsTitle = true
                                        });
                                    }
                                    else
                                    {
                                        descriptions.Add(new ProductDescription
                                        {
                                            Description = description,
                                            IsTitle = false
                                        });
                                    }
                                }
                                readAllCells = true;
                                break;
                            case Variables.ADDITIONAL_INFORMATION:
                                Cell previousAdditionalInformationCell = cell.NextSibling<Cell>();
                                Cell additionalInformationCell = previousAdditionalInformationCell.NextSibling<Cell>();
                                string additionalInformation = GetCellValue(sharedStringTable, additionalInformationCell);

                                if (additionalInformation != "")
                                {
                                    additionalInformations.Add(additionalInformation);
                                }
                                readAllCells = true;
                                break;
                            case Variables.TABLE:
                                Cell ifAvailableColorsCell = cell.PreviousSibling<Cell>();
                                string ifAvailableColors = GetCellValue(sharedStringTable, ifAvailableColorsCell);
                                if (ifAvailableColors == "+")
                                {
                                    Cell previousColorCell = cell.NextSibling<Cell>();
                                    Cell colorCell = previousColorCell.NextSibling<Cell>();
                                    string color = GetCellValue(sharedStringTable, colorCell);
                                    char[] separators = new char[] { '[', ']' };
                                    string[] colorParts = color.Split(separators, StringSplitOptions.RemoveEmptyEntries);

                                    if (colorParts.Length == 2)
                                    {
                                        availableColors.Add(new ProductAvailableColors
                                        {
                                            Color = colorParts[0],
                                            Description = colorParts[1]
                                        });
                                    }
                                    readAllCells = true;
                                    break;
                                }
                                else if (ifAvailableColors == "-")
                                {
                                    readAllCells = true;
                                    break;
                                }
                                else
                                {
                                    Cell tableIntroductionTitleCell = cell.NextSibling<Cell>();
                                    string tableIntroductionTitle = GetCellValue(sharedStringTable, tableIntroductionTitleCell);
                                    Cell tableIntroductionDescriptionCell = tableIntroductionTitleCell.NextSibling<Cell>();
                                    string tableIntroductionDescription = GetCellValue(sharedStringTable, tableIntroductionDescriptionCell);

                                    if (tableIntroductionTitle != "")
                                    {
                                        tableIntroduction.Add(new ProductTableIntroduction
                                        {
                                            Title = tableIntroductionTitle,
                                            Description = tableIntroductionDescription
                                        });
                                    }
                                    readAllCells = true;
                                    break;
                                }
                            case Variables.TAGS:
                                Cell previousTagCell = cell.NextSibling<Cell>();
                                Cell tagCell = previousTagCell.NextSibling<Cell>();
                                string tag = GetCellValue(sharedStringTable, tagCell);

                                if (tag != "")
                                {
                                    tags.Add(tag);
                                }
                                readAllCells = true;
                                break;
                        }
                        if (readAllCells)
                        {
                            break;
                        }
                    }
                }
            }
        }
        product.Descriptions = descriptions;
        product.AdditionalInformations = additionalInformations;
        product.AvailableColors = availableColors;
        product.TableIntroduction = tableIntroduction;
        product.Tags = tags;
        return product;
    }

    private string GetCellValue(SharedStringTable sharedStringTable, Cell cell)
    {
        if (cell != null)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                int sharedStringIndex = int.Parse(cell.CellValue.Text);
                return sharedStringTable.ElementAt(sharedStringIndex).InnerText;
            }
            else
            {
                return cell.InnerText;
            }
        }
        else
        {
            return "";
        }
    }
}