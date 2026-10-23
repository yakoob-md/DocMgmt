using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Collections.Generic;

namespace DocMgmt.Services;

public class ExtractedDocumentContent
{
	public bool IsExcel { get; set; }
	public List<List<string>> GridRows { get; set; } = new();
	public string RawBodyText { get; set; } = string.Empty;
}

public class NativeDocumentParser
{
	public ExtractedDocumentContent ParseFromStream(Stream stream, string extension)
	{
		var result = new ExtractedDocumentContent();

		// Secure native configurations to mitigate XXE vulnerabilities on Linux hosts
		var xmlSettings = new XmlReaderSettings
		{
			DtdProcessing = DtdProcessing.Prohibit,
			XmlResolver = null
		};

		using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
		{
			if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
			{
				result.IsExcel = true;
				var sharedStrings = new List<string>();

				// 1. Gather master structural shared text strings
				var stringsEntry = archive.GetEntry("xl/sharedStrings.xml");
				if (stringsEntry != null)
				{
					using var sStream = stringsEntry.Open();
					using var reader = XmlReader.Create(sStream, xmlSettings);
					var xmlDoc = new XmlDocument();
					xmlDoc.Load(reader);
					var tNodes = xmlDoc.GetElementsByTagName("t");
					foreach (XmlNode node in tNodes)
					{
						sharedStrings.Add(node.InnerText);
					}
				}

				// 2. Read cell locations inside sheet 1
				var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
				if (sheetEntry != null)
				{
					using var sStream = sheetEntry.Open();
					using var reader = XmlReader.Create(sStream, xmlSettings);
					var xmlDoc = new XmlDocument();
					xmlDoc.Load(reader);
					var rowNodes = xmlDoc.GetElementsByTagName("row");

					foreach (XmlNode rowNode in rowNodes)
					{
						var rowData = new List<string>();
						foreach (XmlNode cellNode in rowNode.ChildNodes)
						{
							string cellVal = "";
							var vNode = cellNode["v"];
							if (vNode != null)
							{
								if (cellNode.Attributes["t"]?.Value == "s")
								{
									int index = int.Parse(vNode.InnerText);
									cellVal = sharedStrings[index];
								}
								else
								{
									cellVal = vNode.InnerText;
								}
							}
							rowData.Add(cellVal);
						}
						result.GridRows.Add(rowData);
					}
				}
			}
			else if (extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
			{
				result.IsExcel = false;
				var docEntry = archive.GetEntry("word/document.xml");
				if (docEntry != null)
				{
					using var dStream = docEntry.Open();
					using var reader = XmlReader.Create(dStream, xmlSettings);
					var xmlDoc = new XmlDocument();
					xmlDoc.Load(reader);
					result.RawBodyText = xmlDoc.DocumentElement?.InnerText ?? "";
				}
			}
		}
		return result;
	}
}
