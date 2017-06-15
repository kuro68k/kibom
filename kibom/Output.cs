using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.RtfRendering;
using PdfSharp.Pdf;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace kibom
{
	class Output
	{
		public static void OutputXLSX(List<DesignatorGroup> groups, HeaderBlock header, string file)
		{
			ExcelPackage p = new ExcelPackage();
			p.Workbook.Properties.Author = "KiBOM";
			p.Workbook.Properties.Title = header.title;
			p.Workbook.Properties.Company = header.company;
			p.Workbook.Properties.Comments = header.source;

			string sheetName = "Bill of Materials";
			p.Workbook.Worksheets.Add(sheetName);
			ExcelWorksheet ws = p.Workbook.Worksheets[1];
			ExcelRange r;
			ws.Name = sheetName;
			ws.Cells.Style.Font.Size = 11;			// default font for whole sheet
			ws.Cells.Style.Font.Name = "Calibri";

			ws.Column(1).Width = 18;
			ws.Column(2).Width = 35;
			ws.Column(3).Width = 37;
			ws.Column(4).Width = 28;
			ws.Column(5).Width = 37;
			ws.Column(6).Width = 10;

			// header block
			ws.Cells[1, 1].Value = "Title";
			ws.Cells[1, 2].Value = header.title;
			ws.Cells[2, 1].Value = "Date";
			ws.Cells[2, 2].Value = header.date;
			ws.Cells[3, 1].Value = "Source";
			ws.Cells[3, 2].Value = header.source;
			ws.Cells[4, 1].Value = "Revision";
			ws.Cells[4, 2].Value = header.revision;
			ws.Cells[5, 1].Value = "Company";
			ws.Cells[5, 2].Value = header.company;
			r = ws.Cells[1, 1, 5, 1];
			r.Style.Font.Bold = true;
			ws.Cells[1, 2].Style.Font.Bold = true;

			// table header
			ws.Cells[7, 1].Value = "Count";
			ws.Cells[7, 2].Value = "References";
			ws.Cells[7, 3].Value = "Values";
			ws.Cells[7, 4].Value = "MPN";
			ws.Cells[7, 5].Value = "Form factor";
			ws.Cells[7, 6].Value = "Precision";
			r = ws.Cells[7, 1, 7, 6];
			r.Style.Font.Bold = true;
			r.Style.Fill.PatternType = ExcelFillStyle.Solid;
			r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
			r.Style.Font.Color.SetColor(System.Drawing.Color.White);

			int row = 8;
			foreach (DesignatorGroup g in groups)
			{
				// check for groups that are entire "no part"
				bool all_no_part = true;
				foreach (Component c in g.comp_list)
				{
					if (!c.no_part)
						all_no_part = false;
				}
				if (all_no_part)
					continue;

				// header
				DefaultComp def = Component.FindDefaultComp(g.designator);
				if (def != null)
				{
					ws.Cells[row, 1].Value = def.long_name;
					ws.Cells[row, 2].Value = g.comp_list.Count.ToString() + " values";
					if (def.has_default)
						ws.Cells[row, 3].Value = def.default_type + " unless otherwise stated";
				}
				else
				{
					ws.Cells[row, 1].Value = g.designator;
					ws.Cells[row, 2].Value = g.comp_list.Count.ToString();
				}
				ws.Cells[row, 1].Style.Font.Bold = true;
				r = ws.Cells[row, 1, row, 6];
				r.Style.Fill.PatternType = ExcelFillStyle.Solid;
				r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
				row++;

				// component list
				foreach (Component c in g.comp_list)
				{
					if (c.no_part)
						continue;

					string footprint = c.footprint_normalized;
					if (footprint == "")
						footprint = c.footprint;
					ws.Cells[row, 1].Value = c.count + 1;
					ws.Cells[row, 2].Value = c.reference;
					ws.Cells[row, 3].Value = c.value;
					ws.Cells[row, 4].Value = c.part_no;
					ws.Cells[row, 5].Value = footprint;
					ws.Cells[row, 6].Value = c.precision;
					row++;
				}
				row++;
			}

			r = ws.Cells[1, 1, row, 6];
			r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
			r.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
			r.Style.WrapText = true;

			byte[] bin = p.GetAsByteArray();
			File.WriteAllBytes(file, bin);
		}
		
		public static void OutputTSV(List<DesignatorGroup> groups, HeaderBlock header, string file)
		{
			Console.WriteLine("Generating " + file + "...");
			using (StreamWriter sw = new StreamWriter(file))
			{
				sw.WriteLine("Title\t" + header.title);
				sw.WriteLine("Date\t" + header.date);
				sw.WriteLine("Source\t" + header.source);
				sw.WriteLine("Revsision\t" + header.revision);
				if (header.company != "")
					sw.WriteLine("Company\t" + header.company);
				sw.WriteLine("");
				//sw.WriteLine("Type\tDesignation\tValue\tPart Name");

				foreach (DesignatorGroup g in groups)
				{
					// check for groups that are entire "no part"
					bool all_no_part = true;
					foreach (Component c in g.comp_list)
					{
						if (!c.no_part)
							all_no_part = false;
					}
					if (all_no_part)
						continue;

					// header
					//sw.WriteLine("Group: " + g.designator + " (" + g.comp_list.Count.ToString() + ")");
					DefaultComp def = Component.FindDefaultComp(g.designator);
					if (def != null)
					{
						sw.Write(def.long_name);
						sw.Write("\t" + g.comp_list.Count.ToString() + " values");
						if (def.has_default)
							sw.Write("\t" + def.default_type + " unless otherwise stated");
						sw.WriteLine();
					}
					else
						sw.WriteLine(g.designator + "\t" + g.comp_list.Count.ToString());

					// component list
					foreach (Component c in g.comp_list)
					{
						if (c.no_part)
							continue;

						string footprint = c.footprint_normalized;
						if (footprint == "")
							footprint = c.footprint;
						sw.WriteLine(	(c.count + 1).ToString() +
										"\t" + c.reference +
										"\t" + c.value +
										"\t" + c.part_no +
										"\t" + footprint +
										"\t" + c.precision);
					}
					sw.WriteLine();
				}
			}
		}

		public static void OutputPDF(List<DesignatorGroup> groups, HeaderBlock header, string file, bool rtf = false)
		{
			Console.WriteLine("Generating " + file + "...");

			// document setup
			var doc = new Document();
			doc.DefaultPageSetup.PageFormat = PageFormat.A4;
			doc.DefaultPageSetup.Orientation = Orientation.Landscape;
			doc.DefaultPageSetup.TopMargin = "1.5cm";
			doc.DefaultPageSetup.BottomMargin = "1.5cm";
			doc.DefaultPageSetup.LeftMargin = "1.5cm";
			doc.DefaultPageSetup.RightMargin = "1.5cm";
			doc.Styles["Normal"].Font.Name = "Arial";

			var footer = new Paragraph();
			footer.AddTab();
			footer.AddPageField();
			footer.AddText(" of ");
			footer.AddNumPagesField();
			footer.Format.Alignment = ParagraphAlignment.Center;

			// generate content
			var section = doc.AddSection();
			section.Footers.Primary.Add(footer.Clone());
			section.Footers.EvenPage.Add(footer.Clone());
			PDFCreateHeader(ref section, header);
			var para = section.AddParagraph();
			var table = PDFCreateTable(ref section);

			// BOM table
			int i = 1;
			foreach (DesignatorGroup g in groups)
			{
				// check for groups that are entire "no part"
				bool all_no_part = true;
				foreach (Component c in g.comp_list)
				{
					if (!c.no_part)
						all_no_part = false;
				}
				if (all_no_part)
					continue;

				// group header row
				var row = table.AddRow();
				row.Shading.Color = Colors.LightGray;
				row.Cells[0].MergeRight = 6;
				row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

				DefaultComp def = Component.FindDefaultComp(g.designator);
				if (def != null)
				{
					var p = row.Cells[0].AddParagraph(def.long_name);
					p.Format.Font.Bold = true;
					if (def.has_default)
						row.Cells[0].AddParagraph("All " + def.default_type + " unless otherwise stated");
				}
				else
				{
					var p = row.Cells[0].AddParagraph(g.designator);
					p.Format.Font.Bold = true;
				}

				foreach (Component c in g.comp_list)
				{
					if (c.no_part)
						continue;

					row = table.AddRow();
					row.Cells[0].AddParagraph(i.ToString());
					row.Cells[1].AddParagraph(c.count.ToString());
					row.Cells[2].AddParagraph(c.reference);
					row.Cells[3].AddParagraph(c.value);

					string temp = c.footprint_normalized;
					if (c.code != null)
						temp += ", " + c.code;
					if (c.precision != null)
						temp += ", " + c.precision;
					row.Cells[4].AddParagraph(temp);
					//row.Cells[4].AddParagraph(c.footprint_normalized);

					row.Cells[5].AddParagraph(c.part_no);
					row.Cells[6].AddParagraph(c.note);
				}
			}

			// generate PDF file
			if (!rtf)
			{
				var pdfRenderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always);
				pdfRenderer.Document = doc;
				pdfRenderer.RenderDocument();
				pdfRenderer.PdfDocument.Save(file);
			}
			else
			{
				var rtfRenderer = new RtfDocumentRenderer();
				rtfRenderer.Render(doc, file, null);
			}
		}

		static void PDFCreateHeader(ref Section section, HeaderBlock header)
		{
			var table = section.AddTable();

			table.Borders.Width = 0;
			table.TopPadding = 1;
			table.BottomPadding = 2;
			table.LeftPadding = 0;
			table.RightPadding = 10;

			table.AddColumn();
			table.AddColumn("20cm");

			var row = table.AddRow();
			row.Cells[0].AddParagraph("Title");
			row.Cells[0].Format.Font.Bold = true;
			row.Cells[1].AddParagraph(header.title);
			row = table.AddRow();
			row.Cells[0].AddParagraph("Date");
			row.Cells[0].Format.Font.Bold = true;
			row.Cells[1].AddParagraph(header.date);
			row = table.AddRow();
			row.Cells[0].AddParagraph("Source");
			row.Cells[0].Format.Font.Bold = true;
			row.Cells[1].AddParagraph(header.source);
			row = table.AddRow();
			row.Cells[0].AddParagraph("Revision");
			row.Cells[0].Format.Font.Bold = true;
			row.Cells[1].AddParagraph(header.revision);
			if (header.company != "")
			{
				row = table.AddRow();
				row.Cells[0].AddParagraph("Company");
				row.Cells[0].Format.Font.Bold = true;
				row.Cells[1].AddParagraph(header.company);
			}
			table.AddRow();
		}

		static Table PDFCreateTable(ref Section section)
		{
			var table = section.AddTable();

			table.Borders.Width = 1;
			table.TopPadding = 1;
			table.BottomPadding = 2;
			table.LeftPadding = 5;
			table.RightPadding = 5;

			var col = table.AddColumn("1.5cm");	// item
			col.Format.Alignment = ParagraphAlignment.Center;
			col = table.AddColumn("1.5cm");	// quantity
			col.Format.Alignment = ParagraphAlignment.Center;
			col = table.AddColumn("3.5cm");	// reference
			col.Format.Alignment = ParagraphAlignment.Left;
			col = table.AddColumn("4.5cm");	// value
			col.Format.Alignment = ParagraphAlignment.Left;
			col = table.AddColumn("4.5cm");	// type
			col.Format.Alignment = ParagraphAlignment.Left;
			//col = table.AddColumn("2.5cm");	// mechanical/size
			//col.Format.Alignment = ParagraphAlignment.Left;
			col = table.AddColumn("5.5cm");	// manufacturer part number
			col.Format.Alignment = ParagraphAlignment.Left;
			col = table.AddColumn("6cm");	// notes
			col.Format.Alignment = ParagraphAlignment.Left;

			var row = table.AddRow();
			row.HeadingFormat = true;
			row.Format.Alignment = ParagraphAlignment.Center;
			row.Format.Font.Bold = true;
			row.Shading.Color = Colors.LightGray;
			row.Cells[0].AddParagraph("No.");
			row.Cells[1].AddParagraph("Qty.");
			row.Cells[2].AddParagraph("Reference");
			row.Cells[3].AddParagraph("Value");
			row.Cells[4].AddParagraph("Type");
			//row.Cells[5].AddParagraph("Size");
			row.Cells[5].AddParagraph("Manufacturer Part No.");
			row.Cells[6].AddParagraph("Notes");

			return table;
		}

	}
}
