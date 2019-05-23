﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.RtfRendering;
using PdfSharp.Pdf;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Keio.Utils;

namespace kibom
{
	class Output
	{
		public static void OutputXLSX(string path, List<DesignatorGroup> groups, HeaderBlock header, string file, string template, string footer, bool nfl)
		{
			ExcelPackage p = null;
			ExcelWorksheet ws;
			ExcelRange r;
			int table_x;
			int table_y;

			if (string.IsNullOrEmpty(template))
			{
				XLSXCreateDefaultSheet(ref p, header, out ws);
				table_x = 1;
				table_y = 8;
			}
			else
				XLXSLoadSheet(template, ref p, header, out ws, out table_x, out table_y);


			int row = table_y;
			table_x--;
			foreach (DesignatorGroup g in groups)
			{
				// check for groups that are entirly "no part" or "no fit"
				bool all_no_part = true;
				foreach (Component c in g.comp_list)
				{
					if (!c.no_part && !c.no_fit)
						all_no_part = false;
				}
				if (all_no_part)
					continue;

				// header
				DefaultComp def = Component.FindDefaultComp(g.designator);
				if (def != null)
				{
					ws.Cells[row, table_x + 1].Value = def.long_name;
					ws.Cells[row, table_x + 2].Value = g.comp_list.Count.ToString() + " value(s)";
					if (def.has_default)
						ws.Cells[row, 3].Value = def.default_type + " unless otherwise stated";
				}
				else
				{
					ws.Cells[row, table_x + 1].Value = g.designator;
					ws.Cells[row, table_x + 2].Value = g.comp_list.Count.ToString() + " value(s)";
				}
				XLXSStyleHeader(row, table_x, ref ws);
				row++;

				// component list
				foreach (Component c in g.comp_list)
				{
					if (c.no_part)
						continue;

					string footprint = c.footprint_normalized;
					if (footprint == "")
						footprint = c.footprint;
					ws.Cells[row, table_x + 1].Value = c.count + 1;
					ws.Cells[row, table_x + 2].Value = c.reference;
					ws.Cells[row, table_x + 3].Value = c.value;
					ws.Cells[row, table_x + 4].Value = c.part_no;
					ws.Cells[row, table_x + 5].Value = footprint;
					ws.Cells[row, table_x + 6].Value = c.precision;
					r = ws.Cells[row, table_x + 1, row, table_x + 6];
					r.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.LightGray);
					r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
					r.Style.Border.Left.Color.SetColor(System.Drawing.Color.LightGray);
					row++;
				}
				row++;
			}

			r = ws.Cells[1, table_x + 1, row, table_x + 6];
			r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
			r.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
			r.Style.WrapText = true;

			if (nfl)	// no fit part list
				XLXSNoFitList(ref row, table_x, ref ws, groups);

			XLXSFooter(path, footer, ref ws, ref row);

			// generate output file
			byte[] bin = p.GetAsByteArray();
			File.WriteAllBytes(file, bin);
		}

		// Excel has a limitation where it can't do auto cell sizing with merged cells, so we have to calculate it manually
		// https://stackoverflow.com/questions/41639278/autofit-row-height-of-merged-cell-in-epplus
		private static double XLXSMeasureTextHeight(string text, ExcelFont font, double width)
		{
			if (string.IsNullOrEmpty(text))
				return 0.0;
			var bitmap = new Bitmap(1, 1);
			var graphics = Graphics.FromImage(bitmap);

			var pixelWidth = Convert.ToInt32(width * 7.5);  // 7.5 pixels per excel column width
			var drawingFont = new System.Drawing.Font(font.Name, font.Size);
			var size = graphics.MeasureString(text, drawingFont, pixelWidth);

			// 72 DPI and 96 points per inch.  Excel height in points with max of 409 per Excel requirements.
			return Math.Min(Convert.ToDouble(size.Height) * 72 / 96, 409);
		}

		private static void XLXSStyleHeader(int row, int table_x, ref ExcelWorksheet ws)
		{
			ExcelRange r;
			ws.Cells[row, table_x + 1].Style.Font.Bold = true;
			r = ws.Cells[row, table_x + 1, row, table_x + 6];
			r.Style.Fill.PatternType = ExcelFillStyle.Solid;
			r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
			r.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.LightGray);
		}

		private static void XLXSNoFitList(ref int row, int table_x, ref ExcelWorksheet ws, List<DesignatorGroup> groups)
		{
			row++;
			// header
			ws.Cells[row, table_x + 1].Value = "No fit";
			XLXSStyleHeader(row, table_x, ref ws);
			row++;

			StringBuilder sb = new StringBuilder();
			foreach (DesignatorGroup g in groups)
			{
				bool no_fit_found = false;
				bool newline = true;
				foreach (Component c in g.comp_list)
				{
					if (c.no_fit)
					{
						no_fit_found = true;
						if (!newline)
							sb.Append(", ");
						newline = false;
						sb.Append(c.reference);
					}
				}
				if (no_fit_found)
					sb.Append(Environment.NewLine);
			}

			string s = sb.ToString();
			if (string.IsNullOrEmpty(s))
				s = "None";

			ws.Cells[row, table_x + 1, row, table_x + 6].Merge = true;
			ws.Cells[row, table_x + 1].Style.WrapText = true;
			double width = 0;
			for (int i = table_x + 1; i < table_x + 6; i++)
				width += ws.Column(i).Width;
			double height = XLXSMeasureTextHeight(s + Environment.NewLine + ".", ws.Cells[row, table_x + 1].Style.Font, width);
			ws.Row(row).Height = height;
			ws.Cells[row, table_x + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
			ws.Cells[row, table_x + 1].Value = s;
		}

		private static void XLXSFooter(string path, string footer, ref ExcelWorksheet ws, ref int row)
		{
			// footer
			if (!string.IsNullOrEmpty(footer))
			{
				//ws.Row(row).PageBreak = true;
				row++;

				ExcelPackage footer_p = null;
				ExcelWorksheet footer_ws;

				if (!File.Exists(footer))
				{
					footer = path + footer;
					if (!File.Exists(footer))
						throw new Exception("File not found " + footer);
				}
				FileInfo fi = new FileInfo(footer);
				footer_p = new ExcelPackage(fi);
				footer_ws = footer_p.Workbook.Worksheets[1];

				int footer_row = 1;
				int footer_empty_coint = 0;
				do
				{
					bool empty_row = true;
					for (int x = 1; x < 6; x++)
					{
						if (footer_ws.Cells[footer_row, x].Value != null)
							empty_row = false;
					}
					footer_ws.Cells[footer_row, 1, footer_row, 6].Copy(ws.Cells[row, 1, row, 6]);
					footer_row++;
					row++;

					if (empty_row)
						footer_empty_coint++;
					else
						footer_empty_coint = 0;

				} while (footer_empty_coint < 2);
			}
		}

		private static void XLXSLoadSheet(string template, ref ExcelPackage p, HeaderBlock header, out ExcelWorksheet ws, out int table_x, out int table_y)
		{
			if (!File.Exists(template))
				throw new Exception("File not found " + template);
			FileInfo fi = new FileInfo(template);
			p = new ExcelPackage(fi);
			ws = p.Workbook.Worksheets[1];

			table_x = -1;
			table_y = -1;
			for (int x = 1; x < 50; x++)
			{
				for (int y = 1; y < 100; y++)
				{
					if (ws.Cells[y, x].Value == null)
						continue;
					switch (ws.Cells[y, x].Value.ToString().ToLower())
					{
						case "(title)":
							ws.Cells[y, x].Value = header.title;
							break;
						case "(number)":
							ws.Cells[y, x].Value = header.comment1;
							break;
						case "(documentnumber)":
							ws.Cells[y, x].Value = "not implemented";
							break;
						case "(date)":
							ws.Cells[y, x].Value = header.date;
							break;
						case "(revision)":
							ws.Cells[y, x].Value = header.revision;
							break;
						case "(source)":
							ws.Cells[y, x].Value = header.source;
							break;
						case "(table)":
							table_x = x;
							table_y = y;
							break;
					}
				}
			}
			if (table_x == -1)
				throw new Exception("(table) location not found.");
		}

		private static void XLSXCreateDefaultSheet(ref ExcelPackage p, HeaderBlock header, out ExcelWorksheet ws)
		{
			p = new ExcelPackage();
			p.Workbook.Properties.Author = "KiBOM";
			p.Workbook.Properties.Title = header.title;
			p.Workbook.Properties.Company = header.company;
			p.Workbook.Properties.Comments = header.source;

			string sheetName = "Bill of Materials";
			p.Workbook.Worksheets.Add(sheetName);
			ws = p.Workbook.Worksheets[1];
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
		}

		public static void OutputTSV(string path, List<DesignatorGroup> groups, HeaderBlock header, string file)
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

		private static void PrettyMeasureWidths(List<DesignatorGroup> groups,
												out int count_width,
												out int reference_width,
												out int value_width,
												out int footprint_width,
												out int precision_width)
		{
			count_width = "No.".Length;
			reference_width = "Reference".Length;
			value_width = "Value".Length;
			footprint_width = "Footprint".Length;
			precision_width = "Precision".Length;

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

				foreach (Component c in g.comp_list)
				{
					count_width = Math.Max(count_width, (c.count + 1).ToString().Length);
					reference_width = Math.Max(reference_width, c.reference.Length);
					value_width = Math.Max(value_width, c.value.Length);
					if (!string.IsNullOrEmpty(c.precision))
						precision_width = Math.Max(precision_width, c.precision.Length);

					string footprint = c.footprint_normalized;
					if (footprint == "")
						footprint = c.footprint;
					footprint_width = Math.Max(footprint_width, footprint.Length);
				}
			}

			count_width = Math.Min(count_width, 10);
			reference_width = Math.Min(reference_width, 20);
			value_width = Math.Min(value_width, 30);
			footprint_width = Math.Min(footprint_width, 30);
			precision_width = Math.Min(precision_width, 20);
		}

		public static void OutputPretty(string path, List<DesignatorGroup> groups, HeaderBlock header, string file)
		{
			Console.WriteLine("Generating " + file + "...");
			using (StreamWriter sw = new StreamWriter(file))
			{
				int count_width;
				int reference_width;
				int value_width;
				int footprint_width;
				int precision_width;
				PrettyMeasureWidths(groups, out count_width, out reference_width, out value_width, out footprint_width, out precision_width);

				sw.WriteLine("Title:      " + header.title);
				sw.WriteLine("Date:       " + header.date);
				sw.WriteLine("Source:     " + header.source);
				sw.WriteLine("Revsision:  " + header.revision);
				if (header.company != "")
					sw.WriteLine("Company:    " + header.company);
				sw.WriteLine("");

				//sw.WriteLine("No.  Designation    Value          Footprint     Precision");
				sw.Write(TextUtils.FixedLengthString("No.", ' ', count_width) + "  ");
				sw.Write(TextUtils.FixedLengthString("Reference", ' ', reference_width) + "  ");
				sw.Write(TextUtils.FixedLengthString("Value", ' ', value_width) + "  ");
				sw.Write(TextUtils.FixedLengthString("Footprint", ' ', footprint_width) + "  ");
				sw.WriteLine("Precision");
				sw.WriteLine("");

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
						sw.Write("[ " + def.long_name);
						sw.Write(", " + g.comp_list.Count.ToString() + (g.comp_list.Count > 1 ? " values" : " value"));
						if (def.has_default)
							sw.Write(", " + def.default_type + " unless otherwise stated");
						sw.WriteLine(" ]");
					}
					else
						sw.WriteLine("[ " + g.designator + ", " + g.comp_list.Count.ToString() +
									 (g.comp_list.Count > 1 ? " values" : " value") + " ]");

					sw.WriteLine(new string('-', count_width + 2 + reference_width + 2 + value_width + 2 + footprint_width + 2 + precision_width));

					// component list
					foreach (Component c in g.comp_list)
					{
						if (c.no_part)
							continue;

						string footprint = c.footprint_normalized;
						if (footprint == "")
							footprint = c.footprint;
						sw.Write(TextUtils.FixedLengthString((c.count + 1).ToString(), ' ', count_width) + "  ");

						string reference = TextUtils.Reformat(c.reference, reference_width);
						int split_point = reference.IndexOf('\n');
						if (split_point == -1)
							sw.Write(TextUtils.FixedLengthString(reference, ' ', reference_width) + "  ");
						else
							sw.Write(TextUtils.FixedLengthString(reference.Substring(0, split_point - 1), ' ', reference_width) + "  ");
						
						sw.Write(TextUtils.FixedLengthString(c.value, ' ', value_width) + "  ");
						//sw.Write(TextUtils.FixedLengthString(c.part_no, ' ', 10));
						sw.Write(TextUtils.FixedLengthString(footprint, ' ', footprint_width) + "  ");
						if (!string.IsNullOrEmpty(c.precision))
							sw.Write(TextUtils.FixedLengthString(c.precision, ' ', precision_width));
						sw.WriteLine();

						if (split_point != -1)	// need to do the rest of the references
						{
							string indent = new string(' ', count_width + 2);
							do
							{
								reference = reference.Substring(split_point + 1);
								split_point = reference.IndexOf('\n');
								if (split_point == -1)
									sw.WriteLine(indent + reference.Trim());
								else
									sw.WriteLine(indent + reference.Substring(0, split_point - 1).Trim());
							} while (split_point != -1);
						}
					}
					sw.WriteLine();
				}
			}
		}

		public static void OutputPDF(string path, List<DesignatorGroup> groups, HeaderBlock header, string file, bool rtf = false)
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
					row.Cells[0].AddParagraph(i++.ToString());
					row.Cells[1].AddParagraph((c.count + 1).ToString());
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
