using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Reflection;
using Keio.Utils;

namespace kibom
{
	class DesignatorGroup
	{
		public List<Component> comp_list;
		public string designator;
	}

	class Program
	{
		static void Main(string[] args)
		{
            Version v = Assembly.GetExecutingAssembly().GetName().Version;
            Console.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, @"Kibom {0}.{1} (build {2}.{3})", v.Major, v.Minor, v.Build, v.Revision));

			string filename = "";
			string output_filename = "";
			string path = "";
			string outputs = "";
			string template = "";
			string footer = "";
			string path_txt = "";
			bool no_fit_list = false;
			if (!ParseArgs(args, out filename, out path, out outputs, out output_filename, out template, out footer, out path_txt, out no_fit_list))
				return;

			if (string.IsNullOrEmpty(path_txt)) path_txt = path;
			//Console.WriteLine(path_txt);
			if (!Footprint.LoadSubsFile(path_txt /*path*/ /*"./bin/scripting/plugins/kibom/"*/) ||
				!Component.LoadDefaultsFile(path_txt /*path*/ /*"./bin/scripting/plugins/kibom/"*/))
				return;

			XmlDocument doc = new XmlDocument();
			doc.Load(path + filename);
			if (ParseKicadXML(doc, path, filename, outputs, output_filename, template, footer, path_txt, no_fit_list))
				Console.WriteLine("BOM generated.");
		}

		static bool ParseArgs(string[] args,
							  out string filename, out string path,
							  out string outputs, out string output_filename,
							  out string template, out string footer,
							  out string path_txt,
							  out bool no_fit_list)
		{
			filename = "";
			path = "";
			outputs = "";
			output_filename = "";
			template = "";
			footer = "";
			path_txt = "";
			no_fit_list = false;

			string poutputs = string.Empty;
			string ptemplate = string.Empty;
			string pfooter = string.Empty;
			string ppath_txt = string.Empty;
			string pfilename = string.Empty;
			string poutput_filename = string.Empty;
			bool pnfl = false;

			CmdArgs argProcessor = new CmdArgs() {
				{ new CmdArgument("debug", ArgType.Flag, help: "Generate debug output",
									assign: (dynamic d) => { poutputs += "d"; }) },
				{ new CmdArgument("tsv", ArgType.Flag, help: "Generate tab separated value output",
									assign: (dynamic d) => { poutputs += "t"; }) },
				{ new CmdArgument("pretty", ArgType.Flag, help: "Generate pretty text output",
									assign: (dynamic d) => { poutputs += "q"; }) },
				{ new CmdArgument("pdf", ArgType.Flag, help: "Generate PDF output",
									assign: (dynamic d) => { poutputs += "p"; }) },
				{ new CmdArgument("rtf", ArgType.Flag, help: "Generate RTF output",
									assign: (dynamic d) => { poutputs += "r"; }) },
				{ new CmdArgument("xlsx", ArgType.Flag, help: "Generate XLSX (Excel spreadsheet) output",
									assign: (dynamic d) => { poutputs += "x"; }) },
				{ new CmdArgument("t,template", ArgType.String, help: "Template file for XLSX output",
									assign: (dynamic d) => { ptemplate = (string)d; }) },
				{ new CmdArgument("f,footer", ArgType.String, help: "Footer file for XLSX output",
									assign: (dynamic d) => { pfooter = (string)d; }) },
				{ new CmdArgument("path_txt", ArgType.String, help: "Path for bom_default.txt and bom_subs.txt",
									assign: (dynamic d) => { ppath_txt = (string)d; }) },
				{ new CmdArgument("nfl,no-fit-list", ArgType.Flag, help: "Add a list of no-fit parts",
									assign: (dynamic d) => { pnfl = (bool)d; }) },
				{ new CmdArgument("", ArgType.String, anonymous: true, required: true, parameter_help: "bom.xml",
									assign: (dynamic d) => { pfilename = (string)d; }) },
				{ new CmdArgument("", ArgType.String, anonymous: true, parameter_help: "output file",
									assign: (dynamic d) => { poutput_filename = (string)d; }) },
			};

			string[] remainder;
			if (!argProcessor.TryParse(args, out remainder) || string.IsNullOrEmpty(pfilename))
			{
				argProcessor.PrintHelp();
				return false;
			}

			filename = pfilename;
			output_filename = poutput_filename;
			outputs = poutputs;
			template = ptemplate;
			footer = pfooter;
			path_txt = ppath_txt;
			no_fit_list = pnfl;

			if (!File.Exists(filename))
			{
				Console.WriteLine("File not found.");
				return false;
			}
			path = Path.GetDirectoryName(Path.GetFullPath(filename));
			if (!path.EndsWith("\\"))
				path += "\\";
			filename = Path.GetFileName(filename);

			return true;
		}

		static bool ParseKicadXML(XmlDocument doc, string path, string filename, string outputs, string output_filename, string template, string footer, string path_txt, bool nfl)
		{
			HeaderBlock header = new HeaderBlock();
			if (!header.ParseHeader(doc))
			{
				Console.WriteLine("Could not parse XML header block.");
				return false;
			}
			
			// build component list
			List<Component> comp_list = ParseComponents(doc);
			if (comp_list == null)
				return false;

			// group components by designators and sort by value
			List<DesignatorGroup> groups = Component.BuildDesignatorGroups(comp_list);
			Component.SortDesignatorGroups(ref groups);
			List<DesignatorGroup> merged_groups = Component.MergeComponents(groups);
			Component.SortComponents(ref merged_groups);

			// sort groups alphabetically
			merged_groups.Sort((a, b) => a.designator.CompareTo(b.designator));

			if (output_filename == "")
			{
				string base_filename = Path.GetFileNameWithoutExtension(path + filename);
				if (outputs.Contains('t'))
					Output.OutputTSV(path, merged_groups, header, path + base_filename + ".tsv.txt");
				if (outputs.Contains('q'))
					Output.OutputPretty(path, merged_groups, header, path + base_filename + ".txt");
				if (outputs.Contains('x'))
					Output.OutputXLSX(path, merged_groups, header, path + base_filename + ".xlsx", template, footer, nfl);
				if (outputs.Contains('p'))
					Output.OutputPDF(path, merged_groups, header, path + base_filename + ".pdf");
				if (outputs.Contains('r'))
					Output.OutputPDF(path, merged_groups, header, path + base_filename + ".rtf", true);
			}
			else
			{
				if (outputs.Contains('t'))
					Output.OutputTSV(path, merged_groups, header, output_filename);
				if (outputs.Contains('q'))
					Output.OutputPretty(path, merged_groups, header, output_filename);
				if (outputs.Contains('x'))
					Output.OutputXLSX(path, merged_groups, header, output_filename, template, footer, nfl);
				if (outputs.Contains('p'))
					Output.OutputPDF(path, merged_groups, header, output_filename);
				if (outputs.Contains('r'))
					Output.OutputPDF(path, merged_groups, header, output_filename, true);
			}

			// debug output
			if (outputs.Contains('d'))
			{
				foreach (DesignatorGroup g in merged_groups)
				{
					Console.WriteLine("Group: " + g.designator + " (" + g.comp_list.Count.ToString() + ")");
					DefaultComp def = Component.FindDefaultComp(g.designator);
					if (def != null)
					{
						Console.Write("(" + def.long_name);
						if (def.has_default)
							Console.Write(", " + def.default_type + " unless otherwise stated");
						Console.WriteLine(")");
					}
					foreach (Component c in g.comp_list)
						Console.WriteLine(	"\t" + c.reference +
											"\t" + c.value +
											"\t" + c.footprint_normalized);
					Console.WriteLine();
				}
			}
			return true;
		}

		static List<Component> ParseComponents(XmlDocument doc)
		{
			List<Component> comp_list = new List<Component>();

			XmlNode components_node = doc.DocumentElement.SelectSingleNode("components");
			XmlNodeList comp_nodes = components_node.SelectNodes("comp");
			foreach (XmlNode node in comp_nodes)
			{
				var comp = new Component();
				comp.reference = node.Attributes["ref"].Value;
				comp.designator = comp.reference.Substring(0, comp.reference.IndexOfAny("0123456789".ToCharArray()));
				comp.value = node.SelectSingleNode("value").InnerText;
				comp.numeric_value = Component.ValueToNumeric(comp.value);
				comp.footprint = node.SelectSingleNode("footprint").InnerText;
				
				// normalized footprint
				comp.footprint_normalized = Footprint.substitute(comp.footprint, /*true*/false, true);   //remove_unknown, strip_underscore
                if (comp.footprint_normalized == "no part")
					comp.no_part = true;
				if (comp.footprint.Contains(':'))	// contrains library name
					comp.footprint = comp.footprint.Substring(comp.footprint.IndexOf(':') + 1);
				
				// custom BOM fields
				XmlNode fields = node.SelectSingleNode("fields");
				if (fields != null)
				{
					XmlNodeList fields_nodes = fields.SelectNodes("field");
					foreach (XmlNode field in fields_nodes)
					{
						switch(field.Attributes["name"].Value.ToLower())
						{
							case "bom_footprint":
							//case "bom_partno":
							comp.footprint_normalized = field.InnerText;
							break;

							case "precision":
							comp.precision = field.InnerText;
							break;

							case "bom_note":
							comp.note = field.InnerText;
							break;

							case "bom_partno":
							comp.part_no = field.InnerText;
							break;

							case "code":
							comp.code = field.InnerText;
							break;

							case "bom_no_part":
							if ((field.InnerText.ToLower() == "true") ||
								(field.InnerText.ToLower() == "no part"))
								comp.no_part = true;
							break;

							case "bom_no_fit":
							if ((field.InnerText.ToLower() == "true") ||
								(field.InnerText.ToLower() == "no fit"))
								comp.no_fit = true;
							break;
						}
					}
				}

				if (!comp.no_part)
					comp_list.Add(comp);

				//Console.WriteLine(comp.reference + "\t" + comp.value + "\t" + comp.footprint_normalized);
			}

			return comp_list;
		}
	}
}
