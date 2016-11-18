using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

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
			string filename = "";
			string path = "";
			string outputs = "";
			if (!ParseArgs(args, out filename, out path, out outputs))
				return;

			Console.WriteLine(filename);
			Console.WriteLine(path);
			Console.WriteLine(outputs);

			if (!Footprint.LoadSubsFile(path) ||
				!Component.LoadDefaultsFile(path))
				return;

			XmlDocument doc = new XmlDocument();
			doc.Load(path + filename);
			ParseKicadXML(doc, path, filename, outputs);
		}

		static bool ParseArgs(string[] args, out string filename, out string path, out string outputs)
		{
			filename = "";
			path = "";
			outputs = "";

			if (args.Count() < 1)
			{
				Console.WriteLine("kibom <bom.xml> [-tsv] [-pdf] [-rtf] [-debug]");
				return false;
			}

			filename = args[0];
			if (!File.Exists(filename))
			{
				Console.WriteLine("File not found.");
				return false;
			}
			path = Path.GetDirectoryName(filename);
			filename = Path.GetFileName(filename);

			// parse other args
			for (int i = 1; i < args.Count(); i++)
			{
				switch(args[i].ToLower())
				{
					case "-tsv":
					outputs += "t";
					break;

					case "-pdf":
					outputs += "p";
					break;

					case "-rtf":
					outputs += "r";
					break;

					case "-debug":
					outputs += "r";
					break;
					
					default:
					Console.WriteLine("Unknown argument: \"" + args[i] + "\"");
					return false;
				}
			}
			
			return true;
		}

		static bool ParseKicadXML(XmlDocument doc, string path, string filename, string outputs)
		{
			if (!ParseHeader(doc))
				return false;
			
			// build component list
			List<Component> comp_list = ParseComponents(doc);
			if (comp_list == null)
				return false;

			// group components by designators and sort by value
			List<DesignatorGroup> groups = BuildDesignatorGroups(comp_list);
			SortDesignatorGroups(ref groups);
			List<DesignatorGroup> merged_groups = Component.MergeComponents(groups);

			// sort groups alphabetically
			merged_groups.Sort((a, b) => a.designator.CompareTo(b.designator));

			string base_filename = Path.GetFileNameWithoutExtension(path + filename);
			if (outputs.Contains('t'))
				Output.OutputTSV(merged_groups, base_filename + ".txt");
			if (outputs.Contains('p'))
				Output.OutputPDF(merged_groups, base_filename + ".pdf");

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
				comp.footprint_normalized = Footprint.substitute(comp.footprint, true, true);
				
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
						}
					}
				}
				
				if (!comp.footprint.Contains("no part"))		// ignore pad only parts
					comp_list.Add(comp);

				Console.WriteLine(comp.reference + "\t" + comp.value + "\t" + comp.footprint_normalized);
			}

			return comp_list;
		}

		static List<DesignatorGroup> BuildDesignatorGroups(List<Component> comp_list)
		{
			var groups = new List<DesignatorGroup>();

			foreach(Component comp in comp_list)
			{
				bool found = false;
				for (int i = 0; i < groups.Count; i++)
				{
					if (groups[i].designator == comp.designator)
					{
						groups[i].comp_list.Add(comp);
						found = true;
						break;
					}
				}
				if (!found)
				{
					var new_group = new DesignatorGroup();
					new_group.designator = comp.designator;
					new_group.comp_list = new List<Component>();
					new_group.comp_list.Add(comp);
					groups.Add(new_group);
				}
			}

			return groups;
		}

		static void SortDesignatorGroups(ref List<DesignatorGroup> groups)
		{
			foreach (DesignatorGroup g in groups)
			{
				// sort by value
				//g.comp_list.Sort((a, b) => a.value.CompareTo(b.value));
				g.comp_list.Sort((a, b) => a.numeric_value.CompareTo(b.numeric_value));
			}
		}

		static bool ParseHeader(XmlDocument doc)
		{
			XmlNode header_node = doc.DocumentElement.SelectSingleNode("design");
			Console.WriteLine(header_node.SelectSingleNode("date").InnerText);

			return true;
		}

	}
}
