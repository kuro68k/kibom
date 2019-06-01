using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace kibom
{
	class DefaultComp
	{
		public string designator;
		public string long_name;
		public string default_type;
		public bool has_default = false;

		public DefaultComp(string _designator, string _long_name, string _default_type)
		{
			designator = _designator;
			long_name = _long_name;
			default_type = _default_type;
			if (default_type.ToUpper() != "N/A")
				has_default = true;
		}
	}

	class Component
	{
		public string reference;
		public string designator;
		public string value;
		public double numeric_value;
		public string footprint;
		public string footprint_normalized;
		public string precision;
		public int count;
		public string part_no = "";
		public string note = "";
		public string code;
        public bool no_part = false;
		public bool no_fit = false;

		static List<DefaultComp> default_list = new List<DefaultComp>();

		public Component(string reference, string value, string footprint)
		{
			this.reference = reference;
			this.value = value;
			this.designator = this.reference.Substring (0, this.reference.IndexOfAny ("0123456789".ToCharArray ()));
			this.numeric_value = Component.ValueToNumeric(this.value);

			// normalized footprint
	    		if (footprint == null) {
				Console.Error.WriteLine("No footprint defined on node '" + this.reference + "." + this.value);
				this.footprint = "no part";
				this.footprint_normalized = "no part";
			}
			else {
				this.footprint_normalized = Footprint.substitute(footprint, true, true);
				if (footprint.Contains (':')) {     // contrains library name
					this.footprint = footprint.Substring (footprint.IndexOf (':') + 1);
				} 
				else {
					this.footprint = footprint;
				}
			}
			this.no_part = (this.footprint_normalized == "no part");
		}

		static public DefaultComp FindDefaultComp(string designator)
		{
			foreach (DefaultComp d in default_list)
			{
				if (d.designator == designator)
					return d;
			}
			return null;
		}

		public static bool LoadDefaultsFile(string path)
		{
			try
			{
				using (StreamReader sr = new StreamReader(path + "bom_defaults.txt"))
				{
					int line_no = 0;
					while (!sr.EndOfStream)
					{
						line_no++;
						string line = sr.ReadLine();
						line = line.Trim();
						if ((line == "") || (line.StartsWith("#")))
							continue;

						string designator = line.Trim();
						string long_name = sr.ReadLine().Trim();
						string default_type = sr.ReadLine().Trim();

						default_list.Add(new DefaultComp(designator, long_name, default_type));
					}
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("bom_defaults.txt could not be read:");
				Console.WriteLine(e.Message);
				return false;
			}
			return true;
		}

		public static double ValueToNumeric(string value)
		{
			double nv;	// nominal value
			string ns;	// numeric string
			string si;	// SI units
			int idx = value.IndexOfAny("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray());
			if (idx != -1)
				ns = value.Substring(0, idx);
			else
				ns = "";
			si = value.Trim("0123456789".ToCharArray());
			
			if (!double.TryParse(ns, out nv))
				return -1;
			//si = si.ToLower();
			si = si.Substring(0, 1) + si.Substring(1).ToLower();

			if ((si == "p") ||
				(si == "pf") ||
				(si == "pv") ||
				(si == "ph"))
				nv /= 10e12;
			if ((si == "n") ||
				(si == "nf") ||
				(si == "nv") ||
				(si == "nh"))
				nv /= 10e9;
			if ((si == "u") ||
				(si == "uf") ||
				(si == "uv") ||
				(si == "uh"))
				nv /= 10e6;
			if ((si == "m") ||
				(si == "mf") ||
				(si == "mv") ||
				(si == "mh"))
				nv /= 10e3;
			if ((si == "k") ||
				(si == "kf") ||
				(si == "kv") ||
				(si == "kh"))
				nv *= 10e3;
			if ((si == "M") ||
				(si == "Mf") ||
				(si == "Mv") ||
				(si == "Mh"))
				nv *= 10e6;
			if ((si == "G") ||
				(si == "Gf") ||
				(si == "Gv") ||
				(si == "Gh"))
				nv *= 10e9;
			if ((si == "T") ||
				(si == "Tf") ||
				(si == "Tv") ||
				(si == "Th"))
				nv *= 10e12;
			return nv;
		}

#region Group Building

		// create groups of components with the same designator, unsorted
		public static List<DesignatorGroup> BuildDesignatorGroups(List<Component> comp_list)
		{
			var groups = new List<DesignatorGroup>();

			foreach (Component comp in comp_list)
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

		// sort DesignatorGroup by values
		public static void SortDesignatorGroups(ref List<DesignatorGroup> groups)
		{
			foreach (DesignatorGroup g in groups)
			{
				// sort by value
				//if (g.designator == "U")
				//	DumpDesignatorGroup(g);
				g.comp_list.Sort((a, b) => a.numeric_value.CompareTo(b.numeric_value));
				//if (g.designator == "U")
				//{
				//	Console.WriteLine();
				//	DumpDesignatorGroup(g);
				//}
			}
		}

		private static void DumpDesignatorGroup(DesignatorGroup g)
		{
			foreach (Component c in g.comp_list)
				Console.WriteLine(c.reference);
		}

#endregion

#region Merging

		// Designator groups contain multiple components with the same designators.
		// Sort the components into batches of identical ones.
		public static List<DesignatorGroup> MergeComponents(List<DesignatorGroup> groups)
		{
			var new_groups = new List<DesignatorGroup>();

			foreach (DesignatorGroup g in groups)
			{
				var new_g = new DesignatorGroup();
				new_g.comp_list = new List<Component>();
				new_g.designator = g.designator;

				foreach (Component c in g.comp_list)
				{
					if (new_g.comp_list.Count == 0)
						new_g.comp_list.Add(c);
					else
					{
						// search through existing lists for matching components
						int found = -1;
						//if (c.designator == "U")
						//	System.Diagnostics.Debugger.Break();
						for (int i = 0; i < new_g.comp_list.Count; i++)
						{
							if ((new_g.comp_list[i].value == c.value) &&
								(new_g.comp_list[i].footprint == c.footprint) &&
								(new_g.comp_list[i].code == c.code) &&
								(new_g.comp_list[i].note == c.note) &&
								(new_g.comp_list[i].part_no == c.part_no) &&
								(new_g.comp_list[i].precision == c.precision))
							{
								found = i;
								break;
							}
						}

						if (found == -1)	// create new value group
							new_g.comp_list.Add(c);
						else				// add to existing group
						{
							new_g.comp_list[found].reference += ", " + c.reference;
							new_g.comp_list[found].count++;
						}
					}
				}

				// sort references
				for (int i = 0; i < new_g.comp_list.Count; i++)
					new_g.comp_list[i].reference = SortCommaSeparatedString(new_g.comp_list[i].reference);
				
				// start a new designator group
				new_groups.Add(new_g);
			}

			return new_groups;
		}

		static string SortCommaSeparatedString(string s)
		{
			string[] stringArray = s.Split(',');
			int[] intArray = new int[stringArray.Count()];

			if (stringArray.Count() < 2)
				return s;

			for (int i = 0; i < stringArray.Count(); i++)
			{
				if (!int.TryParse(Regex.Replace(stringArray[i], "[^0-9.]", ""), out intArray[i]))
					intArray[i] = -1;
			}

			bool sorted;
			do
			{
				sorted = true;
				for (int i = 1; i < stringArray.Count(); i++)
				{
					if (intArray[i - 1] > intArray[i])
					{
						sorted = false;
						int tempi = intArray[i - 1];
						intArray[i - 1] = intArray[i];
						intArray[i] = tempi;
						string temps = stringArray[i - 1];
						stringArray[i - 1] = stringArray[i];
						stringArray[i] = temps;
					}
				}
			} while (!sorted);

			string returnValue = "";
			for (int i = stringArray.GetLowerBound(0); i <= stringArray.GetUpperBound(0); i++)
			{
				returnValue = returnValue + stringArray[i].Trim() + ", ";
			}
			return returnValue.Remove(returnValue.Length - 2, 1);
		}

		public static void SortComponents(ref List<DesignatorGroup> groups)
		{
			foreach (DesignatorGroup g in groups)
			{
				g.comp_list.Sort((a, b) => CompareDesignators(a.reference, b.reference));
			}
		}

		static int CompareDesignators(string a, string b)
		{
			if (a.Contains(','))
				a = a.Substring(0, a.IndexOf(','));
			if (b.Contains(','))
				b = b.Substring(0, b.IndexOf(','));
			int ai, bi;
			int.TryParse(a.Substring(1), out ai);
			int.TryParse(b.Substring(1), out bi);
			return ai.CompareTo(bi);
		}

#endregion

	}
}
