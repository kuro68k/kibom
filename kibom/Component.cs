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

		static List<DefaultComp> default_list = new List<DefaultComp>();

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
			ns = value.Substring(0, value.IndexOfAny("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()));
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

			//if (value.Contains("m"))
			//	nv /= 1000;
			//else if (value.Contains("n"))
			//	nv /= 10e-6;
			//else if (value.Contains("p"))
			//	nv /= 10e-9;
			//else if (value.Contains("k") || value.Contains("K"))
			//	nv *= 1000;
			//else if (value.Contains("M"))
			//	nv *= 10e6;
			//else if (value.Contains("G"))
			//	nv *= 10e9;
			return nv;
		}

		public static List<DesignatorGroup> MergeComponents(List<DesignatorGroup> groups)
		{
			var new_groups = new List<DesignatorGroup>();

			foreach (DesignatorGroup g in groups)
			{
				var new_g = new DesignatorGroup();
				new_g.comp_list = new List<Component>();
				new_g.designator = g.designator;

				Component last_c = null;

				foreach (Component c in g.comp_list)
				{
					if (last_c == null)					// first item
					{
						last_c = c;
						last_c.count = 1;
					}
					else
					{
						if ((last_c.value != c.value) ||	// new value group
							(last_c.footprint != c.footprint) ||
							(last_c.code != c.code) ||
							(last_c.note != c.note) ||
							(last_c.part_no != c.part_no) ||
							(last_c.precision != c.precision))
						{
							last_c.reference = SortCommaSeparatedString(last_c.reference);
							new_g.comp_list.Add(last_c);
							last_c = c;
							last_c.count = 1;
						}
						else							// same, add to value group
						{
							last_c.reference += ", " + c.reference;
							last_c.count++;
						}
					}
				}

				last_c.reference = SortCommaSeparatedString(last_c.reference);
				new_g.comp_list.Add(last_c);
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
				returnValue = returnValue + stringArray[i] + ",";
			}
			return returnValue.Remove(returnValue.Length - 1, 1);
		}

	}
}
