using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace kibom
{
	class Sub
	{
		public string search_substring;
		public string replacement_string;

		public Sub(string search, string replace)
		{
			search_substring = search;
			replacement_string = replace;
		}
	}

	class Footprint
	{
		static List<Sub> sub_list = new List<Sub>();

		// load bom_subs.txt
		public static bool LoadSubsFile(string path)
		{
			try
			{
				using (StreamReader sr = new StreamReader(path + "bom_subs.txt"))
				{
					int line_no = 0;
					while (!sr.EndOfStream)
					{
						line_no++;
						string line = sr.ReadLine();
						line = line.Trim();
						if ((line == "") || (line.StartsWith("#")))
							continue;

						if (!line.Contains('\t'))
						{
							Console.WriteLine("Line {0} did not contain a tab character.", line_no);
							return false;
						}
						string search = line.Substring(0, line.IndexOf('\t'));
						search = search.Trim();
						string replace = line.Substring(line.IndexOf('\t'));
						replace = replace.Trim();

						sub_list.Add(new Sub(search, replace));
					}
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("bom_subs.txt could not be read:");
				Console.WriteLine(e.Message);
				return false;
			}
			return true;
		}

		// remove_unknown=true returns "" when substring not found
		public static string substitute(string s, bool remove_unknown = false, bool strip_underscore = false)
		{
			for(int i = 0; i < sub_list.Count(); i++)
			{
				if (s.Contains(sub_list[i].search_substring))
				{
					if (strip_underscore)
						return sub_list[i].replacement_string.Replace('_', ' ');
					return sub_list[i].replacement_string;
				}
			}
			if (remove_unknown)
				return "";
			
			if (strip_underscore)
				return s.Replace('_', ' ');
			return s;
		}
	}
}
