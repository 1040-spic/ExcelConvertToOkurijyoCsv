using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Configuration;
using ExcelConvertToOkumarukunnCsv.Dto;
using System.Globalization;

namespace ExcelConvertToOkumarukunnCsv.Dao
{
	class SlipMapSettingFileDao
	{
		private const string KEY_ENV = @"Env";

		public List<SlipMapSettingDto> ReadSlipMapSettingFile()
		{
			var list = new List<SlipMapSettingDto>();

			try
			{
				var file = getString("Path.SlipMapSetting");
				using (var fs = new FileStream(
							file,
							FileMode.Open,
							FileAccess.Read,
							FileShare.ReadWrite
						)
					)
				{
					var sr = new StreamReader(fs, Encoding.UTF8);

					while (sr.Peek() != -1)
					{
						var line = sr.ReadLine().Split('\t');

						list.Add(new SlipMapSettingDto()
						{
							SlipType = line[0],
							Text = line[1]
						}
						);
					}

					sr.Close();
				}
			}
			catch
			{
				throw;
			}
			finally
			{
			}

			return list;
		}

		private string getString(string key)
		{
			var val = ConfigurationManager.AppSettings.Get(key);

			return val;
		}
	}
}
