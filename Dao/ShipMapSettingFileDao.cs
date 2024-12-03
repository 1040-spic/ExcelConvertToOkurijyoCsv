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
	class ShipMapSettingFileDao
	{
		private const string KEY_ENV = @"Env";

		public List<ShipMapSettingDto> ReadShipMapSettingFile()
		{
			var list = new List<ShipMapSettingDto>();

			try
			{
				var file = getString("Path.ShipMapSetting");
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

						list.Add(new ShipMapSettingDto()
						{
							ShipCd = line[0].Substring(line[0].IndexOf(":") + 1, 1),
							OwnShipCd = line[1]
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
