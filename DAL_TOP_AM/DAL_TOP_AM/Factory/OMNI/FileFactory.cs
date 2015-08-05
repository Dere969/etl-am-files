using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace DAL_TOP_AM.Factory.OMNI
{
	public class FileFactory
	{
		#region Select

		public static List<Entities.Trade_OMNI> Select()
		{
			List<Entities.Trade_OMNI> response = new List<Entities.Trade_OMNI>();
			try
			{
				string[] lines = null;
				lines = File.ReadAllLines(FileTypeFactory.cFileName);
				for (Int32 i = 0; i < lines.Length; i++)
				{
					if (i == 0 && FileTypeFactory.cBlnFirstRowContainsHeader)
						continue;
					response.Add(FileTypeFactory.Select(lines[i]));
				}
			}
			catch (Exception ex)
			{
				DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
				throw ex;
			}
			return response;
		}

		#endregion
	}
}