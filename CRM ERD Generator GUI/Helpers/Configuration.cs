#region Imports

using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using CRM_ERD_Generator_GUI.Model;

#endregion

namespace CRM_ERD_Generator_GUI.Helpers
{
	public static class Configuration
	{
		public static string FileName;

		public static Settings.SerializableSettings LoadConfigs()
		{
			try
			{
				Status.Update("Reading settings ...");

				var file = FileName + ".dat";

				if (!File.Exists(file))
				{
					Status.Update("[ERROR] Settings file does not exist!");
				}

				//Open the file written above and read values from it.
				using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.Read))
				{
					var bformatter = new BinaryFormatter { Binder = new Binder() };
					var settings = (Settings.SerializableSettings)bformatter.Deserialize(stream);

					Status.Update("Done reading settings!");

					return settings;
				}
			}
			catch (Exception ex)
			{
				Status.Update("Failed to read settings! => " + ex.Message);
				return null;
			}
		}

		public static void SaveConfigs(Settings.SerializableSettings settings)
		{
			Status.Update("Writing settings ...");

			var file = FileName + ".dat";

			if (!File.Exists(file))
			{
				File.Create(file).Dispose();
				Status.Update("Created a new settings file.");
			}

			using (var stream = File.Open(file, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
			{
				// clear the file to start from scratch
				stream.SetLength(0);

				var bformatter = new BinaryFormatter();
				bformatter.Serialize(stream, settings);

				Status.Update("Done writing settings!");
			}
		}
	}

	public class Binder : SerializationBinder
	{
		public override Type BindToType(string assemblyName, string typeName)
		{
			var sShortAssemblyName = assemblyName.Split(',')[0];
			var ayAssemblies = AppDomain.CurrentDomain.GetAssemblies();

			return (from ayAssembly in ayAssemblies where sShortAssemblyName == ayAssembly.FullName.Split(',')[0] select ayAssembly.GetType(typeName)).FirstOrDefault();
		}
	}
}
