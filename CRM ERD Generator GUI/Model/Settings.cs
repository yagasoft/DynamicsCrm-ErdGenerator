#region Imports

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.ServiceModel.Description;
using System.Text;
using CRM_ERD_Generator_GUI.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;

#endregion

namespace CRM_ERD_Generator_GUI.Model
{
	[Serializable]
	public class Settings : INotifyPropertyChanged
	{
		[Serializable]
		public class SerializableSettings
		{
			public bool _UseSSL;
			public bool _UseIFD;
			public bool _UseOnline;
			public bool _UseOffice365;
			public string _EntitiesToIncludeString;
			public string _CrmOrg;
			public string _Password;
			public string _Username;
			public string _Domain;
			public string _Namespace;
			public string _ProjectName;
			public string _ServerName;
			public string _ServerPort;
			public string _HomeRealm;
			public bool _UseWindowsAuth;
			public string _EntitiesString;
			public string _SelectPrefixes;
			public bool _SplitFiles;
			public bool _RelateSelectedOnly;
			public bool _CustomOnly;
			public int _Depth;
		}

		public SerializableSettings SerializedSettings
		{
			get
			{
				return new SerializableSettings
				       {
					       _UseSSL = UseSSL,
					       _UseIFD = UseIFD,
					       _UseOnline = UseOnline,
					       _UseOffice365 = UseOffice365,
					       _EntitiesToIncludeString = EntitiesToIncludeString,
					       _CrmOrg = CrmOrg,
					       _Password = Password,
					       _Username = Username,
					       _Domain = Domain,
					       _Namespace = Namespace,
					       _ProjectName = ProjectName,
					       _ServerName = ServerName,
					       _ServerPort = ServerPort,
					       _HomeRealm = HomeRealm,
					       _UseWindowsAuth = UseWindowsAuth,
					       _EntitiesString = EntitiesString,
					       _SelectPrefixes = SelectPrefixes,
						   _SplitFiles = SplitFiles,
						   _RelateSelectedOnly = RelateSelectedOnly,
						   _CustomOnly = CustomOnly,
						   _Depth = Depth
				       };
			}

			set
			{
				if (value == null)
				{
					return;
				}

				UseSSL = value._UseSSL;
				UseIFD = value._UseIFD;
				UseOnline = value._UseOnline;
				UseOffice365 = value._UseOffice365;
				EntitiesToIncludeString = value._EntitiesToIncludeString ?? EntitiesToIncludeString;
				CrmOrg = value._CrmOrg ?? CrmOrg;
				Password = value._Password ?? Password;
				Username = value._Username ?? Username;
				Domain = value._Domain ?? Domain;
				Namespace = value._Namespace ?? Namespace;
				ProjectName = value._ProjectName ?? ProjectName;
				ServerName = value._ServerName ?? ServerName;
				ServerPort = value._ServerPort ?? ServerPort;
				HomeRealm = value._HomeRealm ?? HomeRealm;
				UseWindowsAuth = value._UseWindowsAuth;
				EntitiesString = value._EntitiesString ?? EntitiesString;
				SelectPrefixes = value._SelectPrefixes ?? SelectPrefixes;
				SplitFiles = value._SplitFiles;
				RelateSelectedOnly = value._RelateSelectedOnly;
				CustomOnly = value._CustomOnly;
				Depth = value._Depth;
			}
		}

		public Settings(SerializableSettings serSettings = null)
		{
			EntityList = new ObservableCollection<string>();
			EntitiesSelected = new ObservableCollection<string>();

			CrmSdkUrl = @"https://disco.crm.dynamics.com/XRMServices/2011/Discovery.svc";
			ProjectName = "";
			Domain = "";
			T4Path = "";
			Template = "";
			CrmOrg = "";
			EntitiesString = "account,contact,lead,opportunity,systemuser";
			EntitiesToIncludeString = "account,contact,lead,opportunity,systemuser";
			OutputPath = "";
			Username = "@x.onmicrosoft.com";
			Password = "";
			Namespace = "";
			Dirty = false;

			SerializedSettings = serSettings;
		}

		#region boiler-plate INotifyPropertyChanged

		public event PropertyChangedEventHandler PropertyChanged;

		protected virtual void OnPropertyChanged(string propertyName)
		{
			var handler = PropertyChanged;
			if (handler != null)
			{
				handler(this, new PropertyChangedEventArgs(propertyName));
			}
		}

		protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
		{
			if (EqualityComparer<T>.Default.Equals(field, value))
			{
				return false;
			}
			field = value;
			Dirty = true;
			OnPropertyChanged(propertyName);
			return true;
		}

		#endregion

		private bool _UseSSL;
		private bool _UseIFD;
		private bool _UseOnline;
		private bool _UseOffice365;
		private string _EntitiesToIncludeString;
		private string _CrmOrg;
		private string _Password;
		private string _Username;
		private string _Domain;
		private bool _RelateSelectedOnly;
		private bool _CustomOnly;
		private string _Namespace;
		private string _ProjectName;
		private string _ServerName = "";
		private string _ServerPort = "";
		private string _HomeRealm = "";
		private bool _UseWindowsAuth;
		private string _EntitiesString;
		private string _SelectPrefixes = "";
		private bool _SplitFiles;
		private int _Depth;
		
		public bool UseSSL
		{
			get { return _UseSSL; }
			set
			{
				if (SetField(ref _UseSSL, value))
				{
					ReEvalReadOnly();
				}
			}
		}

		public bool UseIFD
		{
			get { return _UseIFD; }
			set
			{
				if (SetField(ref _UseIFD, value))
				{
					if (value)
					{
						UseOnline = false;
						UseOffice365 = false;
						UseSSL = true;
						UseWindowsAuth = false;
					}
					ReEvalReadOnly();
				}
			}
		}

		public bool UseOnline
		{
			get { return _UseOnline; }
			set
			{
				if (SetField(ref _UseOnline, value))
				{
					if (value)
					{
						UseIFD = false;
						UseOffice365 = true;
						UseSSL = true;
						UseWindowsAuth = false;
					}
					else
					{
						UseOffice365 = false;
					}
					ReEvalReadOnly();
				}
			}
		}

		public bool UseOffice365
		{
			get { return _UseOffice365; }
			set
			{
				if (SetField(ref _UseOffice365, value))
				{
					if (value)
					{
						UseIFD = false;
						UseOnline = true;
						UseSSL = true;
						UseWindowsAuth = false;
					}
					ReEvalReadOnly();
				}
			}
		}

		public string EntitiesToIncludeString
		{
			get
			{
				var sb = new StringBuilder();
				foreach (var value in _EntitiesSelected)
				{
					if (sb.Length != 0)
					{
						sb.Append(',');
					}
					sb.Append(value);
				}
				return sb.ToString();
			}
			set
			{
				var newList = new ObservableCollection<string>();
				var split = value.Split(',').Select(p => p.Trim()).ToList();
				foreach (var s in split)
				{
					newList.Add(s);
					if (!_EntityList.Contains(s))
					{
						_EntityList.Add(s);
					}
				}
				EntitiesSelected = newList;
				SetField(ref _EntitiesToIncludeString, value);
				OnPropertyChanged("EnableExclude");
			}
		}

		public string CrmOrg
		{
			get { return _CrmOrg; }
			set { SetField(ref _CrmOrg, value); }
		}

		public string Password
		{
			get { return _Password; }
			set { SetField(ref _Password, value); }
		}

		public string Username
		{
			get { return _Username; }
			set { SetField(ref _Username, value); }
		}

		public string Domain
		{
			get { return _Domain; }
			set { SetField(ref _Domain, value); }
		}

		public bool RelateSelectedOnly
		{
			get { return _RelateSelectedOnly; }
			set { SetField(ref _RelateSelectedOnly, value); }
		}

		public bool CustomOnly
		{
			get { return _CustomOnly; }
			set { SetField(ref _CustomOnly, value); }
		}

		public int Depth
		{
			get { return _Depth; }
			set { SetField(ref _Depth, value); }
		}

		public string Namespace
		{
			get { return _Namespace; }
			set { SetField(ref _Namespace, value); }
		}

		public string ProjectName
		{
			get { return _ProjectName; }
			set { SetField(ref _ProjectName, value); }
		}

		public string ServerName
		{
			get { return _ServerName; }
			set { SetField(ref _ServerName, value); }
		}

		public string ServerPort
		{
			get
			{
				if (UseOnline || UseOffice365)
				{
					return "";
				}
				return _ServerPort;
			}
			set { SetField(ref _ServerPort, value); }
		}

		public string HomeRealm
		{
			get { return _HomeRealm; }
			set { SetField(ref _HomeRealm, value); }
		}

		public bool UseWindowsAuth
		{
			get { return _UseWindowsAuth; }
			set
			{
				SetField(ref _UseWindowsAuth, value);
				ReEvalReadOnly();
			}
		}
		
		public string EntitiesString
		{
			get
			{
				var sb = new StringBuilder();

				foreach (var value in _EntityList)
				{
					if (sb.Length != 0)
					{
						sb.Append(',');
					}
					sb.Append(value);
				}

				_EntitiesString = sb.ToString();

				return _EntitiesString;
			}
			set
			{
				var split = value.Split(',').Select(p => p.Trim()).ToList();

				foreach (var s in split.Where(s => !_EntityList.Contains(s)))
				{
					_EntityList.Add(s);
				}

				_EntitiesString = value;
			}
		}
		
		public string SelectPrefixes
		{
			get
			{
				return _SelectPrefixes;
			}
			set
			{
				_SelectPrefixes = value;
			}
		}
		
		public bool SplitFiles
		{
			get
			{
				return _SplitFiles;
			}
			set
			{
				_SplitFiles = value;
			}
		}

		#region Non serialisable

		[NonSerialized] private string _CrmSdkUrl;
		[NonSerialized] private string _Template;
		[NonSerialized] private string _T4Path;
		[NonSerialized] private string _OutputPath;
		[NonSerialized] private string _Folder = "";
		[NonSerialized] private bool _NewTemplate;
		[NonSerialized] private ObservableCollection<string> _OnLineServers = new ObservableCollection<string>();
		[NonSerialized] private ObservableCollection<string> _OrgList = new ObservableCollection<string>();
		[NonSerialized] private ObservableCollection<string> _TemplateList = new ObservableCollection<string>();
		[NonSerialized] public ObservableCollection<string> _EntitiesSelected;
		[NonSerialized] public ObservableCollection<string> _EntityList;

		public string CrmSdkUrl
		{
			get { return _CrmSdkUrl; }
			set { SetField(ref _CrmSdkUrl, value); }
		}

		public string Template
		{
			get { return _Template; }
			set
			{
				SetField(ref _Template, value);
				NewTemplate = !File.Exists(Path.Combine(_Folder, _Template));
			}
		}

		public string T4Path
		{
			get { return _T4Path; }
			set { SetField(ref _T4Path, value); }
		}

		public string OutputPath
		{
			get { return _OutputPath; }
			set { SetField(ref _OutputPath, value); }
		}

		public string Folder
		{
			get { return _Folder; }
			set { SetField(ref _Folder, value); }
		}

		public bool NewTemplate
		{
			get { return _NewTemplate; }
			set { SetField(ref _NewTemplate, value); }
		}


		public ObservableCollection<string> OnLineServers
		{
			get { return _OnLineServers; }
			set { SetField(ref _OnLineServers, value); }
		}


		public ObservableCollection<string> OrgList
		{
			get { return _OrgList; }
			set { SetField(ref _OrgList, value); }
		}


		public ObservableCollection<string> TemplateList
		{
			get { return _TemplateList; }
			set { SetField(ref _TemplateList, value); }
		}

		public ObservableCollection<string> EntityList
		{
			get { return _EntityList; }
			set { SetField(ref _EntityList, value); }
		}

		public ObservableCollection<string> EntitiesSelected
		{
			get { return _EntitiesSelected; }
			set { SetField(ref _EntitiesSelected, value); }
		}

		public IOrganizationService CrmConnection { get; set; }

		public bool Dirty { get; set; }

		#endregion

		#region Read Only Properties

		private void ReEvalReadOnly()
		{
			OnPropertyChanged("NeedServer");
			OnPropertyChanged("NeedOnlineServer");
			OnPropertyChanged("NeedServerPort");
			OnPropertyChanged("NeedHomeRealm");
			OnPropertyChanged("NeedCredentials");
			OnPropertyChanged("CanUseWindowsAuth");
			OnPropertyChanged("CanUseSSL");
		}

		public bool NeedServer
		{
			get { return !(UseOnline || UseOffice365); }
		}

		public bool NeedOnlineServer
		{
			get { return (UseOnline || UseOffice365); }
		}

		public bool NeedServerPort
		{
			get { return !(UseOffice365 || UseOnline); }
		}

		public bool NeedHomeRealm
		{
			get { return !(UseIFD || UseOffice365 || UseOnline); }
		}

		public bool NeedCredentials
		{
			get { return !UseWindowsAuth; }
		}

		public bool CanUseWindowsAuth
		{
			get { return !(UseIFD || UseOnline || UseOffice365); }
		}

		public bool CanUseSSL
		{
			get { return !(UseOnline || UseOffice365 || UseIFD); }
		}

		#endregion

		#region Conntection Strings

		public AuthenticationProviderType AuthType
		{
			get
			{
				if (UseIFD)
				{
					return AuthenticationProviderType.Federation;
				}
				else if (UseOffice365)
				{
					return AuthenticationProviderType.OnlineFederation;
				}
				else if (UseOnline)
				{
					return AuthenticationProviderType.LiveId;
				}

				return AuthenticationProviderType.ActiveDirectory;
			}
		}

		public string GetDiscoveryCrmConnectionString()
		{
			var connectionString = string.Format("Url={0}://{1}:{2};",
				UseSSL ? "https" : "http",
				UseIFD ? ServerName : UseOffice365 ? "disco." + ServerName : UseOnline ? "dev." + ServerName : ServerName,
				ServerPort.Length == 0 ? (UseSSL ? 443 : 80) : int.Parse(ServerPort));

			if (!UseWindowsAuth)
			{
				if (!UseIFD)
				{
					if (!string.IsNullOrEmpty(Domain))
					{
						connectionString += string.Format("Domain={0};", Domain);
					}
				}

				var sUsername = Username;
				if (UseIFD)
				{
					if (!string.IsNullOrEmpty(Domain))
					{
						sUsername = string.Format("{0}\\{1}", Domain, Username);
					}
				}

				connectionString += string.Format("Username={0};Password={1};", sUsername, Password);
			}

			if (UseOnline && !UseOffice365)
			{
				ClientCredentials deviceCredentials;

				do
				{
					deviceCredentials = DeviceIdManager.LoadDeviceCredentials() ??
					                    DeviceIdManager.RegisterDevice();
				} while (deviceCredentials.UserName.Password.Contains(";")
				         || deviceCredentials.UserName.Password.Contains("=")
				         || deviceCredentials.UserName.Password.Contains(" ")
				         || deviceCredentials.UserName.UserName.Contains(";")
				         || deviceCredentials.UserName.UserName.Contains("=")
				         || deviceCredentials.UserName.UserName.Contains(" "));

				connectionString += string.Format("DeviceID={0};DevicePassword={1};",
					deviceCredentials.UserName.UserName,
					deviceCredentials.UserName.Password);
			}

			if (UseIFD && !string.IsNullOrEmpty(HomeRealm))
			{
				connectionString += string.Format("HomeRealmUri={0};", HomeRealm);
			}

			return connectionString;
		}


		public string GetOrganizationCrmConnectionString()
		{
			var currentServerName = string.Empty;

			var orgDetails = ConnectionHelper.GetOrganizationDetails(this);
			if (UseOffice365 || UseOnline)
			{
				currentServerName = string.Format("{0}.{1}", orgDetails.UrlName, ServerName);
			}
			else if (UseIFD)
			{
				var serverNameParts = ServerName.Split('.');

				serverNameParts[0] = orgDetails.UrlName;


				currentServerName = string.Format("{0}:{1}",
					string.Join(".", serverNameParts),
					ServerPort.Length == 0 ? (UseSSL ? 443 : 80) : int.Parse(ServerPort));
			}
			else
			{
				currentServerName = string.Format("{0}:{1}/{2}",
					ServerName,
					ServerPort.Length == 0 ? (UseSSL ? 443 : 80) : int.Parse(ServerPort),
					CrmOrg);
			}

			//var connectionString = string.Format("Url={0}://{1};",
			//                                     UseSSL ? "https" : "http",
			//                                     currentServerName);

			var connectionString = string.Format("Url={0};",
				orgDetails.Endpoints[EndpointType.OrganizationService]/*.Replace("/XRMServices/2011/Organization.svc", "")*/);

			if (!UseWindowsAuth)
			{
				if (!UseIFD)
				{
					if (!string.IsNullOrEmpty(Domain))
					{
						connectionString += string.Format("Domain={0};", Domain);
					}
				}

				var username = Username;
				if (UseIFD)
				{
					if (!string.IsNullOrEmpty(Domain))
					{
						username = string.Format("{0}\\{1}", Domain, Username);
					}
				}

				connectionString += string.Format("Username={0};Password={1};", username, Password);
			}

			//if (UseOnline)
			//{
			//	ClientCredentials deviceCredentials;

			//	do
			//	{
			//		deviceCredentials = DeviceIdManager.LoadDeviceCredentials() ??
			//							DeviceIdManager.RegisterDevice();
			//	} while (deviceCredentials.UserName.Password.Contains(";")
			//			 || deviceCredentials.UserName.Password.Contains("=")
			//			 || deviceCredentials.UserName.Password.Contains(" ")
			//			 || deviceCredentials.UserName.UserName.Contains(";")
			//			 || deviceCredentials.UserName.UserName.Contains("=")
			//			 || deviceCredentials.UserName.UserName.Contains(" "));

			//	connectionString += string.Format("DeviceID={0};DevicePassword={1};",
			//		deviceCredentials.UserName.UserName,
			//		deviceCredentials.UserName.Password);
			//}

			if (UseIFD && !string.IsNullOrEmpty(HomeRealm))
			{
				connectionString += string.Format("HomeRealmUri={0};", HomeRealm);
			}

			//append timeout in seconds to connectionstring
			//connectionString += string.Format("Timeout={0};", Timeout.ToString(@"hh\:mm\:ss"));
			return connectionString;
		}

		#endregion
	}
}
