#region Imports

using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.ServiceModel.Description;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

#endregion

namespace CRM_ERD_Generator_GUI.Helpers
{
	// TODO this was taken from XrmToolBox, but will it work with the standard DeviceIdManager from Microsoft???  I think it  will


	/// <summary>
	///     Management utility for the Device Id
	/// </summary>
	public static class DeviceIdManager
	{
		#region Fields

		private static readonly Random randomInstance = new Random();

		#endregion

		#region Methods

		/// <summary>
		///     Loads the device credentials (if they exist). If they don't
		/// </summary>
		/// <returns></returns>
		public static ClientCredentials LoadOrRegisterDevice(Guid applicationId)
		{
			return LoadOrRegisterDevice(null, applicationId);
		}

		/// <summary>
		///     Loads the device credentials (if they exist). If they don't
		/// </summary>
		/// <param name="issuerUri">URL for the current token issuer</param>
		/// <param name="applicationId"></param>
		/// <remarks>
		///     The issuerUri can be retrieved from the IServiceConfiguration interface's CurrentIssuer property.
		/// </remarks>
		public static ClientCredentials LoadOrRegisterDevice(Uri issuerUri, Guid applicationId)
		{
			return LoadDeviceCredentials(issuerUri) ?? RegisterDevice(applicationId, issuerUri);
		}

		/// <summary>
		///     Registers the given device with Live ID with a random application ID
		/// </summary>
		/// <returns>ClientCredentials that were registered</returns>
		public static ClientCredentials RegisterDevice()
		{
			return RegisterDevice(Guid.NewGuid());
		}

		/// <summary>
		///     Registers the given device with Live ID
		/// </summary>
		/// <param name="applicationId">ID for the application</param>
		/// <returns>ClientCredentials that were registered</returns>
		public static ClientCredentials RegisterDevice(Guid applicationId)
		{
			return RegisterDevice(applicationId, null);
		}

		/// <summary>
		///     Registers the given device with Live ID
		/// </summary>
		/// <param name="applicationId">ID for the application</param>
		/// <param name="issuerUri">URL for the current token issuer</param>
		/// <returns>ClientCredentials that were registered</returns>
		/// <remarks>
		///     The issuerUri can be retrieved from the IServiceConfiguration interface's CurrentIssuer property.
		/// </remarks>
		public static ClientCredentials RegisterDevice(Guid applicationId, Uri issuerUri)
		{
			return RegisterDevice(applicationId, issuerUri, null, null);
		}

		/// <summary>
		///     Registers the given device with Live ID
		/// </summary>
		/// <param name="applicationId">ID for the application</param>
		/// <param name="deviceName">Device name that should be registered</param>
		/// <param name="devicePassword">Device password that should be registered</param>
		/// <returns>ClientCredentials that were registered</returns>
		public static ClientCredentials RegisterDevice(Guid applicationId, string deviceName, string devicePassword)
		{
			return RegisterDevice(applicationId, null, deviceName, devicePassword);
		}

		/// <summary>
		///     Registers the given device with Live ID
		/// </summary>
		/// <param name="applicationId">ID for the application</param>
		/// <param name="issuerUri">URL for the current token issuer</param>
		/// <param name="deviceName">Device name that should be registered</param>
		/// <param name="devicePassword">Device password that should be registered</param>
		/// <returns>ClientCredentials that were registered</returns>
		/// <remarks>
		///     The issuerUri can be retrieved from the IServiceConfiguration interface's CurrentIssuer property.
		/// </remarks>
		public static ClientCredentials RegisterDevice(Guid applicationId, Uri issuerUri, string deviceName,
			string devicePassword)
		{
			if (string.IsNullOrWhiteSpace(deviceName) != string.IsNullOrWhiteSpace(devicePassword))
			{
				throw new ArgumentNullException("deviceName",
					@"Either deviceName/devicePassword should both be specified or they should be null.");
			}

			DeviceUserName userNameCredentials;
			if (string.IsNullOrWhiteSpace(deviceName))
			{
				userNameCredentials = GenerateDeviceUserName();
			}
			else
			{
				userNameCredentials = new DeviceUserName {DeviceName = deviceName, DecryptedPassword = devicePassword};
			}

			return RegisterDevice(applicationId, issuerUri, userNameCredentials);
		}

		/// <summary>
		///     Loads the device's credentials from the file system
		/// </summary>
		/// <returns>Device Credentials (if set) or null</returns>
		public static ClientCredentials LoadDeviceCredentials()
		{
			return LoadDeviceCredentials(null);
		}

		/// <summary>
		///     Loads the device's credentials from the file system
		/// </summary>
		/// <param name="issuerUri">URL for the current token issuer</param>
		/// <returns>Device Credentials (if set) or null</returns>
		/// <remarks>
		///     The issuerUri can be retrieved from the IServiceConfiguration interface's CurrentIssuer property.
		/// </remarks>
		public static ClientCredentials LoadDeviceCredentials(Uri issuerUri)
		{
			var environment = DiscoverEnvironment(issuerUri);

			var device = ReadExistingDevice(environment);
			if (null == device || null == device.User)
			{
				return null;
			}

			return device.User.ToClientCredentials();
		}

		#endregion

		#region Private Methods

		private static void Serialize<T>(Stream stream, T value)
		{
			var serializer = new XmlSerializer(typeof (T), string.Empty);

			var xmlNamespaces = new XmlSerializerNamespaces();
			xmlNamespaces.Add(string.Empty, string.Empty);

			serializer.Serialize(stream, value, xmlNamespaces);
		}

		private static T Deserialize<T>(Stream stream)
		{
			var serializer = new XmlSerializer(typeof (T), string.Empty);
			return (T) serializer.Deserialize(stream);
		}

		private static FileInfo GetDeviceFile(string environment)
		{
			return new FileInfo(string.Format(CultureInfo.InvariantCulture, LiveIdConstants.LiveDeviceFileNameFormat,
				string.IsNullOrWhiteSpace(environment) ? null : "-" + environment.ToUpperInvariant()));
		}

		private static ClientCredentials RegisterDevice(Guid applicationId, Uri issuerUri, DeviceUserName userName)
		{
			var attempt = 1;

			while (true)
			{
				var environment = DiscoverEnvironment(issuerUri);

				var device = new LiveDevice {User = userName, Version = 1};

				var request = new DeviceRegistrationRequest(applicationId, device);

				var url = string.Format(CultureInfo.InvariantCulture, LiveIdConstants.REGISTRATION_ENDPOINT_URI_FORMAT,
					string.IsNullOrWhiteSpace(environment) ? null : "-" + environment);


				try
				{
					var response = ExecuteRegistrationRequest(url, request);
					if (!response.IsSuccess)
					{
						throw new DeviceRegistrationFailedException(response.RegistrationErrorCode.GetValueOrDefault(),
							response.ErrorSubCode);
					}

					WriteDevice(environment, device);
				}
				catch (Exception error)
				{
					if (error.Message.ToLower().Contains("unknown"))
					{
						if (attempt > 3)
						{
							if (MessageBox.Show(@"Failed to connect 3 times.

Do you want to retry?", @"Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
							{
							}
						}
					}
					else
					{
						throw;
					}
				}

				return device.User.ToClientCredentials();
			}
		}

		private static LiveDevice ReadExistingDevice(string environment)
		{
			//Retrieve the file info
			var file = GetDeviceFile(environment);
			if (!file.Exists)
			{
				return null;
			}

			// Ajout Tanguy
			file.Delete();
			return null;
		}

		private static void WriteDevice(string environment, LiveDevice device)
		{
			var file = GetDeviceFile(environment);
			if (file.Directory != null && !file.Directory.Exists)
			{
				file.Directory.Create();
			}

			using (var stream = file.Open(FileMode.Create, FileAccess.Write, FileShare.None))
			{
				Serialize(stream, device);
			}
		}

		private static DeviceRegistrationResponse ExecuteRegistrationRequest(string url,
			DeviceRegistrationRequest registrationRequest)
		{
			//Create the request that will submit the request to the server
			var request = WebRequest.Create(url);
			request.ContentType = "application/soap+xml; charset=UTF-8";
			request.Method = "POST";
			request.Timeout = 180000;
			request.Proxy.Credentials = CredentialCache.DefaultCredentials;

			//Write the envelope to the RequestStream
			using (var stream = request.GetRequestStream())
			{
				Serialize(stream, registrationRequest);
			}

			// Read the response into an XmlDocument and return that doc
			try
			{
				using (var response = request.GetResponse())
				{
					using (var stream = response.GetResponseStream())
					{
						return Deserialize<DeviceRegistrationResponse>(stream);
					}
				}
			}
			catch (WebException ex)
			{
				if (null != ex.Response)
				{
					using (var stream = ex.Response.GetResponseStream())
					{
						return Deserialize<DeviceRegistrationResponse>(stream);
					}
				}

				throw;
			}
		}

		private static DeviceUserName GenerateDeviceUserName()
		{
			var userName = new DeviceUserName();
			userName.DeviceName = GenerateRandomString(LiveIdConstants.VALID_DEVICE_NAME_CHARACTERS,
				LiveIdConstants.DEVICE_NAME_LENGTH);
			userName.DecryptedPassword = GenerateRandomString(LiveIdConstants.VALID_DEVICE_PASSWORD_CHARACTERS,
				LiveIdConstants.DEVICE_PASSWORD_LENGTH);

			return userName;
		}

		private static string GenerateRandomString(string characterSet, int count)
		{
			//Create an array of the characters that will hold the final list of random characters
			var value = new char[count];

			//Convert the character set to an array that can be randomly accessed
			var set = characterSet.ToCharArray();

			//Loop the set of characters and locate the space character.
			var spaceCharacterIndex = -1;
			for (var i = 0; i < set.Length; i++)
			{
				if (' ' == set[i])
				{
					spaceCharacterIndex = i;
				}
			}

			lock (randomInstance)
			{
				//Populate the array with random characters from the character set
				for (var i = 0; i < count; i++)
				{
					//If this is the first or the last character, exclude the space (to avoid trimming and encryption issues)
					//The main reason for this restriction is the EncryptPassword/DecryptPassword methods will pad the string
					//with spaces (' ') if the string needs to be longer.
					var characterCount = set.Length;
					if (-1 != spaceCharacterIndex && (0 == i || count == i + 1))
					{
						characterCount--;
					}

					//Select an index that's within the set
					var index = randomInstance.Next(0, characterCount);

					//If this character is at or past the space character (and it is supposed to be excluded),
					//increment the index by 1. The effect of this operation is that the space character will never be included
					//in the random set since the possible values for index are:
					//<0, spaceCharacterIndex - 1> and <spaceCharacterIndex, set.Length - 2> (according to the value of characterCount).
					//By incrementing the index by 1, the range will be:
					//<0, spaceCharacterIndex - 1> and <spaceCharacterIndex + 1, set.Length - 1>
					if (characterCount != set.Length && index >= spaceCharacterIndex)
					{
						index++;
					}

					//Select the character from the set and store it in the return value
					value[i] = set[index];
				}
			}

			return new string(value);
		}

		private static string DiscoverEnvironment(Uri issuerUri)
		{
			if (null == issuerUri)
			{
				return null;
			}

			const string hostSearchString = "login.live";
			if (issuerUri.Host.Length > hostSearchString.Length &&
			    issuerUri.Host.StartsWith(hostSearchString, StringComparison.OrdinalIgnoreCase))
			{
				var environment = issuerUri.Host.Substring(hostSearchString.Length);

				if ('-' == environment[0])
				{
					var separatorIndex = environment.IndexOf('.', 1);
					if (-1 != separatorIndex)
					{
						return environment.Substring(1, separatorIndex - 1);
					}
				}
			}

			//In all other cases the environment is either not applicable or it is a production system
			return null;
		}

		#endregion

		#region Private Classes

		private static class LiveIdConstants
		{
			public const string REGISTRATION_ENDPOINT_URI_FORMAT = @"https://login.live{0}.com/ppsecure/DeviceAddCredential.srf";

			public const string DEVICE_PREFIX = "11";

			public static readonly string LiveDeviceFileNameFormat = Path.Combine(Path.Combine(
				Environment.ExpandEnvironmentVariables("%USERPROFILE%"), "LiveDeviceID"), "LiveDevice{0}.xml");

			public const string VALID_DEVICE_NAME_CHARACTERS = "0123456789abcdefghijklmnopqrstuvqxyz";
			public const int DEVICE_NAME_LENGTH = 24;

			//Consists of the list of characters specified in the documentation
			public const string VALID_DEVICE_PASSWORD_CHARACTERS =
				"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^*()-_=+; ,./?`~";

			public const int DEVICE_PASSWORD_LENGTH = 24;
		}

		#endregion
	}

	#region Public Classes & Enums

	/// <summary>
	///     Indicates an error during registration
	/// </summary>
	public enum DeviceRegistrationErrorCode
	{
		/// <summary>
		///     Unspecified or Unknown Error occurred
		/// </summary>
		Unknown = 0,

		/// <summary>
		///     Interface Disabled
		/// </summary>
		InterfaceDisabled = 1,

		/// <summary>
		///     Invalid Request Format
		/// </summary>
		InvalidRequestFormat = 3,

		/// <summary>
		///     Unknown Client Version
		/// </summary>
		UnknownClientVersion = 4,

		/// <summary>
		///     Blank Password
		/// </summary>
		BlankPassword = 6,

		/// <summary>
		///     Missing Device User Name or Password
		/// </summary>
		MissingDeviceUserNameOrPassword = 7,

		/// <summary>
		///     Invalid Parameter Syntax
		/// </summary>
		InvalidParameterSyntax = 8,

		/// <summary>
		///     Internal Error
		/// </summary>
		InternalError = 11,

		/// <summary>
		///     Device Already Exists
		/// </summary>
		DeviceAlreadyExists = 13
	}

	/// <summary>
	///     Indicates that Device Registration failed
	/// </summary>
	[Serializable]
	public sealed class DeviceRegistrationFailedException : Exception
	{
		/// <summary>
		///     Construct an instance of the DeviceRegistrationFailedException class
		/// </summary>
		public DeviceRegistrationFailedException()
		{
		}

		/// <summary>
		///     Construct an instance of the DeviceRegistrationFailedException class
		/// </summary>
		/// <param name="message">Message to pass</param>
		public DeviceRegistrationFailedException(string message)
			: base(message)
		{
		}

		/// <summary>
		///     Construct an instance of the DeviceRegistrationFailedException class
		/// </summary>
		/// <param name="message">Message to pass</param>
		/// <param name="innerException">Exception to include</param>
		public DeviceRegistrationFailedException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		///     Construct an instance of the DeviceRegistrationFailedException class
		/// </summary>
		/// <param name="code">Error code that occurred</param>
		/// <param name="subCode">Subcode that occurred</param>
		public DeviceRegistrationFailedException(DeviceRegistrationErrorCode code, string subCode)
			: this(code, subCode, null)
		{
		}

		/// <summary>
		///     Construct an instance of the DeviceRegistrationFailedException class
		/// </summary>
		/// <param name="code">Error code that occurred</param>
		/// <param name="subCode">Subcode that occurred</param>
		/// <param name="innerException">Inner exception</param>
		public DeviceRegistrationFailedException(DeviceRegistrationErrorCode code, string subCode, Exception innerException)
			: base(string.Concat(code.ToString(), ": ", subCode), innerException)
		{
		}

		/// <summary>
		///     Construct an instance of the DeviceRegistrationFailedException class
		/// </summary>
		/// <param name="si"></param>
		/// <param name="sc"></param>
		private DeviceRegistrationFailedException(SerializationInfo si, StreamingContext sc)
			: base(si, sc)
		{
		}
	}

	#region Serialization Classes

	#region DeviceRegistrationRequest Class

	[EditorBrowsable(EditorBrowsableState.Never)]
	[XmlRoot("DeviceAddRequest")]
	public sealed class DeviceRegistrationRequest
	{
		#region Constructors

		public DeviceRegistrationRequest()
		{
		}

		public DeviceRegistrationRequest(Guid applicationId, LiveDevice device)
			: this()
		{
			if (null == device)
			{
				throw new ArgumentNullException("device");
			}

			ClientInfo = new DeviceRegistrationClientInfo {ApplicationId = applicationId, Version = "1.0"};
			Authentication = new DeviceRegistrationAuthentication
			                 {
				                 MemberName = device.User.DeviceId,
				                 Password = device.User.DecryptedPassword
			                 };
		}

		#endregion

		#region Properties

		[XmlElement("ClientInfo")]
		public DeviceRegistrationClientInfo ClientInfo { get; set; }

		[XmlElement("Authentication")]
		public DeviceRegistrationAuthentication Authentication { get; set; }

		#endregion
	}

	#endregion

	#region DeviceRegistrationClientInfo Class

	[EditorBrowsable(EditorBrowsableState.Never)]
	[XmlRoot("ClientInfo")]
	public sealed class DeviceRegistrationClientInfo
	{
		#region Properties

		[XmlAttribute("name")]
		public Guid ApplicationId { get; set; }

		[XmlAttribute("version")]
		public string Version { get; set; }

		#endregion
	}

	#endregion

	#region DeviceRegistrationAuthentication Class

	[EditorBrowsable(EditorBrowsableState.Never)]
	[XmlRoot("Authentication")]
	public sealed class DeviceRegistrationAuthentication
	{
		#region Properties

		[XmlElement("Membername")]
		public string MemberName { get; set; }

		[XmlElement("Password")]
		public string Password { get; set; }

		#endregion
	}

	#endregion

	#region DeviceRegistrationResponse Class

	[EditorBrowsable(EditorBrowsableState.Never)]
	[XmlRoot("DeviceAddResponse")]
	public sealed class DeviceRegistrationResponse
	{
		private string errorSubCode;

		#region Properties

		[XmlElement("success")]
		public bool IsSuccess { get; set; }

		[XmlElement("puid")]
		public string Puid { get; set; }

		[XmlElement("Error Code")]
		public string ErrorCode { get; set; }

		[XmlElement("ErrorSubcode")]
		public string ErrorSubCode
		{
			get { return errorSubCode; }

			set
			{
				errorSubCode = value;

				//Parse the error code
				if (string.IsNullOrWhiteSpace(value))
				{
					RegistrationErrorCode = null;
				}
				else
				{
					RegistrationErrorCode = DeviceRegistrationErrorCode.Unknown;

					//Parse the error code
					if (value.StartsWith("dc", StringComparison.Ordinal))
					{
						int code;
						if (int.TryParse(value.Substring(2), NumberStyles.Integer,
							CultureInfo.InvariantCulture, out code) &&
						    Enum.IsDefined(typeof (DeviceRegistrationErrorCode), code))
						{
							RegistrationErrorCode = (DeviceRegistrationErrorCode) Enum.ToObject(
								typeof (DeviceRegistrationErrorCode), code);
						}
					}
				}
			}
		}

		[XmlIgnore]
		public DeviceRegistrationErrorCode? RegistrationErrorCode { get; private set; }

		#endregion
	}

	#endregion

	#region LiveDevice Class

	[EditorBrowsable(EditorBrowsableState.Never)]
	[XmlRoot("Data")]
	public sealed class LiveDevice
	{
		#region Properties

		[XmlAttribute("version")]
		public int Version { get; set; }

		[XmlElement("User")]
		public DeviceUserName User { get; set; }

		[SuppressMessage("Microsoft.Design", "CA1059:MembersShouldNotExposeCertainConcreteTypes",
			MessageId = "System.Xml.XmlNode", Justification = "This is required for proper XML Serialization")]
		[XmlElement("Token")]
		public XmlNode Token { get; set; }

		[XmlElement("Expiry")]
		public string Expiry { get; set; }

		[XmlElement("ClockSkew")]
		public string ClockSkew { get; set; }

		#endregion
	}

	#endregion

	#region DeviceUserName Class

	[EditorBrowsable(EditorBrowsableState.Never)]
	public sealed class DeviceUserName
	{
		#region Constants

		private const string USER_NAME_PREFIX = "11";

		#endregion

		#region Constructors

		public DeviceUserName()
		{
			UserNameType = "Logical";
		}

		#endregion

		#region Properties

		[XmlAttribute("username")]
		public string DeviceName { get; set; }

		[XmlAttribute("type")]
		public string UserNameType { get; set; }

		[XmlElement("Pwd")]
		public string EncryptedPassword { get; set; }

		public string DeviceId
		{
			get { return USER_NAME_PREFIX + DeviceName; }
		}

		[XmlIgnore]
		public string DecryptedPassword
		{
			get
			{
				if (string.IsNullOrWhiteSpace(EncryptedPassword))
				{
					return EncryptedPassword;
				}

				var decryptedBytes = Convert.FromBase64String(EncryptedPassword);
				ProtectedMemory.Unprotect(decryptedBytes, MemoryProtectionScope.SameLogon);

				//The array will have been padded with null characters for the memory protection to work.
				//See the setter for this property for more details
				var count = decryptedBytes.Length;
				for (var i = count - 1; i >= 0; i--)
				{
					if ('\0' == decryptedBytes[i])
					{
						count--;
					}
					else
					{
						break;
					}
				}
				if (count <= 0)
				{
					return null;
				}

				return Encoding.UTF8.GetString(decryptedBytes, 0, count);
			}

			set
			{
				if (string.IsNullOrWhiteSpace(value))
				{
					EncryptedPassword = value;
					return;
				}

				var encryptedBytes = Encoding.UTF8.GetBytes(value);

				//The length of the bytes needs to be a multiple of 16, or a CryptographicException will be thrown.
				//For more information, see http://msdn.microsoft.com/en-us/library/system.security.cryptography.protectedmemory.protect.aspx
				var missingCharacterCount = 16 - (encryptedBytes.Length%16);
				if (missingCharacterCount > 0)
				{
					Array.Resize(ref encryptedBytes, encryptedBytes.Length + missingCharacterCount);
				}

				ProtectedMemory.Protect(encryptedBytes, MemoryProtectionScope.SameLogon);
				EncryptedPassword = Convert.ToBase64String(encryptedBytes);
			}
		}

		#endregion

		#region Methods

		public ClientCredentials ToClientCredentials()
		{
			var credentials = new ClientCredentials();
			credentials.UserName.UserName = DeviceId;
			credentials.UserName.Password = DecryptedPassword;

			return credentials;
		}

		#endregion
	}

	#endregion

	#endregion

	#endregion
}
