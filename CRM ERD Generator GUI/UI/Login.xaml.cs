#region Imports

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CRM_ERD_Generator_GUI.Helpers;
using CRM_ERD_Generator_GUI.Model;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Application = System.Windows.Forms.Application;
using MultiSelectComboBoxClass = CrmCodeGenerator.Controls.MultiSelectComboBox;

#endregion

namespace CRM_ERD_Generator_GUI.UI
{
	/// <summary>
	///     Interaction logic for Login.xaml
	/// </summary>
	public partial class Login
	{
		private readonly Settings settings;
		private DiagramBuilder builder;

		private EntityMetadata[] allEntities;

		public Login()
		{
			Assembly.Load("Xceed.Wpf.Toolkit");

			Status.Login = this;

			InitializeComponent();

			Configuration.FileName = "settings";
			settings = new Settings(Configuration.LoadConfigs());

			txtPassword.Password = settings.Password; // PasswordBox doesn't allow 2 way binding
			DataContext = settings;
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			IntSpinnerDepth.Value = settings.Depth;

			if (settings.OrgList.Contains(settings.CrmOrg) == false)
			{
				settings.OrgList.Add(settings.CrmOrg);
			}

			Organization.SelectedItem = settings.CrmOrg;
		}

		private void Cancel_Click(object sender, RoutedEventArgs e)
		{
			Close();
		}

		private void RefreshOrgs(object sender, RoutedEventArgs e)
		{
			settings.Password = ((PasswordBox) ((Button) sender).CommandParameter).Password;
			// PasswordBox doesn't allow 2 way binding, so we have to manually read it
			UpdateStatus("Refreshing organisations ...", true);
			try
			{
				//var orgs = QuickConnection.GetOrganizations(settings.CrmSdkUrl, settings.Domain, settings.Username, settings.Password);
				//var newOrgs = new ObservableCollection<String>(orgs);
				//settings.OrgList = newOrgs;

				var newOrgs = ConnectionHelper.GetOrgList(settings);
				settings.OrgList = newOrgs;
				UpdateStatus("Finished refreshing organisations.", false);
			}
			catch (Exception ex)
			{
				var error = "[ERROR] " + ex.Message
				            + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
				UpdateStatus(error, false);
				UpdateStatus("Unable to refresh organizations, check connection information", false);
			}

			UpdateStatus("", false);
		}

		private void EntitiesRefresh_Click(object sender, RoutedEventArgs events)
		{
			settings.Password = ((PasswordBox) ((Button) sender).CommandParameter).Password;
			// PasswordBox doesn't allow 2 way binding, so we have to manually read it

			UpdateStatus("Refreshing entities...", true);

			RefreshEntityList();

			UpdateStatus("Finished refreshing entities.", false);
		}

		private void RefreshEntityList(bool refresh = true)
		{
			if (refresh)
			{
				Update_AllEntities();
			}

			if (allEntities == null)
			{
				return;
			}

			var entities = allEntities.Where(e => !EntityHelper.NonStandard.Contains(e.LogicalName));

			var origSelection = settings.EntitiesToIncludeString;
			var newList = new ObservableCollection<string>();
			foreach (var entity in entities.OrderBy(e => e.LogicalName))
			{
				newList.Add(entity.LogicalName);
			}

			settings.EntityList = newList;
			settings.EntitiesToIncludeString = origSelection;
		}

		private void Update_AllEntities()
		{
			try
			{
				// TODO REMOVE THIS var connection = QuickConnection.Connect(settings.CrmSdkUrl, settings.Domain, settings.Username, settings.Password, settings.CrmOrg);
				var connString = CrmConnection.Parse(settings.GetOrganizationCrmConnectionString());
				var connection = new OrganizationService(connString);

				var entityProperties = new MetadataPropertiesExpression
				                       {
					                       AllProperties = false
				                       };
				entityProperties.PropertyNames.AddRange("LogicalName");
				var entityQueryExpression = new EntityQueryExpression
				                            {
					                            Properties = entityProperties
				                            };
				var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
				                                     {
					                                     Query = entityQueryExpression,
					                                     ClientVersionStamp = null
				                                     };

				allEntities =
					((RetrieveMetadataChangesResponse) connection.Execute(retrieveMetadataChangesRequest)).EntityMetadata.ToArray();
			}
			catch (Exception ex)
			{
				var error = "[ERROR] " + ex.Message
				            + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
				UpdateStatus(error, false);
				UpdateStatus("Unable to refresh entities, check connection information", false);
			}
		}

		private void IncludeNonStandardEntities_Click(object sender, RoutedEventArgs e)
		{
			if (allEntities != null)
			{
				RefreshEntityList();
				// if we don't have the entire list of entities don't do anything (eg if they havn't entered a username & password)
			}
		}

		private void Logon_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (IntSpinnerDepth.Value.HasValue)
				{
					settings.Depth = IntSpinnerDepth.Value.Value;
				}

				builder = new DiagramBuilder(settings);

				settings.Password = ((PasswordBox)((Button)sender).CommandParameter).Password;
				// PasswordBox doesn't allow 2 way binding, so we have to manually read it
				settings.Dirty = true;

				//  TODO Because the EntitiesSelected is a collection, the Settings class can't see when an item is added or removed.  when I have more time need to get the observable to bubble up.

				// save user's 'split files'
				if (CheckBoxSplitFiles.IsChecked.HasValue)
				{
					settings.SplitFiles = CheckBoxSplitFiles.IsChecked.Value;
				}

				Configuration.SaveConfigs(settings.SerializedSettings);

				UpdateStatus("Generating entity diagrams, this might take a while depending on CRM server/connection speed ... ", true);

				// check user's 'split files'
				if (settings.SplitFiles)
				{
					UpdateStatus("Generator will split entity diagrams into separate pages.", true);
				}

				RegisterMapperEvents();

				Task.Factory.StartNew(
					() =>
					{
						try
						{
							builder.Generate();
						}
						catch (Exception ex)
						{
							var error = "\n\n[ERROR] " + ex.Message
							            + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
							UpdateStatus(error, false);
							UpdateStatus(ex.StackTrace, false);
							UpdateStatus("Unable to map entities, see error above.", false);
						}
					});
			}
			catch (Exception ex)
			{
				var error = "\n\n[ERROR] " + ex.Message
				            + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
				UpdateStatus(error, false);
				UpdateStatus(ex.StackTrace, false);
				UpdateStatus("Unable to map entities, see error above.", false);
			}
		}

		private void RegisterMapperEvents()
		{
			builder.PropertyChanged
				+= (o, args) =>
				   {
					   try
					   {
						   switch (args.PropertyName)
						   {
							   case "LogMessage":
								   lock (builder.LoggingLock)
								   {
									   UpdateStatus(builder.LogMessage, true);
								   }
								   break;

							   case "Cancelled":
								   if (builder.Cancelled)
								   {
									   UpdateStatus("\n\nCancelled generator!", false);
								   }
								   break;

							   case "Error":
								   UpdateStatus("\n\nGenerator produced an error! => " + builder.Error.Message, false);
								   break;

							   case "Done":
								   Configuration.SaveConfigs(settings.SerializedSettings);
								   UpdateStatus("\n\nGenerator finished!", false);
								   break;
						   }
					   }
					   catch
					   {
						   // ignored
					   }
				   };
		}

		// credit: http://stackoverflow.com/questions/636383/how-can-i-find-wpf-controls-by-name-or-type (CrimsonX)
		private void SetEnabledChildren(DependencyObject obj, bool enabled, params string[] exclusions)
		{
			var childrenCount = VisualTreeHelper.GetChildrenCount(obj);

			if (childrenCount > 0)
			{
				for (var i = 0; i < childrenCount; i++)
				{
					var child = VisualTreeHelper.GetChild(obj, i);
					SetEnabledChildren(child, enabled, exclusions);
				}
			}

			if (obj is TextBox || obj is CheckBox || obj is Button
			    || obj is MultiSelectComboBoxClass || obj is PasswordBox || obj is ComboBox)
			{
				var frameworkElement = (FrameworkElement) obj;

				if (exclusions.All(exclusion => frameworkElement.Name != exclusion))
				{
					Dispatcher.Invoke(() => frameworkElement.IsEnabled = enabled);
				}
			}
		}

		internal void UpdateStatus(string message, bool working)
		{
			Dispatcher.Invoke(() => SetEnabledChildren(Inputs, !working, "ButtonCancel"));

			if (!string.IsNullOrWhiteSpace(message))
			{
				Status.Update(message);
			}

			Application.DoEvents();
			// Needed to allow the output window to update (also allows the cursor wait and form disable to show up)
		}

		private void ButtonSelectCustom_OnClick_Click(object sender, RoutedEventArgs e)
		{
			// get all prefixes
			var prefixes = settings.SelectPrefixes.Split(',').Select(prefix => prefix + "_");
			// get custom entity names from the fetched list
			var customEntities = settings.EntityList.Where(entity => prefixes.Any(entity.StartsWith));
			// get only unselected entities
			var newEntities = customEntities.Where(entity => !settings.EntitiesSelected.Contains(entity)).ToList();

			if (newEntities.Any())
			{
				// add the unselected entities to the list of inclusions
				settings.EntitiesToIncludeString += "," + newEntities.Aggregate((entity1, entity2) => entity1 + "," + entity2);
			}
		}

		private void ButtonCancel_Click(object sender, RoutedEventArgs e)
		{
			builder.Cancel = true;
		}
	}


	#region TextBox scroll behaviour

	// credit: http://stackoverflow.com/questions/10097417/how-do-i-create-an-autoscrolling-textbox
	public class TextBoxBehaviour
	{
		static readonly Dictionary<TextBox, Capture> associations = new Dictionary<TextBox, Capture>();

		public static bool GetScrollOnTextChanged(DependencyObject dependencyObject)
		{
			return (bool)dependencyObject.GetValue(ScrollOnTextChangedProperty);
		}

		public static void SetScrollOnTextChanged(DependencyObject dependencyObject, bool value)
		{
			dependencyObject.SetValue(ScrollOnTextChangedProperty, value);
		}

		public static readonly DependencyProperty ScrollOnTextChangedProperty =
            DependencyProperty.RegisterAttached("ScrollOnTextChanged", typeof(bool), typeof(TextBoxBehaviour), new UIPropertyMetadata(false, OnScrollOnTextChanged));

		static void OnScrollOnTextChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
		{
			var textBox = dependencyObject as TextBox;
			if (textBox == null)
			{
				return;
			}
			bool oldValue = (bool)e.OldValue, newValue = (bool)e.NewValue;
			if (newValue == oldValue)
			{
				return;
			}
			if (newValue)
			{
				textBox.Loaded += TextBoxLoaded;
				textBox.Unloaded += TextBoxUnloaded;
			}
			else
			{
				textBox.Loaded -= TextBoxLoaded;
				textBox.Unloaded -= TextBoxUnloaded;
				if (associations.ContainsKey(textBox))
				{
					associations[textBox].Dispose();
				}
			}
		}

		static void TextBoxUnloaded(object sender, RoutedEventArgs routedEventArgs)
		{
			var textBox = (TextBox)sender;
			associations[textBox].Dispose();
			textBox.Unloaded -= TextBoxUnloaded;
		}

		static void TextBoxLoaded(object sender, RoutedEventArgs routedEventArgs)
		{
			var textBox = (TextBox)sender;
			textBox.Loaded -= TextBoxLoaded;
			associations[textBox] = new Capture(textBox);
		}

		class Capture : IDisposable
		{
			private TextBox TextBox { get; set; }

			public Capture(TextBox textBox)
			{
				TextBox = textBox;
				TextBox.TextChanged += OnTextBoxOnTextChanged;
			}

			private void OnTextBoxOnTextChanged(object sender, TextChangedEventArgs args)
			{
				TextBox.ScrollToEnd();
			}

			public void Dispose()
			{
				TextBox.TextChanged -= OnTextBoxOnTextChanged;
			}
		}

	}

	#endregion
}
