#region Imports

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.ServiceModel;
using CRM_ERD_Generator_GUI.Model;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using VisioApi = Microsoft.Office.Interop.Visio;

#endregion

namespace CRM_ERD_Generator_GUI
{
	/// <summary>
	///     Create a Visio diagram detailing relationships between Microsoft CRM entities.
	///     First, this sample reads in all the entity names. It then creates a visio object for
	///     the entity and all of the entities related to the entity, and links them together.
	///     Finally,it saves the file to disk.
	/// </summary>
	public class DiagramBuilder
	{
		#region Class Level Members

		/// <summary>
		///     Stores the organization service proxy.
		/// </summary>
		public IOrganizationService Service;

		// Specify which language code to use in the sample. If you are using a language
		// other than US English, you will need to modify this value accordingly.
		// See http://msdn.microsoft.com/en-us/library/0h88fahh.aspx
		public const int LANGUAGE_CODE = 1033;

		private VisioApi.Application application;
		private VisioApi.Document document;
		private List<Guid> processedRelationships;

		private const double X_POS1 = 0;
		private const double Y_POS1 = 0;
		private const double X_POS2 = 1.75;
		private const double Y_POS2 = 0.6;

		private const double SHDW_PATTERN = 0;
		private const double BEGIN_ARROW_MANY = 29;
		private const double BEGIN_ARROW = 0;
		private const double END_ARROW = 29;
		private const double LINE_COLOR_MANY = 10;
		private const double LINE_COLOR = 8;
		private const double LINE_PATTERN_MANY = 2;
		private const double LINE_PATTERN = 1;
		private const string LINE_WEIGHT = "2pt";
		private const double ROUNDING = 0.0625;
		private const double HEIGHT = 0.5;
		private const short NAME_CHARACTER_SIZE = 12;
		private const short FONT_STYLE = 0;
		private const short VISIO_SECTION_OJBECT_INDEX = 1;
		private string versionName;

		// Excluded entities.
		// These entities exist in the metadata but are not to be drawn in the diagram.
		private readonly string[] excludedEntities =
		{
			"attributemap", "bulkimport", "businessunitmap",
			"commitment", "displaystringmap", "documentindex",
			"entitymap", "importconfig", "integrationstatus",
			"internaladdress", "privilegeobjecttypecodes",
			"roletemplate", "roletemplateprivileges",
			"statusmap", "stringmap", "stringmapbit"
		};

		// Excluded relationship list.
		// Those entity relationships that should not be included in the diagram.
		private readonly Hashtable excludedRelationsTable = new Hashtable();
		private readonly string[] excludedRelations = {"owningteam", "organizationid"};

		private readonly Queue<string> mainEntities = new Queue<string>();
		private readonly Queue<string> toBeDrawnEntities = new Queue<string>();
		private readonly List<string> pageDrawnEntities = new List<string>();
		private readonly List<string> fileDrawnEntities = new List<string>();

		private int depth;

		private readonly List<EntityMetadata> entitiesCache = new List<EntityMetadata>();

		public readonly object LoggingLock = new object();

		public Settings Settings { get; set; }

		private string logMessage;

		public string LogMessage
		{
			get { return logMessage; }
			set
			{
				logMessage = value;
				OnPropertyChanged();
			}
		}

		private Exception error;

		public Exception Error
		{
			get { return error; }
			set
			{
				error = value;
				OnPropertyChanged();
			}
		}

		private bool done;

		public bool Done
		{
			get { return done; }
			set
			{
				done = value;
				OnPropertyChanged();
			}
		}

		private bool cancel;

		public bool Cancel
		{
			get { return cancel; }
			set
			{
				cancel = value;
				OnPropertyChanged();
			}
		}

		private bool cancelled;

		public bool Cancelled
		{
			get { return cancelled; }
			set
			{
				cancelled = value;
				OnPropertyChanged();
			}
		}

		#endregion Class Level Members

		#region event handler

		protected void OnMessage(string message, string extendedMessage = "")
		{
			lock (LoggingLock)
			{
				LogMessage = message + (string.IsNullOrEmpty(extendedMessage) ? "" : " => " + extendedMessage);
			}
		}


		public event PropertyChangedEventHandler PropertyChanged;

		protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
		{
			var handler = PropertyChanged;
			if (handler != null)
			{
				handler(this, new PropertyChangedEventArgs(propertyName));
			}
		}

		#endregion

		public DiagramBuilder(Settings settings)
		{
			Settings = settings;
			Service = GetConnection();

			// Do the same for excluded relationships.
			foreach (var t in excludedRelations)
			{
				excludedRelationsTable.Add(t.GetHashCode(), t);
			}

			processedRelationships = new List<Guid>();
		}

		public void Generate()
		{
			var builder = this;

			try
			{
				// Load Visio and create a new document.
				// Not showing the UI increases rendering speed  
				builder.application = application = new VisioApi.Application
				                                    {
					                                    Visible = false,
					                                    ScreenUpdating = 1,
					                                    EventsEnabled = 0,
					                                    ShowChanges = false,
					                                    ShowStatusBar = 0,
					                                    AlertResponse = 0,
					                                    ShowMenus = 0,
					                                    ShowProgress = 0,
					                                    ShowToolbar = 0,
					                                    UndoEnabled = false,
				                                    };
				builder.versionName = application.Version;
				builder.document = document = application.Documents.Add(string.Empty);

				Settings.EntitiesSelected.ToList().ForEach(mainEntities.Enqueue);

				builder.processedRelationships = new List<Guid>();

				if (mainEntities.Count <= 0)
				{
					OnMessage("No entities passed as argument; fetching all entities ...");
					GetEntities().ToList().ForEach(entityName => mainEntities.Enqueue(entityName.LogicalName));
				}

				OnMessage(string.Format("Parsing {0} ...",
					mainEntities.Aggregate((entity1, entity2) => entity1 + ", " + entity2)));

				depth = Settings.Depth;

				// loop through entities to be drawn
				while (mainEntities.Count > 0 && depth >= 0)
				{
					if (Cancel)
					{
						break;
					}

					var tempQueue = new Queue<string>();

					mainEntities.ToList().ForEach(tempQueue.Enqueue);
					mainEntities.Clear();
					builder.Parse(tempQueue);

					depth--;
				}

				// save the diagram in the current directory without overwriting
				var filename = "CRM_ERD.vsd";
				var index = 1;
				while (File.Exists(filename))
				{
					filename = "CRM_ERD_" + (index < 10 ? "0" + index++ : index++.ToString()) + ".vsd";
				}
				document.SaveAs(Directory.GetCurrentDirectory() + "\\" + filename);
				OnMessage(string.Format("\n\nSaved diagrams to {0}.", Directory.GetCurrentDirectory() + "\\" + filename));

				if (!Cancel)
				{
					Done = true;
				}
				else
				{
					Cancelled = true;
				}
			}
			catch (FaultException<OrganizationServiceFault> ex)
			{
				OnMessage("\n\nThe application terminated with an error.");
				OnMessage(string.Format("Timestamp: {0}", ex.Detail.Timestamp));
				OnMessage(string.Format("Code: {0}", ex.Detail.ErrorCode));
				OnMessage(string.Format("Message: {0}", ex.Detail.Message));
				OnMessage(string.Format("Plugin Trace: {0}", ex.Detail.TraceText));
				OnMessage(string.Format(
					"Inner Fault: {0}",
					null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault"));

				Error = ex;
			}
			catch (TimeoutException ex)
			{
				OnMessage("\n\nThe application terminated with an error.");
				OnMessage(string.Format("Message: {0}", ex.Message));
				OnMessage(string.Format("Stack Trace: {0}", ex.StackTrace));
				OnMessage(string.Format("Inner Fault: {0}",
					string.IsNullOrEmpty(ex.InnerException.Message) ? "No Inner Fault" : ex.InnerException.Message));

				Error = ex;
			}
			catch (Exception ex)
			{
				OnMessage("\n\nThe application terminated with an error.");
				OnMessage(string.Format(ex.Message));

				// Display the details of the inner exception.
				if (ex.InnerException != null)
				{
					OnMessage(string.Format(ex.InnerException.Message));

					var fe
						= ex.InnerException
						  as FaultException<OrganizationServiceFault>;
					if (fe != null)
					{
						OnMessage(string.Format("Timestamp: {0}", fe.Detail.Timestamp));
						OnMessage(string.Format("Code: {0}", fe.Detail.ErrorCode));
						OnMessage(string.Format("Message: {0}", fe.Detail.Message));
						OnMessage(string.Format("Plugin Trace: {0}", fe.Detail.TraceText));
						OnMessage(string.Format(
							"Inner Fault: {0}",
							null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault"));
					}
				}

				Error = ex;
			}
				// Additional exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
				// SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.
			finally
			{
				// close the Visio application. 
				application.Quit();
			}
		}

		private void Parse(Queue<string> entities)
		{
			var totalCount = entities.Count;
			var currentEntityIndex = 0;

			while (entities.Count > 0)
			{
				if (Cancel)
				{
					return;
				}

				var entityTemp = entities.Dequeue();

				if (fileDrawnEntities.Contains(entityTemp))
				{
					continue;
				}

				fileDrawnEntities.Add(entityTemp);

				if (Settings.SplitFiles)
				{
					pageDrawnEntities.Clear();
					toBeDrawnEntities.Clear();
					processedRelationships.Clear();

					document.Pages.Add();
				}

				pageDrawnEntities.Add(entityTemp);

				OnMessage(string.Format("\n\nParsing entity \"{2}\" {0} / {1} ...", ++currentEntityIndex, totalCount, entityTemp));

				BuildDiagram(new[] { entityTemp }, entityTemp);

				while (toBeDrawnEntities.Count > 0)
				{
					if (Cancel)
					{
						return;
					}

					var entityTemp2 = toBeDrawnEntities.Dequeue();

					if (!fileDrawnEntities.Contains(entityTemp2) && !mainEntities.Contains(entityTemp2))
					{
						if (Settings.RelateSelectedOnly)
						{
							if (Settings.EntitiesSelected.Contains(entityTemp2))
							{
								mainEntities.Enqueue(entityTemp2);
							}
						}
						else
						{
							mainEntities.Enqueue(entityTemp2);
						}
					}
				}
			}
		}

		/// <summary>
		///     Create a new page in a Visio file showing all the direct entity relationships participated in
		///     by the passed-in array of entities.
		/// </summary>
		/// <param name="entities">Core entities for the diagram</param>
		/// <param name="pageTitle">Page title</param>
		private void BuildDiagram(IEnumerable<string> entities, string pageTitle)
		{
			// Get the default page of our new document
			var page = document.Pages[document.Pages.Count];
			page.Name = pageTitle;

			var entityNames = entities.ToArray();

			// Get the metadata for each passed-in entity, draw it, and draw its relationships.
			foreach (var entityName in entityNames)
			{
				if (Cancel)
				{
					return;
				}

				OnMessage(string.Format("Processing entity: {0} ...", entityName));

				var entity = GetEntityMetadata(entityName);

				// Create a Visio rectangle shape.
				VisioApi.Shape rect = null;

				try
				{
					// There is no "Get Try", so we have to rely on an exception to tell us it does not exists
					// We have to skip some entities because they may have already been added by relationships of another entity
					rect = page.Shapes.ItemU[entity.SchemaName];
				}
				catch (COMException)
				{
					var label
						= entity.DisplayName.UserLocalizedLabel == null
							  ? ""
							  : entity.DisplayName.UserLocalizedLabel.Label + "\n";
					if (entity.OwnershipType != null)
					{
						rect = DrawEntityRectangle(
							page, label + "(" + entity.SchemaName + ")",
							entity.SchemaName,
							entity.OwnershipType.Value);
					}
					OnMessage(string.Format("Finished drawing entity: {0}.", entityName));
				}

				if (rect == null)
				{
					continue;
				}

				CacheEntities(entity);

				if (Cancel)
				{
					return;
				}

				// Draw all relationships TO this entity.
				DrawRelationships(entity, rect, entity.ManyToManyRelationships, false);
				OnMessage("Finished drawing M-N relationships.");
				DrawRelationships(entity, rect, entity.ManyToOneRelationships, false);
				OnMessage("Finished drawing N-1 relationships.");

				// Draw all relationshipos FROM this entity
				DrawRelationships(entity, rect, entity.OneToManyRelationships, true);
				OnMessage("Finished drawing 1-N relationships.");
			}

			// Arrange the shapes to fit the page.
			page.Layout();
			page.ResizeToFitContents();
		}

		private void CacheEntities(EntityMetadata entity)
		{
			var names = new List<string>();

			foreach (var entityRelationship in entity.ManyToManyRelationships)
			{
				var currentManyToManyRelationship = entityRelationship;
				names.Add(string.Compare(entity.LogicalName
					          , currentManyToManyRelationship.Entity1LogicalName, StringComparison.Ordinal) != 0
					          ? currentManyToManyRelationship.Entity1LogicalName
					          : currentManyToManyRelationship.Entity2LogicalName);
			}

			foreach (var entityRelationship in entity.ManyToOneRelationships)
			{
				var currentOneToManyRelationship = entityRelationship;
				names.Add(currentOneToManyRelationship.ReferencedEntity);
			}


			foreach (var entityRelationship in entity.OneToManyRelationships)
			{
				var currentOneToManyRelationship = entityRelationship;
				names.Add(currentOneToManyRelationship.ReferencingEntity);
			}

			if (Settings.RelateSelectedOnly)
			{
				names = names.Where(name => Settings.EntitiesSelected.Contains(name)).ToList();
			}

			if (Cancel)
			{
				return;
			}

			GetEntities(names.Distinct().ToArray());
		}

		/// <summary>
		///     Draw on a Visio page the entity relationships defined in the passed-in relationship collection.
		/// </summary>
		/// <param name="entity">Core entity</param>
		/// <param name="rect">Shape representing the core entity</param>
		/// <param name="relationshipCollection">Collection of entity relationships to draw</param>
		/// <param name="areReferencingRelationships">Whether or not the core entity is the referencing entity in the relationship</param>
		private void DrawRelationships(EntityMetadata entity, VisioApi.Shape rect
			, IEnumerable<RelationshipMetadataBase> relationshipCollection,
			bool areReferencingRelationships)
		{
			AttributeMetadata attribute2 = null;
			AttributeMetadata attribute = null;
			var metadataId = Guid.NewGuid();
			var isManyToMany = false;

			// Draw each relationship in the relationship collection.
			foreach (var entityRelationship in relationshipCollection)
			{
				if (Cancel)
				{
					return;
				}

				EntityMetadata entity2 = null;

				if (Settings.CustomOnly && (!entityRelationship.IsCustomRelationship.HasValue
				              || !entityRelationship.IsCustomRelationship.Value))
				{
					continue;
				}

				if (entityRelationship is ManyToManyRelationshipMetadata)
				{
					isManyToMany = true;
					var currentManyToManyRelationship = entityRelationship as ManyToManyRelationshipMetadata;
					// The entity passed in is not necessarily the originator of this relationship.

					var entityTemp = string.Compare(entity.LogicalName
						                 , currentManyToManyRelationship.Entity1LogicalName, StringComparison.Ordinal) != 0
						                 ? currentManyToManyRelationship.Entity1LogicalName
						                 : currentManyToManyRelationship.Entity2LogicalName;

					if (Settings.RelateSelectedOnly && !Settings.EntitiesSelected.Contains(entityTemp))
					{
						continue;
					}

					entity2 = GetEntityMetadata(entityTemp);

					attribute2 = GetAttributeMetadata(entity2, entity2.PrimaryIdAttribute);
					attribute = GetAttributeMetadata(entity, entity.PrimaryIdAttribute);

					if (currentManyToManyRelationship.MetadataId != null)
					{
						metadataId = currentManyToManyRelationship.MetadataId.Value;
					}
				}
				else if (entityRelationship is OneToManyRelationshipMetadata)
				{
					isManyToMany = false;
					var currentOneToManyRelationship = entityRelationship as OneToManyRelationshipMetadata;

					var entityTemp = areReferencingRelationships
								? currentOneToManyRelationship.ReferencingEntity
								: currentOneToManyRelationship.ReferencedEntity;

					if (Settings.RelateSelectedOnly && !Settings.EntitiesSelected.Contains(entityTemp))
					{
						continue;
					}

					entity2 = GetEntityMetadata(entityTemp);
					attribute2 = GetAttributeMetadata(
						entity2,
						areReferencingRelationships
							? currentOneToManyRelationship.ReferencingAttribute
							: currentOneToManyRelationship.ReferencedAttribute);
					attribute = GetAttributeMetadata(
						entity,
						areReferencingRelationships
							? currentOneToManyRelationship.ReferencedAttribute
							: currentOneToManyRelationship.ReferencingAttribute);

					if (currentOneToManyRelationship.MetadataId != null)
					{
						metadataId = currentOneToManyRelationship.MetadataId.Value;
					}
				}

				// Verify relationship is either ManyToManyMetadata or OneToManyMetadata
				if (entity2 == null || (Settings.RelateSelectedOnly && !Settings.EntitiesSelected.Contains(entity2.LogicalName)))
				{
					continue;
				}

				if (!toBeDrawnEntities.Contains(entity2.LogicalName)
				    && !pageDrawnEntities.Contains(entity2.LogicalName))
				{
					if (Settings.RelateSelectedOnly)
					{
						if (Settings.EntitiesSelected.Contains(entity2.LogicalName))
						{
							toBeDrawnEntities.Enqueue(entity2.LogicalName);

						}
					}
					else
					{
						toBeDrawnEntities.Enqueue(entity2.LogicalName);
					}
				}

				if (processedRelationships.Contains(metadataId))
				{
					// Skip relationships we have already drawn
					continue;
				}

				// Record we are drawing this relationship
				processedRelationships.Add(metadataId);

				DrawRelationship(entity, rect, areReferencingRelationships, entity2, attribute, attribute2, entityRelationship,
					isManyToMany);
			}
		}

		#region Draw

		private void DrawRelationship(EntityMetadata entity, VisioApi.Shape rect, bool areReferencingRelationships,
			EntityMetadata entity2, AttributeMetadata attribute, AttributeMetadata attribute2,
			RelationshipMetadataBase entityRelationship, bool isManyToMany)
		{
			// Do not draw relationships involving the entity itself, BusinessUnit,
			// or those that are intentionally excluded.
			if (string.Compare(entity2.LogicalName, "systemuser", StringComparison.Ordinal) != 0 &&
			    string.Compare(entity2.LogicalName, rect.Name, StringComparison.Ordinal) != 0 &&
			    string.Compare(entity.LogicalName, "systemuser", StringComparison.Ordinal) != 0 &&
			    !excludedRelationsTable.ContainsKey(attribute.LogicalName.GetHashCode()))
			{
				// Either find or create a shape that represents this secondary entity, and add the name of
				// the involved attribute to the shape's text.
				VisioApi.Shape rect2;
				try
				{
					rect2 = rect.ContainingPage.Shapes.ItemU[entity2.SchemaName];

					if (rect2.Text.IndexOf(attribute2.SchemaName, StringComparison.Ordinal) == -1)
					{
						rect2.CellsSRC[VISIO_SECTION_OJBECT_INDEX
							, (short) VisioApi.VisRowIndices.visRowXFormOut
							, (short) VisioApi.VisCellIndices.visXFormHeight].ResultIU += HEIGHT;
						var label
							= attribute2.DisplayName.UserLocalizedLabel == null
								  ? ""
								  : attribute2.DisplayName.UserLocalizedLabel.Label + "\n";
						rect2.Text += "\n" + label + "(" + attribute2.SchemaName + ")";

						// If the attribute is a primary key for the entity, append a [PK] label to the attribute name to indicate this.
						if (string.CompareOrdinal(entity2.PrimaryIdAttribute, attribute2.LogicalName) == 0)
						{
							rect2.Text += "  [PK]";
						}
					}
				}
				catch (COMException)
				{
					var label
						= entity2.DisplayName.UserLocalizedLabel == null
							  ? ""
							  : entity2.DisplayName.UserLocalizedLabel.Label + "\n";

					entity2.OwnershipType = entity2.OwnershipType ?? OwnershipTypes.None;

					rect2 = DrawEntityRectangle(
						rect.ContainingPage, label + "(" + entity2.SchemaName + ")",
						entity2.SchemaName,
						entity2.OwnershipType.Value);
					rect2.CellsSRC[VISIO_SECTION_OJBECT_INDEX
						, (short) VisioApi.VisRowIndices.visRowXFormOut
						, (short) VisioApi.VisCellIndices.visXFormHeight].ResultIU += HEIGHT;
					label
						= attribute2.DisplayName.UserLocalizedLabel == null
							  ? ""
							  : attribute2.DisplayName.UserLocalizedLabel.Label + "\n";
					rect2.Text += "\n" + label + "(" + attribute2.SchemaName + ")";

					// If the attribute is a primary key for the entity, append a [PK] label to the attribute name to indicate so.
					if (string.CompareOrdinal(entity2.PrimaryIdAttribute, attribute2.LogicalName) == 0)
					{
						rect2.Text += "  [PK]";
					}
				}

				// Add the name of the involved attribute to the core entity's text, if not already present.
				if (rect.Text.IndexOf(attribute.SchemaName, StringComparison.Ordinal) == -1)
				{
					rect.CellsSRC[VISIO_SECTION_OJBECT_INDEX
						, (short) VisioApi.VisRowIndices.visRowXFormOut
						, (short) VisioApi.VisCellIndices.visXFormHeight].ResultIU += HEIGHT;
					var label
						= attribute.DisplayName.UserLocalizedLabel == null
							  ? ""
							  : attribute.DisplayName.UserLocalizedLabel.Label + "\n";
					rect.Text += "\n" + label + "(" + attribute.SchemaName + ")";

					// If the attribute is a primary key for the entity, append a [PK] label to the attribute name to indicate so.
					if (string.CompareOrdinal(entity.PrimaryIdAttribute, attribute.LogicalName) == 0)
					{
						rect.Text += "  [PK]";
					}
				}

				// Update the style of the entity name
				var characters = rect.Characters;
				var characters2 = rect2.Characters;

				//set the font family of the text to segoe for the visio 2013.
				if (versionName == "15.0")
				{
					characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterFont] = FONT_STYLE;
					characters2.CharProps[(short) VisioApi.VisCellIndices.visCharacterFont] = FONT_STYLE;
				}

				entity2.OwnershipType = entity2.OwnershipType ?? OwnershipTypes.None;

				switch (entity2.OwnershipType.Value)
				{
					case OwnershipTypes.BusinessOwned:
						// set the font color of the text
						characters2.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visBlack;
						break;
					case OwnershipTypes.OrganizationOwned:
						// set the font color of the text
						characters2.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visBlack;
						break;
					case OwnershipTypes.UserOwned:
						// set the font color of the text
						characters2.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visWhite;
						break;
					default:
						// set the font color of the text
						characters2.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visBlack;
						break;
				}

				entity.OwnershipType = entity2.OwnershipType ?? OwnershipTypes.None;

				switch (entity.OwnershipType.Value)
				{
					case OwnershipTypes.BusinessOwned:
						// set the font color of the text
						characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visBlack;
						break;
					case OwnershipTypes.OrganizationOwned:
						// set the font color of the text
						characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visBlack;
						break;
					case OwnershipTypes.UserOwned:
						// set the font color of the text
						characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visWhite;
						break;
					default:
						// set the font color of the text
						characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
							= (short) VisioApi.VisDefaultColors.visBlack;
						break;
				}

				// Draw the directional, dynamic connector between the two entity shapes.
				if (areReferencingRelationships)
				{
					DrawDirectionalDynamicConnector(rect, rect2, entityRelationship.SchemaName, isManyToMany);
				}
				else
				{
					DrawDirectionalDynamicConnector(rect2, rect, entityRelationship.SchemaName, isManyToMany);
				}
			}
			else
			{
				Debug.WriteLine(string.Format("<{0} - {1}> not drawn.", rect.Name, entity2.LogicalName), "Relationship");
			}
		}

		/// <summary>
		///     Draw an "Entity" Rectangle
		/// </summary>
		/// <param name="page">The Page on which to draw</param>
		/// <param name="entityName">The name of the entity</param>
		/// <param name="schemaName"></param>
		/// <param name="ownership">The ownership type of the entity</param>
		/// <returns>The newly drawn rectangle</returns>
		private VisioApi.Shape DrawEntityRectangle(VisioApi.Page page, string entityName, string schemaName,
			OwnershipTypes ownership)
		{
			var rect = page.DrawRectangle(X_POS1, Y_POS1, X_POS2, Y_POS2);
			rect.Name = schemaName;
			rect.Text = entityName + " ";

			// Determine the shape fill color based on entity ownership.
			string fillColor;

			// Update the style of the entity name
			var characters = rect.Characters;
			characters.Begin = 0;
			characters.End = entityName.Length;
			characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterStyle]
				= (short) VisioApi.VisCellVals.visBold;
			characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterSize] = NAME_CHARACTER_SIZE;
			//set the font family of the text to segoe for the visio 2013.
			if (versionName == "15.0")
			{
				characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterFont] = FONT_STYLE;
			}

			switch (ownership)
			{
				case OwnershipTypes.BusinessOwned:
					fillColor = "RGB(255,140,0)"; // orange
					// set the font color of the text
					characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
						= (short) VisioApi.VisDefaultColors.visBlack;
					break;
				case OwnershipTypes.OrganizationOwned:
					fillColor = "RGB(127, 186, 0)"; // green
					// set the font color of the text
					characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
						= (short) VisioApi.VisDefaultColors.visBlack;
					break;
				case OwnershipTypes.UserOwned:
					fillColor = "RGB(0,24,143)"; // blue 
					// set the font color of the text
					characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
						= (short) VisioApi.VisDefaultColors.visWhite;
					break;
				default:
					fillColor = "RGB(255,255,255)"; // White
					// set the font color of the text
					characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterColor]
						= (short) VisioApi.VisDefaultColors.visBlack;
					break;
			}

			// Set the fill color, placement properties, and line weight of the shape.
			rect.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowMisc
				, (short) VisioApi.VisCellIndices.visLOFlags]
				.FormulaU = ((int) VisioApi.VisCellVals.visLOFlagsPlacable).ToString();
			rect.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowFill
				, (short) VisioApi.VisCellIndices.visFillForegnd].FormulaU = fillColor;
			return rect;
		}

		/// <summary>
		///     Draw a directional, dynamic connector between two entities, representing an entity relationship.
		/// </summary>
		/// <param name="shapeFrom">Shape initiating the relationship</param>
		/// <param name="shapeTo">Shape referenced by the relationship</param>
		/// <param name="relationName"></param>
		/// <param name="isManyToMany">Whether or not it is a many-to-many entity relationship</param>
		private void DrawDirectionalDynamicConnector(VisioApi.Shape shapeFrom, VisioApi.Shape shapeTo
			, string relationName, bool isManyToMany)
		{
			// Add a dynamic connector to the page.
			VisioApi.Shape connectorShape = shapeFrom.ContainingPage.Drop(application.ConnectorToolDataObject, 0.0, 0.0);

			connectorShape.Text = relationName;
			var characters = connectorShape.Characters;
			characters.Begin = 0;
			characters.End = relationName.Length;
			characters.CharProps[(short) VisioApi.VisCellIndices.visCharacterSize] = 5;

			// Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
			connectorShape.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowFill
				, (short) VisioApi.VisCellIndices.visFillShdwPattern].ResultIU = SHDW_PATTERN;
			connectorShape.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowLine
				, (short) VisioApi.VisCellIndices.visLineBeginArrow]
				.ResultIU = isManyToMany ? BEGIN_ARROW_MANY : BEGIN_ARROW;
			connectorShape.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowLine
				, (short) VisioApi.VisCellIndices.visLineEndArrow].ResultIU = END_ARROW;
			connectorShape.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowLine
				, (short) VisioApi.VisCellIndices.visLineColor]
				.ResultIU = isManyToMany ? LINE_COLOR_MANY : LINE_COLOR;
			connectorShape.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowLine
				, (short) VisioApi.VisCellIndices.visLinePattern].ResultIU = LINE_PATTERN;
			connectorShape.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowFill
				, (short) VisioApi.VisCellIndices.visLineRounding].ResultIU = ROUNDING;

			// Connect the starting point.
			connectorShape
				.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowXForm1D
					, (short) VisioApi.VisCellIndices.vis1DBeginX]
				.GlueTo(shapeFrom
					.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowXFormOut
						, (short) VisioApi.VisCellIndices.visXFormPinX]);

			// Connect the ending point.
			connectorShape
				.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowXForm1D
					, (short) VisioApi.VisCellIndices.vis1DEndX]
				.GlueTo(shapeTo
					.CellsSRC[VISIO_SECTION_OJBECT_INDEX, (short) VisioApi.VisRowIndices.visRowXFormOut
						, (short) VisioApi.VisCellIndices.visXFormPinX]);
		}

		#endregion

		#region CRM

		private EntityMetadata[] GetEntities(params string[] entityNames)
		{
			var entitiesTemp = new List<EntityMetadata>();
			if (entityNames.Any())
			{
				var entityNamesList = new List<string>(entityNames);
				foreach (var entityName in entityNames)
				{
					if (entitiesCache.Any(md => md.LogicalName == entityName))
					{
						//OnMessage(string.Format("[Cached] Fetching entity: {0} ... ", entityName));
						entitiesTemp.Add(entitiesCache.Find(md => md.LogicalName == entityName));
						entityNamesList.Remove(entityName);
					}
				}

				entityNames = entityNamesList.ToArray();
			}

			var entityFilter = new MetadataFilterExpression(LogicalOperator.And);
			if (entityNames.Any())
			{
				OnMessage(string.Format("Fetching entities: {0} ... ", entityNames.Aggregate((e1, e2) => e1 + ", " + e2)));
				entityFilter.Conditions.Add(new MetadataConditionExpression("LogicalName"
					, MetadataConditionOperator.In, entityNames));
			}
			else
			{
				return new EntityMetadata[0];
			}

			var entityProperties = new MetadataPropertiesExpression
			{
				AllProperties = false
			};
			entityProperties.PropertyNames.AddRange(
				"LogicalName", "PrimaryIdAttribute", "DisplayName", "SchemaName", "OwnershipType", "OneToManyRelationships"
				, "ManyToOneRelationships", "ManyToManyRelationships", "Attributes");

			var attributeProperties = new MetadataPropertiesExpression
			{
				AllProperties = false
			};
			attributeProperties.PropertyNames.AddRange("DisplayName", "LogicalName", "SchemaName");

			var relationshipProperties = new MetadataPropertiesExpression
			{
				AllProperties = false
			};
			relationshipProperties.PropertyNames.AddRange("IsCustomRelationship", "ReferencedAttribute", "ReferencedEntity",
				"ReferencingEntity", "ReferencingAttribute", "SchemaName", "Entity1LogicalName",
				"Entity2LogicalName");

			var entityQueryExpression = new EntityQueryExpression
			{
				Criteria = entityFilter,
				Properties = entityProperties,
				AttributeQuery = new AttributeQueryExpression
				{
					Properties = attributeProperties
				},
				RelationshipQuery = new RelationshipQueryExpression
				{
					Properties = relationshipProperties
				}
			};

			var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
			                                     {
				                                     Query = entityQueryExpression,
				                                     ClientVersionStamp = null
			                                     };

			entitiesTemp.AddRange(
				((RetrieveMetadataChangesResponse) Service.Execute(retrieveMetadataChangesRequest)).EntityMetadata);

			entitiesTemp.Where(entity => entitiesCache.All(md => md.LogicalName != entity.LogicalName)).ToList().ForEach(entitiesCache.Add);

			return entitiesTemp.ToArray();
		}

		/// <summary>
		///     Retrieves an entity from the local copy of CRM Metadata
		/// </summary>
		/// <param name="entityName">The name of the entity to find</param>
		/// <returns>NULL if the entity was not found, otherwise the entity's metadata</returns>
		private EntityMetadata GetEntityMetadata(string entityName)
		{
			if (entitiesCache.Any(md => md.LogicalName == entityName))
			{
				//OnMessage(string.Format("[Cached] Fetching entity: {0} ... ", entityName));
				return entitiesCache.Find(md => md.LogicalName == entityName);
			}

			OnMessage(string.Format("Fetching entity: {0} ... ", entityName));

			var entityFilter = new MetadataFilterExpression(LogicalOperator.And);
			entityFilter.Conditions.Add(new MetadataConditionExpression("LogicalName"
				, MetadataConditionOperator.Equals, entityName));
			var entityProperties = new MetadataPropertiesExpression
			                       {
				                       AllProperties = false
			                       };
			entityProperties.PropertyNames.AddRange(
				"LogicalName", "PrimaryIdAttribute", "DisplayName", "SchemaName", "OwnershipType", "OneToManyRelationships"
				, "ManyToOneRelationships", "ManyToManyRelationships", "Attributes");

			var attributeProperties = new MetadataPropertiesExpression
			                          {
				                          AllProperties = false
			                          };
			attributeProperties.PropertyNames.AddRange("DisplayName", "LogicalName", "SchemaName");

			var relationshipProperties = new MetadataPropertiesExpression
			                             {
				                             AllProperties = false
			                             };
			relationshipProperties.PropertyNames.AddRange("IsCustomRelationship", "ReferencedAttribute", "ReferencedEntity",
				"ReferencingEntity", "ReferencingAttribute", "SchemaName", "Entity1LogicalName",
				"Entity2LogicalName");

			var entityQueryExpression = new EntityQueryExpression
			                            {
				                            Criteria = entityFilter,
				                            Properties = entityProperties,
				                            AttributeQuery = new AttributeQueryExpression
				                                             {
					                                             Properties = attributeProperties
				                                             },
				                            RelationshipQuery = new RelationshipQueryExpression
				                                                {
					                                                Properties = relationshipProperties
				                                                }
			                            };

			var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
			                                     {
				                                     Query = entityQueryExpression,
				                                     ClientVersionStamp = null
			                                     };
			var results = (RetrieveMetadataChangesResponse) Service.Execute(retrieveMetadataChangesRequest);

			results.EntityMetadata.ToList().ForEach(entitiesCache.Add);

			return results.EntityMetadata.FirstOrDefault(md => md.LogicalName == entityName);
		}

		/// <summary>
		///     Retrieves an attribute from an EntityMetadata object
		/// </summary>
		/// <param name="entity">The entity metadata that contains the attribute</param>
		/// <param name="attributeName">The name of the attribute to find</param>
		/// <returns>NULL if the attribute was not found, otherwise the attribute's metadata</returns>
		private static AttributeMetadata GetAttributeMetadata(EntityMetadata entity, string attributeName)
		{
			return entity.Attributes.FirstOrDefault(attrib => attrib.LogicalName == attributeName);
		}

		internal IOrganizationService GetConnection()
		{
			OnMessage("Creating connection to CRM ...");

			var service = new OrganizationService(CrmConnection.Parse(Settings.GetOrganizationCrmConnectionString()));

			OnMessage("Connection created.");

			return service;
		}

		#endregion
	}
}
