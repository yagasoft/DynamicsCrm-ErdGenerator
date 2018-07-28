using System;
using System.Collections.ObjectModel;
using System.Linq;
using CRM_ERD_Generator_GUI.Model;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Discovery;

namespace CRM_ERD_Generator_GUI.Helpers
{
    public class ConnectionHelper
    {
        public static OrganizationDetail GetOrganizationDetails(Settings settings)
        {
            var orgs = GetOrganizations(settings);
            var details = orgs.FirstOrDefault(d => d.UrlName == settings.CrmOrg);
            return details;
        }
        public static ObservableCollection<string> GetOrgList(Settings settings)
        {
            var orgs = GetOrganizations(settings);
            var newOrgs = new ObservableCollection<String>(orgs.Select(d => d.UrlName).ToList());
            return newOrgs;
        }
        public static OrganizationDetailCollection GetOrganizations(Settings settings)
        {
            var connection = CrmConnection.Parse(settings.GetDiscoveryCrmConnectionString());
            var service = new DiscoveryService(connection);

            var request = new RetrieveOrganizationsRequest();
            var response = (RetrieveOrganizationsResponse)service.Execute(request);
            return response.Details;
        }
    }
}
