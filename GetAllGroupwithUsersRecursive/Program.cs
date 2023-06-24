using Azure.Core;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace TestGroupSync
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var graphClient = MSGraphClient.GetAuthenticatedClient(
                "your access token"
            );

            string groupName = "RootGroup";
            List<QueryOption> options = new List<QueryOption>
                    {
                         new QueryOption("$filter", $"displayName eq '{groupName}'")
                    };
            var groups = await graphClient.Groups.Request(options).GetAsync();

            if (groups != null && groups.Any())
            {
                var x = groups.FirstOrDefault().DisplayName;
            }

            var objects = await GetGroupsAndUsers(graphClient, "group id to retrieve e.g. 4943ba8a-367d-49df-80d5-9fb133910664", "");

        }
        static async Task<List<UserOrGroupVM>> GetGroupsAndUsers(GraphServiceClient graphClient, string groupId, string parentId)
        {
            List<UserOrGroupVM> objects = new List<UserOrGroupVM>();

            try
            {
                // Retrieve the specified group
                Group group = await graphClient.Groups[groupId].Request().GetAsync();

                if (group != null)
                {
                    objects.Add(new UserOrGroupVM
                    {
                        //Email = group.Mail,
                        AadObjectId = Guid.Parse(group.Id),
                        IsGroup = true,
                        Name = group.DisplayName,
                        ParentId = string.IsNullOrEmpty(parentId) ? Guid.Empty : Guid.Parse(parentId),
                    });

                    List<QueryOption> options = new List<QueryOption>
                    {
                         new QueryOption("$select", "id,displayName,userPrincipalName,mail,assignedLicenses,department,officeLocation,mobilePhone")
                    };

                    // Retrieve the members of the group
                    var members = await graphClient.Groups[groupId].Members.Request(options).GetAsync();

                    do
                    {
                        if (members != null && members.Count > 0)
                        {
                            foreach (DirectoryObject member in members)
                            {
                                // If the member is a group, recursively retrieve its members
                                if (member is Group)
                                {
                                    List<UserOrGroupVM> subObjects = await GetGroupsAndUsers(graphClient, member.Id, groupId);
                                    objects.AddRange(subObjects);
                                }
                                else if (member is User)
                                {
                                    var graphUser = (User)member;
                                    if (!graphUser.AssignedLicenses.Any()) continue;//if no active licenses then we should not consider the user, may be it is a resource account (room, equipment etc.)
                                    objects.Add(new UserOrGroupVM
                                    {
                                        AadObjectId = Guid.Parse(graphUser.Id),
                                        Name = graphUser.DisplayName,
                                        UserPrincipalName = string.IsNullOrEmpty(graphUser.Mail) ? graphUser.UserPrincipalName : graphUser.Mail,
                                        NotificationEmail = string.IsNullOrEmpty(graphUser.Mail) ? graphUser.UserPrincipalName : graphUser.Mail,
                                        Department = graphUser.Department,
                                        OfficeLocation = graphUser.OfficeLocation,
                                        MobilePhone = graphUser.MobilePhone,
                                        ParentId = Guid.Parse(groupId),

                                    });
                                }
                            }
                        }
                    } while (members.NextPageRequest != null && (members = await members.NextPageRequest.GetAsync()).Count > 0);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            return objects;
        }
    }

    public class MSGraphClient
    {
        public static GraphServiceClient GetAuthenticatedClient(string accessToken)
        {
            var delegateAuthProvider = new DelegateAuthenticationProvider((requestMessage) =>
            {
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                return Task.FromResult(0);
            });

            var graphClient = new GraphServiceClient(delegateAuthProvider);

            return graphClient;
        }
    }
    public class UserOrGroupVM
    {
        public Guid AadObjectId { get; set; }
        public string Name { get; set; }
        public string UserPrincipalName { get; set; }
        public string NotificationEmail { get; set; }
        public string Department { get; set; }
        public string OfficeLocation { get; set; }
        public string MobilePhone { get; set; }
        public bool IsGroup { get; set; }
        public Guid ParentId { get; set; }
    }
}
