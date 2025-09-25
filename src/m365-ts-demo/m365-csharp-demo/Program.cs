using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

var cfg = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json").Build();

var cred = new ClientSecretCredential(
    cfg["TenantId"], cfg["ClientId"], cfg["ClientSecret"]);

var graph = new GraphServiceClient(cred, scopes:new[]{"https://graph.microsoft.com/.default"});

var user = await graph.Me.GetAsync();
Console.WriteLine($"Hello {user.DisplayName}");

var messages = await graph.Me.Messages.GetAsync(req =>
{
    req.QueryParameters.Top = 5;
    req.QueryParameters.Select = new[] { "subject", "from" };
});
foreach (var m in messages.Value)
    Console.WriteLine("Mail: " + m.Subject);

var children = await graph.Me.Drive.Root.Children.GetAsync();
Console.WriteLine("Root files count " + children.Value.Count);
