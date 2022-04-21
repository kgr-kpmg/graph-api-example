using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ListSharepointsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly GraphServiceClient _graph;

        public UserController(IConfiguration config, GraphServiceClient graph)
        {
            _config = config;
            _graph = graph;
        }

        [Route("auth")]
        public async Task<IActionResult> Auth([FromQuery(Name = "code")] string code)
        {
            var clientId = _config.GetSection("AzureAd:ClientId").Value;
            var instance = _config.GetSection("AzureAd:Instance").Value;
            var clientSecret = _config.GetSection("AzureAd:ClientSecret").Value;
            var redirectUri = "https://localhost:5001/api/user/auth"; // must be same as in code request
            var scopes = "";

            var uri = $"{instance}organizations/oauth2/v2.0/token";

            var formUrlEncodedContent = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("client_id", clientId),
                new KeyValuePair<string, string>("grant_type", "authorization_code"),
                new KeyValuePair<string, string>("scope", scopes),
                new KeyValuePair<string, string>("code", code),
                new KeyValuePair<string, string>("redirect_uri", redirectUri),
                new KeyValuePair<string, string>("client_secret", clientSecret),
                new KeyValuePair<string, string>("Content-Type", "application/x-www-form-urlencoded")
            });

            var client = new HttpClient();
            var httpResponse = await client.PostAsync(uri, formUrlEncodedContent);
            var response = await JsonSerializer.DeserializeAsync<TokenResponse>(await httpResponse.Content.ReadAsStreamAsync());

            // Bearer token can be used against protected /api/user/sites endpoint in postman
            return Ok(response);
        }

        [Authorize]
        [Route("sites")]
        public async Task<IActionResult> Sites()
        {
            var results = await _graph.Sites
                .Request(new List<QueryOption>
                {
                    new QueryOption("?search", "*")
                })
                .GetAsync();

            var response = results.Select(site => new Site
            {
                Name = site.Name,
                Url = site.WebUrl
            });

            return Ok(response);
        }
    }

    class Site
    {
        public string Name { get; set; }
        public string Url { get; set; }
    }

    class TokenResponse
    {
        [JsonPropertyName("token_type")]
        public string TokenType { get; set; }
        [JsonPropertyName("scope")]
        public string Scope { get; set; }
        [JsonPropertyName("expires_in")]
        public int ExpiresIn { get; set; }
        [JsonPropertyName("access_token")]
        public string AccessToken { get; set; }
        [JsonPropertyName("refresh_token")]
        public string RefreshToken { get; set; }
    }

}
