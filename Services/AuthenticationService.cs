using DocumentLoadSanityCheckerDownload.Models.Configs;
using IdentityModel.Client;

namespace DocumentLoadSanityCheckerDownload.Services
{
    public sealed class AuthenticationService : IHostedService, IDisposable
    {
        #region Private members variables
        private readonly SDxConfig sdxConfig;
        private readonly ILogger<AuthenticationService> logger;
        private Timer? refreshTokenTimer;
        #endregion

        #region Public members
        public TokenResponse? tokenResponse { get; set; }
        #endregion

        #region Constructors
        public AuthenticationService(SDxConfig sdxConfig, ILogger<AuthenticationService> logger)
        {
            this.logger = logger;
            this.sdxConfig = sdxConfig;
        }
        #endregion

        #region Public Methods
        public Task StartAsync(CancellationToken cancellationToken)
        {
            tokenResponse = GetOAuthTokenClientCredentialsFlow(sdxConfig.AuthServerAuthority, sdxConfig.AuthClientId, sdxConfig.AuthClientSecret, sdxConfig.ServerResourceID);
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            logger.LogInformation("Stoping token refresh thread...");

            this.refreshTokenTimer?.Change(Timeout.Infinite, 0);

            return Task.CompletedTask;
        }

        public void Dispose()
        {
            logger.LogInformation("AuthenticationService disposing...");

            this.refreshTokenTimer?.Dispose();
        }
        #endregion

        #region Private Methods
        private TokenResponse GetOAuthTokenClientCredentialsFlow(string authServerAuthority, string authClientId, string authClientSecret, string serverResourceID)
        {
            var client = new HttpClient();

            var discoveryDocument = client.GetDiscoveryDocumentAsync(new DiscoveryDocumentRequest
            {
                Address = authServerAuthority,
                Policy =
                    {
                        ValidateEndpoints = false,
                        ValidateIssuerName = false,
                        RequireHttps = true
                    }
            }).Result;


            var parameters = new Parameters(string.IsNullOrEmpty(serverResourceID) ? new Dictionary<string, string>() : new Dictionary<string, string>() { { "Resource", serverResourceID } });

            var response = client.RequestClientCredentialsTokenAsync(new ClientCredentialsTokenRequest
            {
                Address = discoveryDocument.TokenEndpoint,
                ClientId = authClientId,
                ClientSecret = authClientSecret,
                Scope = serverResourceID,
                Parameters = parameters
            }).Result;

            if (!string.IsNullOrWhiteSpace(response.Error))
            {
                logger.LogError(response.Error);
                throw new ArgumentNullException(response.Error);
            }

            if (!string.IsNullOrWhiteSpace(response.AccessToken))
            {
                logger.LogInformation("Token obtained successfully");
            }

            TokenExpirationCall(response.ExpiresIn, authServerAuthority, authClientId, authClientSecret, serverResourceID);

            return response;
        }

        private void TokenExpirationCall(int timeTillExpires, string authServerAuthority, string authClientId, string authClientSecret, string serverResourceID)
        {
            // Sets token refresh callback before the token expires so that we don't get 401's by using expired tokens
            this.refreshTokenTimer = new Timer(_ =>
            {
                logger.LogInformation("Token expiring obtaining new one...");
                this.tokenResponse = GetOAuthTokenClientCredentialsFlow(authServerAuthority, authClientId, authClientSecret, serverResourceID);
            }
                , null, (int)TimeSpan.FromSeconds(timeTillExpires - 30).TotalMilliseconds, Timeout.Infinite);
            logger.LogInformation("Refresh token timer initialized successfully");
        }

        #endregion
    }
}
