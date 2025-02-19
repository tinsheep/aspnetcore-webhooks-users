// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using Microsoft.IdentityModel.Protocols;

namespace GraphWebhooks
{
    /// <summary>
    /// Extension functions for ChangeNotificationCollection to validate
    /// v2 tokens issued by Microsoft identity platform
    ///
    /// This is similar to the built-in methods in the Graph SDK
    /// https://github.com/microsoftgraph/msgraph-sdk-dotnet-core/blob/dev/src/Microsoft.Graph.Core/Extensions/ITokenValidableExtension.cs
    /// The difference here is the Graph SDK assumes v1 tokens, and Graph can
    /// send v2 tokens
    /// </summary>
    public static class ChangeNotificationCollectionExtensions
    {
        /// <summary>
        /// Validates all tokens contained in a ChangeNotificationCollection. If there are none, returns true.
        /// </summary>
        /// <param name="collection">The ChangeNotificationCollection to validate</param>
        /// <param name="tenantIds">A set of tenant IDs that can appear in the token issuer claim</param>
        /// <param name="appIds">A set of app IDs that can appear in the audience claim</param>
        /// <param name="wellKnownUri">The well-known OpenID config URI (Default: https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration)</param>
        /// <param name="issuerPrefix">The prefix for valid issuers (Default: https://login.microsoftonline.com/)</param>
        /// <returns>true if all tokens are valid, false otherwise</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<bool> AreTokensValidV2(
            this ChangeNotificationCollection collection,
            IEnumerable<Guid> tenantIds,
            IEnumerable<Guid> appIds,
            string wellKnownUri = "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
            string issuerPrefix = "https://login.microsoftonline.com/")
        {
            if ((collection.ValidationTokens == null || !collection.ValidationTokens.Any()) && collection.Value.All(x => x.EncryptedContent == null))
                return true;

            if (tenantIds == null || !tenantIds.Any())
                throw new ArgumentNullException(nameof(tenantIds));
            if (appIds == null || !appIds.Any())
                throw new ArgumentNullException(nameof(appIds));

            var configurationManager = new ConfigurationManager<OpenIdConnectConfiguration>(
                wellKnownUri, new OpenIdConnectConfigurationRetriever());

            var openIdConfig = await configurationManager.GetConfigurationAsync();
            var handler = new JwtSecurityTokenHandler();
            var issuersToValidate = tenantIds.Select(tid => $"{issuerPrefix}{tid}/v2.0");
            var appIdsToValidate = appIds.Select(appId => appId.ToString());

            return collection.ValidationTokens
                .Select(t => IsTokenValid(t, handler, openIdConfig, issuersToValidate, appIdsToValidate))
                .Aggregate((x, y) => x && y);
        }

        /// <summary>
        /// Checks if a given token is valid
        /// </summary>
        /// <param name="token">The token to validate</param>
        /// <param name="handler">The JwtSecurityTokenHandler to use to validate the token</param>
        /// <param name="openIdConnectConfiguration">OpenID configuration information</param>
        /// <param name="issuersToValidate">A set of valid issuers</param>
        /// <param name="audiences">A set of valid audiences</param>
        /// <returns>true if token is valid, false if not</returns>
        private static bool IsTokenValid(
            string token,
            JwtSecurityTokenHandler handler,
            OpenIdConnectConfiguration openIdConnectConfiguration,
            IEnumerable<string> issuersToValidate,
            IEnumerable<string> audiences)
        {
            try
            {
                handler.ValidateToken(token, new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidateAudience = true,
                    ValidateIssuerSigningKey = true,
                    ValidateLifetime = true,
                    ValidIssuers = issuersToValidate,
                    ValidAudiences = audiences,
                    IssuerSigningKeys = openIdConnectConfiguration.SigningKeys
                }, out _);

                return true;
            }
            catch (SecurityTokenValidationException)
            {
                return false;
            }
        }
    }
}
