// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using GraphWebhooks.Models;
using GraphWebhooks.Services;
using Microsoft.Graph.ExternalConnectors;

namespace GraphWebhooks.Controllers
{
    /// <summary>
    /// Implements subscription management endpoints
    /// </summary>
    public class WatchController : Controller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly SubscriptionStore _subscriptionStore;
        private readonly CertificateService _certificateService;
        private readonly ILogger<WatchController> _logger;
        private readonly string _notificationHost;
        private readonly string _teamId;
        private readonly string _channelId;

        public WatchController(
            GraphServiceClient graphClient,
            SubscriptionStore subscriptionStore,
            CertificateService certificateService,
            ILogger<WatchController> logger,
            IConfiguration configuration)
        {
            _graphClient = graphClient ?? throw new ArgumentException(nameof(graphClient));
            _subscriptionStore = subscriptionStore ?? throw new ArgumentException(nameof(subscriptionStore));
            _certificateService = certificateService ?? throw new ArgumentException(nameof(certificateService));
            _logger = logger ?? throw new ArgumentException(nameof(logger));
            _ = configuration ?? throw new ArgumentException(nameof(configuration));
            _teamId = configuration.GetValue<string>("Teams:TeamId");
            _channelId = configuration.GetValue<string>("Teams:ChannelId");

            _notificationHost = configuration.GetValue<string>("NotificationHost");
            if (string.IsNullOrEmpty(_notificationHost) || _notificationHost == "YOUR_NGROK_PROXY")
            {
                throw new ArgumentException("You must configure NotificationHost in appsettings.json");
            }

        }

        /// <summary>
        /// GET /watch/delegated
        /// Creates a new subscription to the authenticated user's inbox and
        /// displays a page that updates with each received notification
        /// </summary>
        /// <returns></returns>
        [AuthorizeForScopes(ScopeKeySection = "GraphScopes")]
        public async Task<IActionResult> Delegated()
        {
            try
            {
                // Delete any existing subscriptions for the user
                await DeleteAllSubscriptions(false);

                // Get the user's ID and tenant ID from the user's identity
                var userId = User.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;
                _logger.LogInformation($"Authenticated user ID {userId}");
                var tenantId = User.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;

                // Get the user from Microsoft Graph
                var user = await _graphClient.Me
                    .Request()
                    .Select(u => new {u.DisplayName, u.Mail, u.UserPrincipalName})
                    .GetAsync();

                _logger.LogInformation($"Authenticated user: {user.DisplayName} ({user.Mail ?? user.UserPrincipalName})");
                // Add the user's display name and email address to the user's
                // identity.
                User.AddUserGraphInfo(user);

                // Create the subscription
                var subscription = new Subscription
                {
                    ChangeType = "created",
                    NotificationUrl = $"{_notificationHost}/listen",
                    Resource = "me/mailfolders/inbox/messages",
                    ClientState = Guid.NewGuid().ToString(),
                    IncludeResourceData = false,
                    // Subscription only lasts for one hour
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddHours(1)
                };

                var newSubscription = await _graphClient.Subscriptions
                    .Request().AddAsync(subscription);

                // Add the subscription to the subscription store
                _subscriptionStore.SaveSubscriptionRecord(new SubscriptionRecord
                {
                    Id = newSubscription.Id,
                    UserId = userId,
                    TenantId = tenantId,
                    ClientState = newSubscription.ClientState
                });

                return View(newSubscription).WithSuccess("Subscription created");
            }
            catch (Exception ex)
            {
                // Throw MicrosoftIdentityWebChallengeUserException to allow
                // Microsoft.Identity.Web to challenge the user for re-auth or consent
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException) throw;

                // Otherwise display the error
                return View().WithError($"Error creating subscription: {ex.Message}",
                    ex.ToString());
            }
        }

        /// <summary>
        /// GET /watch/delegated
        /// Creates a new subscription for the authenticated user's chats
        /// displays a page that updates with each received notification
        /// </summary>
        /// <returns></returns>
        [AuthorizeForScopes(ScopeKeySection = "GraphScopes")]
        public async Task<IActionResult> DelegatedChats()
        {
            try
            {
                // Delete any existing subscriptions for the user
                await DeleteAllSubscriptions(false);

                // Get the user's ID and tenant ID from the user's identity
                var userId = User.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;
                _logger.LogInformation($"Authenticated user ID {userId}");
                var tenantId = User.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;

                // Get the user from Microsoft Graph
                var user = await _graphClient.Me
                    .Request()
                    .Select(u => new { u.DisplayName, u.Mail, u.UserPrincipalName })
                    .GetAsync();

                _logger.LogInformation($"Authenticated user: {user.DisplayName} ({user.Mail ?? user.UserPrincipalName})");
                // Add the user's display name and email address to the user's
                // identity.
                User.AddUserGraphInfo(user);
                
                //this subscription requires the beta endpoint
                //_graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                // Get the encryption certificate (public key)
                var encryptionCertificate = await _certificateService.GetEncryptionCertificate();

                // Create the subscription
                var subscription = new Subscription
                {
                    ChangeType = "created",
                    NotificationUrl = $"{_notificationHost}/listen",
                    Resource = "/users/" + userId + "/chats/getAllMessages",
                    ClientState = Guid.NewGuid().ToString(),
                    IncludeResourceData = true,
                    EncryptionCertificateId = encryptionCertificate.Subject,
                    // Subscription only lasts for one hour
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddHours(1)

                };

                // To get resource data, we must provide a public key that
                // Microsoft Graph will use to encrypt their key
                // See https://docs.microsoft.com/graph/webhooks-with-resource-data#creating-a-subscription
                subscription.AddPublicEncryptionCertificate(encryptionCertificate);

                var newSubscription = await _graphClient.Subscriptions
                    .Request().AddAsync(subscription);

                // Add the subscription to the subscription store
                _subscriptionStore.SaveSubscriptionRecord(new SubscriptionRecord
                {
                    Id = newSubscription.Id,
                    UserId = userId,
                    TenantId = tenantId,
                    ClientState = newSubscription.ClientState
                });

                return View(newSubscription).WithSuccess("Subscription created");
            }
            catch (Exception ex)
            {
                // Throw MicrosoftIdentityWebChallengeUserException to allow
                // Microsoft.Identity.Web to challenge the user for re-auth or consent
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException) throw;

                // Otherwise display the error
                return View().WithError($"Error creating subscription: {ex.Message}",
                    ex.ToString());
            }
        }

        /// <summary>
        /// GET /watch/delegated
        /// Creates a new subscription for a channel chat
        /// displays a page that updates with each received notification
        /// </summary>
        /// <returns></returns>
        [AuthorizeForScopes(ScopeKeySection = "GraphScopes")]
        public async Task<IActionResult> DelegatedChannelChats()
        {
            try
            {

                // Delete any existing subscriptions for the user
                await DeleteAllSubscriptions(false);

                // Get the user's ID and tenant ID from the user's identity
                var userId = User.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;
                _logger.LogInformation($"Authenticated user ID {userId}");
                var tenantId = User.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;

                // Get the user from Microsoft Graph
                var user = await _graphClient.Me
                    .Request()
                    .Select(u => new { u.DisplayName, u.Mail, u.UserPrincipalName })
                    .GetAsync();

                _logger.LogInformation($"Authenticated user: {user.DisplayName} ({user.Mail ?? user.UserPrincipalName})");
                // Add the user's display name and email address to the user's
                // identity.
                User.AddUserGraphInfo(user);

                //this subscription requires the beta endpoint
                //_graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                // Get the encryption certificate (public key)
                var encryptionCertificate = await _certificateService.GetEncryptionCertificate();

                // Create the subscription
                var subscription = new Subscription
                {
                    ChangeType = "created",
                    NotificationUrl = $"{_notificationHost}/listen",
                    Resource = "/teams/" + _teamId + "/channels/" + _channelId + "/messages",
                    ClientState = Guid.NewGuid().ToString(),
                    IncludeResourceData = true,
                    EncryptionCertificateId = encryptionCertificate.Subject,
                    // Subscription only lasts for one hour
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddHours(1)

                };

                // To get resource data, we must provide a public key that
                // Microsoft Graph will use to encrypt their key
                // See https://docs.microsoft.com/graph/webhooks-with-resource-data#creating-a-subscription
                subscription.AddPublicEncryptionCertificate(encryptionCertificate);

                var newSubscription = await _graphClient.Subscriptions
                    .Request().AddAsync(subscription);

                // Add the subscription to the subscription store
                _subscriptionStore.SaveSubscriptionRecord(new SubscriptionRecord
                {
                    Id = newSubscription.Id,
                    UserId = userId,
                    TenantId = tenantId,
                    ClientState = newSubscription.ClientState
                });

                return View(newSubscription).WithSuccess("Subscription created");
            }
            catch (Exception ex)
            {
                // Throw MicrosoftIdentityWebChallengeUserException to allow
                // Microsoft.Identity.Web to challenge the user for re-auth or consent
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException) throw;

                // Otherwise display the error
                return View().WithError($"Error creating subscription: {ex.Message}",
                    ex.ToString());
            }
        }



        /// <summary>
        /// GET /watch/apponly
        /// Creates a new subscription to all Teams channel messages and
        /// displays a page that updates with each received notification
        /// </summary>
        /// <returns></returns>
        public async Task<IActionResult> AppOnly()
        {
            try
            {
                // Delete any existing Teams channel subscriptions
                // This is important as each app is only allowed one active
                // subscription to the /teams/getAllMessages resource
                await DeleteAllSubscriptions(true);

                var tenantId = User.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;

                // Get the encryption certificate (public key)
                var encryptionCertificate = await _certificateService.GetEncryptionCertificate();

                // Create the subscription
                var subscription = new Subscription
                {
                    ChangeType = "created",
                    NotificationUrl = $"{_notificationHost}/listen",
                    Resource = "/teams/getAllMessages",
                    ClientState = Guid.NewGuid().ToString(),
                    IncludeResourceData = true,
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddHours(1),
                    EncryptionCertificateId = encryptionCertificate.Subject
                };

                // To get resource data, we must provide a public key that
                // Microsoft Graph will use to encrypt their key
                // See https://docs.microsoft.com/graph/webhooks-with-resource-data#creating-a-subscription
                subscription.AddPublicEncryptionCertificate(encryptionCertificate);

                var newSubscription = await _graphClient.Subscriptions
                    .Request()
                    .WithAppOnly()
                    .AddAsync(subscription);

                // Add the subscription to the subscription store
                _subscriptionStore.SaveSubscriptionRecord(new SubscriptionRecord
                {
                    Id = newSubscription.Id,
                    UserId = "APP-ONLY",
                    TenantId = tenantId,
                    ClientState = newSubscription.ClientState
                });

                return View(newSubscription).WithSuccess("Subscription created");
            }
            catch (Exception ex)
            {
                return RedirectToAction("Index", "Home")
                    .WithError($"Error creating subscription: {ex.Message}",
                        ex.ToString());
            }
        }

        /// <summary>
        /// GET /watch/apponly
        /// Creates a new subscription to all Meeting updates and
        /// displays a page that updates with each received notification
        /// </summary>
        /// <returns></returns>
        public async Task<IActionResult> AppOnlyMeetings()
        {
            try
            {
                // Delete any existing Teams channel subscriptions
                // This is important as each app is only allowed one active
                // subscription to the /teams/getAllMessages resource
                await DeleteAllSubscriptions(true);

                var tenantId = User.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;

                //this subscription requires the beta endpoint
                _graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                // Get the encryption certificate (public key)
                var encryptionCertificate = await _certificateService.GetEncryptionCertificate();

                // Create the subscription
                var subscription = new Subscription
                {
                    ChangeType = "created",
                    NotificationUrl = $"{_notificationHost}/listen",
                    Resource = "/communications/onlineMeetings/?$filter=JoinWebUrl eq 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_NjVjNzQwY2MtMjAxYS00OGYzLTk1YTgtYzc1MjFiMzlhOWE2%40thread.v2/0?context=%7b%22Tid%22%3a%22a6218add-62d4-4c65-a992-da1ec5d85e38%22%2c%22Oid%22%3a%22e1999dd6-9b42-462e-9763-ad46585b75bc%22%7d'",
                    ClientState = Guid.NewGuid().ToString(),
                    IncludeResourceData = true,
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddHours(1),
                    EncryptionCertificateId = encryptionCertificate.Subject
                };

                // To get resource data, we must provide a public key that
                // Microsoft Graph will use to encrypt their key
                // See https://docs.microsoft.com/graph/webhooks-with-resource-data#creating-a-subscription
                subscription.AddPublicEncryptionCertificate(encryptionCertificate);

                var newSubscription = await _graphClient.Subscriptions
                    .Request()
                    .WithAppOnly()
                    .AddAsync(subscription);

                // Add the subscription to the subscription store
                _subscriptionStore.SaveSubscriptionRecord(new SubscriptionRecord
                {
                    Id = newSubscription.Id,
                    UserId = "APP-ONLY",
                    TenantId = tenantId,
                    ClientState = newSubscription.ClientState
                });

                return View(newSubscription).WithSuccess("Subscription created");
            }
            catch (Exception ex)
            {
                return RedirectToAction("Index", "Home")
                    .WithError($"Error creating subscription: {ex.Message}",
                        ex.ToString());
            }
        }

        /// <summary>
        /// GET /watch/unsubscribe
        /// Deletes the user's inbox subscription and signs the user out
        /// </summary>
        /// <param name="subscriptionId">The ID of the subscription to delete</param>
        /// <returns></returns>
        public async Task<IActionResult> Unsubscribe(string subscriptionId)
        {
            if (string.IsNullOrEmpty(subscriptionId))
            {
                return RedirectToAction("Index", "Home")
                    .WithError("No subscription ID specified");
            }

            try
            {
                var subscription = _subscriptionStore.GetSubscriptionRecord(subscriptionId);

                var appOnly = subscription.UserId == "APP-ONLY";
                // To unsubscribe, just delete the subscription
                await _graphClient.Subscriptions[subscriptionId]
                    .Request()
                    .WithAppOnly(appOnly)
                    .DeleteAsync();

                // Remove the subscription from the subscription store
                _subscriptionStore.DeleteSubscriptionRecord(subscriptionId);

                // Redirect to Microsoft.Identity.Web's signout page
                return RedirectToAction("SignOut", "Account", new { area = "MicrosoftIdentity" });
            }
            catch (Exception ex)
            {
                // Throw MicrosoftIdentityWebChallengeUserException to allow
                // Microsoft.Identity.Web to challenge the user for re-auth or consent
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException) throw;

                // Otherwise display the error
                return RedirectToAction("Index", "Home")
                    .WithError($"Error deleting subscription: {ex.Message}",
                        ex.ToString());
            }
        }

        /// <summary>
        /// GET /watch/justsignout
        /// signs the user out without deleteing the subscription
        /// </summary>
        /// <param name="subscriptionId">The ID of the subscription to delete</param>
        /// <returns></returns>
        public async Task<IActionResult> JustSignOut()
        {

            try
            {

                // Redirect to Microsoft.Identity.Web's signout page
                return RedirectToAction("SignOut", "Account", new { area = "MicrosoftIdentity" });
            }
            catch (Exception ex)
            {
                // Throw MicrosoftIdentityWebChallengeUserException to allow
                // Microsoft.Identity.Web to challenge the user for re-auth or consent
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException) throw;

                // Otherwise display the error
                return RedirectToAction("Index", "Home")
                    .WithError($"Error deleting subscription: {ex.Message}",
                        ex.ToString());
            }
        }

        /// <summary>
        /// Deletes all current subscriptions
        /// </summary>
        /// <param name="appOnly">If true, all app-only subscriptions are removed. If false, all user subscriptions are removed</param>
        private async Task DeleteAllSubscriptions(bool appOnly)
        {
            try
            {
                // Get all current subscriptions
                var subscriptions = await _graphClient.Subscriptions
                    .Request()
                    .WithAppOnly(appOnly)
                    .GetAsync();

                foreach(var subscription in subscriptions.CurrentPage)
                {
                    // Delete the subscription
                    await _graphClient.Subscriptions[subscription.Id]
                        .Request()
                        .WithAppOnly(appOnly)
                        .DeleteAsync();

                    // Remove the subscription from the subscription store
                    _subscriptionStore.DeleteSubscriptionRecord(subscription.Id);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting existing subscriptions");
            }
        }
    }
}
