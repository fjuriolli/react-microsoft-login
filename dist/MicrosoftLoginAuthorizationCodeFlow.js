"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var msal_browser_1 = require("@azure/msal-browser");
var AbstractMicrosoftLogin_1 = require("./AbstractMicrosoftLogin");
var getUserAgentApp = function (_a) {
    var clientId = _a.clientId, tenantUrl = _a.tenantUrl, redirectUri = _a.redirectUri;
    if (clientId && AbstractMicrosoftLogin_1.CLIENT_ID_REGEX.test(clientId)) {
        return new msal_browser_1.PublicClientApplication({
            auth: __assign(__assign(__assign({}, (redirectUri && { redirectUri: redirectUri })), (tenantUrl && { authority: tenantUrl })), { clientId: clientId, navigateToLoginRequestUrl: false }),
        });
    }
};
var MicrosoftLoginAuthorizationCodeFlow = (function (_super) {
    __extends(MicrosoftLoginAuthorizationCodeFlow, _super);
    function MicrosoftLoginAuthorizationCodeFlow(props) {
        return _super.call(this, props) || this;
    }
    MicrosoftLoginAuthorizationCodeFlow.prototype.getMsalInstance = function () {
        return _super.prototype.getState.call(this).msalInstance;
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.createUserAgent = function (userAgentApp) {
        return getUserAgentApp(userAgentApp);
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.isUserAgentCreated = function () {
        return !!_super.prototype.getState.call(this).msalInstance;
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.handleRedirect = function () {
        var msalInstance = this.getMsalInstance();
        var scopes = _super.prototype.getState.call(this).scopes;
        var _a = this.props, authCallback = _a.authCallback, _b = _a.withUserData, withUserData = _b === void 0 ? false : _b;
        msalInstance
            .handleRedirectPromise()
            .then(function (response) {
            var currentAccounts = msalInstance.getAllAccounts();
            if (currentAccounts === null) {
                console.log("Logged in, but no currentAccounts retrieved");
            }
            else if (currentAccounts.length > 1) {
                console.log("Logged in, but many accounts retrieved", currentAccounts);
            }
            else if (currentAccounts.length === 1) {
                var username = currentAccounts[0].username;
                console.log("Logged in with account " + username + ".");
            }
        })
            .catch(function (err) {
            console.error("Error while redirecting user to login", err);
        });
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.loginRedirect = function (_a) {
        var scopes = _a.scopes, prompt = _a.prompt;
        var msalInstance = this.getMsalInstance();
        msalInstance
            .loginPopup({ scopes: scopes, prompt: prompt })
            .then(function (loginResponse) {
            var currentAccounts = msalInstance.getAllAccounts();
            if (currentAccounts === null) {
                console.log("no accounts detected");
            }
            else if (currentAccounts.length > 1) {
                console.log("Multiple accounts chosen", currentAccounts);
            }
            else if (currentAccounts.length === 1) {
                var username = currentAccounts[0].username;
                console.log("Account " + username + " chosen.");
            }
        })
            .catch(function (error) {
            console.log(error);
        });
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.getGraphAPITokenAndUser = function (_a) {
        var _this = this;
        var scopes = _a.scopes, withUserData = _a.withUserData, authCallback = _a.authCallback, isRedirect = _a.isRedirect;
        var msalInstance = this.getMsalInstance();
        return (msalInstance
            .ssoSilent({ scopes: scopes, loginHint: "get_from_username" })
            .catch(function (error) {
            _this.log("Fetch Graph API 'access_token' in silent mode is FAILED", error, true);
            if (isRedirect) {
                _this.log("Fetch Graph API 'access_token' with redirect STARTED");
                msalInstance.acquireTokenRedirect({ scopes: scopes });
            }
            else {
                _this.log("Fetch Graph API 'access_token' with popup STARTED");
                msalInstance.acquireTokenPopup({ scopes: scopes });
            }
        })
            .then(function (authResponseWithAccessToken) {
            _this.log("Fetch Graph API 'access_token' SUCCEDEED", authResponseWithAccessToken);
            if (withUserData) {
                _this.getUserData(authResponseWithAccessToken);
            }
            else {
                _this.log("Login SUCCEDED");
                authCallback(null, { authResponseWithAccessToken: authResponseWithAccessToken });
            }
        })
            .catch(function (error) {
            _this.log("Login FAILED", error, true);
            authCallback(error);
        }));
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.popupLogin = function (_a) {
        var _this = this;
        var scopes = _a.scopes, withUserData = _a.withUserData, authCallback = _a.authCallback, prompt = _a.prompt;
        _super.prototype.log.call(this, "Fetch Azure AD 'token' with popup STARTED");
        var msalInstance = this.getMsalInstance();
        msalInstance
            .loginPopup({ scopes: scopes, prompt: prompt })
            .then(function (authResult) {
            _this.log("Fetch Azure AD 'token' with popup SUCCEDEED", authResult);
            _this.log("Fetch Graph API 'access_token' in silent mode STARTED");
            _this.getGraphAPITokenAndUser({
                scopes: scopes,
                withUserData: withUserData,
                authCallback: authCallback,
                isRedirect: false,
            });
        })
            .catch(function (error) {
            _this.log("Fetch Azure AD 'token' with popup FAILED", error, true);
            authCallback(error);
        });
    };
    MicrosoftLoginAuthorizationCodeFlow.prototype.getUserData = function (authResponseWithAccessToken) {
        var _this = this;
        var authCallback = this.props.authCallback;
        var accessToken = authResponseWithAccessToken.accessToken;
        this.log("Fetch Graph API user data STARTED");
        var options = {
            method: "GET",
            headers: {
                Authorization: "Bearer " + accessToken,
            },
        };
        return fetch("https://graph.microsoft.com/v1.0/me", options)
            .then(function (response) { return response.json(); })
            .then(function (userData) {
            _this.log("Fetch Graph API user data SUCCEDEED", userData);
            _this.log("Login SUCCEDED");
            console.log("userData", userData);
        });
    };
    return MicrosoftLoginAuthorizationCodeFlow;
}(AbstractMicrosoftLogin_1.default));
exports.default = MicrosoftLoginAuthorizationCodeFlow;
//# sourceMappingURL=MicrosoftLoginAuthorizationCodeFlow.js.map