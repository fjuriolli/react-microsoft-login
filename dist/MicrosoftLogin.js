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
var React = require("react");
var msal_1 = require("msal");
var MicrosoftLoginButton_1 = require("./MicrosoftLoginButton");
var CLIENT_ID_REGEX = /[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}/;
var getUserAgentApp = function (_a) {
    var clientId = _a.clientId, tenantUrl = _a.tenantUrl, redirectUri = _a.redirectUri;
    if (clientId && CLIENT_ID_REGEX.test(clientId)) {
        return new msal_1.UserAgentApplication({
            auth: __assign(__assign(__assign({}, (redirectUri && { redirectUri: redirectUri })), (tenantUrl && { authority: tenantUrl })), { clientId: clientId, validateAuthority: true, navigateToLoginRequestUrl: false }),
        });
    }
};
var getScopes = function (graphScopes) {
    var scopes = graphScopes || [];
    if (!scopes.find(function (el) { return el.toLowerCase() === "user.read"; })) {
        scopes.push("user.read");
    }
    return scopes;
};
var MicrosoftLogin = (function (_super) {
    __extends(MicrosoftLogin, _super);
    function MicrosoftLogin(props) {
        var _this = _super.call(this, props) || this;
        _this.login = function () {
            var _a = _this.state, msalInstance = _a.msalInstance, scopes = _a.scopes;
            var _b = _this.props, _c = _b.withUserData, withUserData = _c === void 0 ? false : _c, authCallback = _b.authCallback, _d = _b.forceRedirectStrategy, forceRedirectStrategy = _d === void 0 ? false : _d, prompt = _b.prompt;
            if (msalInstance) {
                _this.log("Login STARTED");
                if (forceRedirectStrategy || _this.checkToIE()) {
                    _this.redirectLogin({ msalInstance: msalInstance, scopes: scopes, prompt: prompt });
                }
                else {
                    _this.popupLogin({
                        msalInstance: msalInstance,
                        scopes: scopes,
                        withUserData: withUserData,
                        authCallback: authCallback,
                        prompt: prompt,
                    });
                }
            }
            else {
                _this.log("Login FAILED", "clientID broken or not provided", true);
            }
        };
        var graphScopes = props.graphScopes, clientId = props.clientId, tenantUrl = props.tenantUrl, redirectUri = props.redirectUri;
        _this.state = {
            msalInstance: getUserAgentApp({ clientId: clientId, tenantUrl: tenantUrl, redirectUri: redirectUri }),
            scopes: getScopes(graphScopes),
        };
        return _this;
    }
    MicrosoftLogin.prototype.componentDidMount = function () {
        var _this = this;
        var _a = this.state, msalInstance = _a.msalInstance, scopes = _a.scopes;
        var _b = this.props, authCallback = _b.authCallback, _c = _b.withUserData, withUserData = _c === void 0 ? false : _c;
        if (!msalInstance) {
            this.log("Initialization", "clientID broken or not provided", true);
        }
        else {
            msalInstance.handleRedirectCallback(function (error, authResponse) {
                if (!error && authResponse) {
                    _this.log("Fetch Azure AD 'token' with redirect SUCCEDEED", authResponse);
                    _this.log("Fetch Graph API 'access_token' in silent mode STARTED");
                    _this.getGraphAPITokenAndUser({
                        msalInstance: msalInstance,
                        scopes: scopes,
                        withUserData: withUserData,
                        authCallback: authCallback,
                        isRedirect: true,
                    });
                }
                else {
                    _this.log("Fetch Azure AD 'token' with redirect FAILED", error, true);
                    authCallback(error);
                }
            });
        }
    };
    MicrosoftLogin.prototype.componentDidUpdate = function (prevProps) {
        var _a = this.props, clientId = _a.clientId, tenantUrl = _a.tenantUrl, redirectUri = _a.redirectUri;
        if (prevProps.clientId !== clientId ||
            prevProps.tenantUrl !== tenantUrl ||
            prevProps.redirectUri !== redirectUri) {
            this.setState({
                msalInstance: getUserAgentApp({ clientId: clientId, tenantUrl: tenantUrl, redirectUri: redirectUri }),
            });
        }
    };
    MicrosoftLogin.prototype.getGraphAPITokenAndUser = function (_a) {
        var _this = this;
        var msalInstance = _a.msalInstance, scopes = _a.scopes, withUserData = _a.withUserData, authCallback = _a.authCallback, isRedirect = _a.isRedirect;
        return msalInstance
            .acquireTokenSilent({ scopes: scopes })
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
        });
    };
    MicrosoftLogin.prototype.popupLogin = function (_a) {
        var _this = this;
        var msalInstance = _a.msalInstance, scopes = _a.scopes, withUserData = _a.withUserData, authCallback = _a.authCallback, prompt = _a.prompt;
        this.log("Fetch Azure AD 'token' with popup STARTED");
        msalInstance
            .loginPopup({ scopes: scopes, prompt: prompt })
            .then(function (authResponse) {
            _this.log("Fetch Azure AD 'token' with popup SUCCEDEED", authResponse);
            _this.log("Fetch Graph API 'access_token' in silent mode STARTED");
            _this.getGraphAPITokenAndUser({
                msalInstance: msalInstance,
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
    MicrosoftLogin.prototype.redirectLogin = function (_a) {
        var msalInstance = _a.msalInstance, scopes = _a.scopes, prompt = _a.prompt;
        this.log("Fetch Azure AD 'token' with redirect STARTED");
        msalInstance.loginRedirect({ scopes: scopes, prompt: prompt });
    };
    MicrosoftLogin.prototype.getUserData = function (authResponseWithAccessToken) {
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
            authCallback(undefined, __assign(__assign({}, userData), authResponseWithAccessToken));
        });
    };
    MicrosoftLogin.prototype.checkToIE = function () {
        var ua = window.navigator.userAgent;
        var msie = ua.indexOf("MSIE ");
        var msie11 = ua.indexOf("Trident/");
        var msedge = ua.indexOf("Edge/");
        var isIE = msie > 0 || msie11 > 0;
        var isEdge = msedge > 0;
        return isIE || isEdge;
    };
    MicrosoftLogin.prototype.log = function (name, content, isError) {
        var debug = this.props.debug;
        if (debug) {
            var style = "background-color: " + (isError ? "#990000" : "#009900") + "; color: #ffffff; font-weight: 700; padding: 2px";
            console.groupCollapsed("MSLogin debug");
            console.log("%c" + name, style);
            content && console.log(content.message || content);
            console.groupEnd();
        }
    };
    MicrosoftLogin.prototype.render = function () {
        var _a = this.props, buttonTheme = _a.buttonTheme, className = _a.className, children = _a.children;
        return children ? (React.createElement("div", { onClick: this.login }, children)) : (React.createElement(MicrosoftLoginButton_1.default, { buttonTheme: buttonTheme || "light", buttonClassName: className, onClick: this.login }));
    };
    return MicrosoftLogin;
}(React.Component));
exports.default = MicrosoftLogin;
//# sourceMappingURL=MicrosoftLogin.js.map