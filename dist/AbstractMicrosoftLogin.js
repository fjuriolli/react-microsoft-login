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
Object.defineProperty(exports, "__esModule", { value: true });
exports.CLIENT_ID_REGEX = void 0;
var React = require("react");
var MicrosoftLoginButton_1 = require("./MicrosoftLoginButton");
var getScopes = function (graphScopes) {
    var scopes = graphScopes || [];
    if (!scopes.find(function (el) { return el.toLowerCase() === "user.read"; })) {
        scopes.push("user.read");
    }
    return scopes;
};
exports.CLIENT_ID_REGEX = /[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}/;
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
                    _this.redirectLogin({ scopes: scopes, prompt: prompt });
                }
                else {
                    _this.popupLogin({
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
            msalInstance: _this.createUserAgent({ clientId: clientId, tenantUrl: tenantUrl, redirectUri: redirectUri }),
            scopes: getScopes(graphScopes),
        };
        return _this;
    }
    MicrosoftLogin.prototype.getState = function () {
        return this.state;
    };
    MicrosoftLogin.prototype.componentDidMount = function () {
        if (!this.isUserAgentCreated()) {
            this.log("Initialization", "clientID broken or not provided", true);
        }
        else {
            this.handleRedirect();
        }
    };
    MicrosoftLogin.prototype.componentDidUpdate = function (prevProps) {
        var _a = this.props, clientId = _a.clientId, tenantUrl = _a.tenantUrl, redirectUri = _a.redirectUri;
        if (prevProps.clientId !== clientId ||
            prevProps.tenantUrl !== tenantUrl ||
            prevProps.redirectUri !== redirectUri) {
            this.setState({
                msalInstance: this.createUserAgent({
                    clientId: clientId,
                    tenantUrl: tenantUrl,
                    redirectUri: redirectUri,
                }),
            });
        }
    };
    MicrosoftLogin.prototype.redirectLogin = function (_a) {
        var scopes = _a.scopes, prompt = _a.prompt;
        this.log("Fetch Azure AD 'token' with redirect STARTED");
        this.loginRedirect({ scopes: scopes, prompt: prompt });
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
//# sourceMappingURL=AbstractMicrosoftLogin.js.map