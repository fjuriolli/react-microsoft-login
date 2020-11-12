import * as React from "react";

import {
  MicrosoftLoginProps,
  MicrosoftLoginState,
  MicrosoftLoginPrompt,
} from "../index";
import MicrosoftLoginButton from "./MicrosoftLoginButton";

export interface UserAgentApp {
  clientId: string;
  tenantUrl?: string;
  redirectUri?: string;
}

export interface GraphAPITokenAndUser {
  // msalInstance: UserAgentApplication;
  scopes: string[];
  withUserData: boolean;
  authCallback: any;
  isRedirect: boolean;
}

export interface PopupLogin {
  // msalInstance: UserAgentApplication;
  scopes: string[];
  withUserData: boolean;
  authCallback: any;
  prompt?: MicrosoftLoginPrompt;
}

export interface RedirectLogin {
  // msalInstance: UserAgentApplication;
  scopes: string[];
  prompt?: MicrosoftLoginPrompt;
}

const getScopes = (graphScopes: string[]) => {
  const scopes = graphScopes || [];
  if (!scopes.find((el: string) => el.toLowerCase() === "user.read")) {
    scopes.push("user.read");
  }
  return scopes;
};

export const CLIENT_ID_REGEX = /[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}/;

export default abstract class MicrosoftLogin extends React.Component<
  MicrosoftLoginProps,
  MicrosoftLoginState
> {
  protected constructor(props: any) {
    super(props);
    const { graphScopes, clientId, tenantUrl, redirectUri } = props;
    this.state = {
      msalInstance: this.createUserAgent({ clientId, tenantUrl, redirectUri }),
      scopes: getScopes(graphScopes),
    };
  }

  protected abstract isUserAgentCreated(): boolean;
  protected abstract createUserAgent(userAgentApp: UserAgentApp): any;
  protected abstract handleRedirect(): void;
  protected abstract loginRedirect(loginRedirect: any): void;
  protected abstract popupLogin({
    /*msalInstance,*/ scopes,
    withUserData,
    authCallback,
    prompt,
  }: PopupLogin): void;
  protected getState(): MicrosoftLoginState {
    return this.state;
  }

  componentDidMount() {
    // const { /*msalInstance,*/ scopes } = this.state;
    // const { authCallback, withUserData = false } = this.props;
    if (!this.isUserAgentCreated()) {
      this.log("Initialization", "clientID broken or not provided", true);
    } else {
      this.handleRedirect();
    }
  }

  componentDidUpdate(prevProps: any) {
    const { clientId, tenantUrl, redirectUri } = this.props;
    if (
      prevProps.clientId !== clientId ||
      prevProps.tenantUrl !== tenantUrl ||
      prevProps.redirectUri !== redirectUri
    ) {
      this.setState({
        msalInstance: this.createUserAgent({
          clientId,
          tenantUrl,
          redirectUri,
        }),
      });
    }
  }

  login = () => {
    const { msalInstance, scopes } = this.state;
    const {
      withUserData = false,
      authCallback,
      forceRedirectStrategy = false,
      prompt,
    } = this.props;

    if (msalInstance) {
      this.log("Login STARTED");
      if (forceRedirectStrategy || this.checkToIE()) {
        this.redirectLogin({ /*msalInstance,*/ scopes, prompt });
      } else {
        this.popupLogin({
          /*msalInstance,*/ scopes,
          withUserData,
          authCallback,
          prompt,
        });
      }
    } else {
      this.log("Login FAILED", "clientID broken or not provided", true);
    }
  };

  redirectLogin({ /*msalInstance,*/ scopes, prompt }: RedirectLogin) {
    this.log("Fetch Azure AD 'token' with redirect STARTED");
    this.loginRedirect({ scopes, prompt });
  }

  checkToIE(): boolean {
    const ua = window.navigator.userAgent;
    const msie = ua.indexOf("MSIE ");
    const msie11 = ua.indexOf("Trident/");
    const msedge = ua.indexOf("Edge/");
    const isIE = msie > 0 || msie11 > 0;
    const isEdge = msedge > 0;
    return isIE || isEdge;
  }

  log(name: string, content?: any, isError?: boolean) {
    const { debug } = this.props;
    if (debug) {
      const style = `background-color: ${
        isError ? "#990000" : "#009900"
      }; color: #ffffff; font-weight: 700; padding: 2px`;
      console.groupCollapsed("MSLogin debug");
      console.log(`%c${name}`, style);
      content && console.log(content.message || content);
      console.groupEnd();
    }
  }

  render() {
    const { buttonTheme, className, children } = this.props;
    return children ? (
      <div onClick={this.login}>{children}</div>
    ) : (
      <MicrosoftLoginButton
        buttonTheme={buttonTheme || "light"}
        buttonClassName={className}
        onClick={this.login}
      />
    );
  }
}
