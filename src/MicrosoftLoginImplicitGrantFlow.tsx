import * as React from "react";
import { UserAgentApplication, AuthResponse, AuthError } from "msal";
import AbstractMicrosoftLogin, {
  UserAgentApp,
  GraphAPITokenAndUser,
  PopupLogin,
  CLIENT_ID_REGEX,
} from "./AbstractMicrosoftLogin";
import { MicrosoftLoginPrompt } from "../index";

import {
  //   MicrosoftLoginProps,
  //   MicrosoftLoginState,
  GraphAPIUserData,
  //   MicrosoftLoginPrompt,
} from "../index";
// import MicrosoftLoginButton from "./MicrosoftLoginButton";
// import AbstractMicrosoftLogin from "./AbstractMicrosoftLogin";
//
// interface UserAgentApp {
//   clientId: string;
//   tenantUrl?: string;
//   redirectUri?: string;
// }
// interface GraphAPITokenAndUser {
//   msalInstance: UserAgentApplication;
//   scopes: string[];
//   withUserData: boolean;
//   authCallback: any;
//   isRedirect: boolean;
// }
// interface PopupLogin {
//   msalInstance: UserAgentApplication;
//   scopes: string[];
//   withUserData: boolean;
//   authCallback: any;
//   prompt?: MicrosoftLoginPrompt;
// }
// interface RedirectLogin {
//   msalInstance: UserAgentApplication;
//   scopes: string[];
//   prompt?: MicrosoftLoginPrompt;
// }

const getUserAgentApp = ({
  clientId,
  tenantUrl,
  redirectUri,
}: UserAgentApp) => {
  if (clientId && CLIENT_ID_REGEX.test(clientId)) {
    return new UserAgentApplication({
      auth: {
        ...(redirectUri && { redirectUri }),
        ...(tenantUrl && { authority: tenantUrl }),
        clientId,
        validateAuthority: true,
        navigateToLoginRequestUrl: false,
      },
    });
  }
};

export default class MicrosoftLoginImplicitGrantFlow extends AbstractMicrosoftLogin {
  constructor(props: any) {
    super(props);
  }

  private getMsalInstance(): UserAgentApplication {
    return super.getState().msalInstance as UserAgentApplication;
  }

  protected createUserAgent(userAgentApp: UserAgentApp): any {
    return getUserAgentApp(userAgentApp);
  }

  protected isUserAgentCreated(): boolean {
    return !!super.getState().msalInstance;
  }

  protected handleRedirect() {
    const msalInstance = this.getMsalInstance();
    const { scopes } = super.getState();
    const { authCallback, withUserData = false } = this.props;

    msalInstance.handleRedirectCallback(
      (error: AuthError, authResponse: AuthResponse) => {
        if (!error && authResponse) {
          super.log(
            "Fetch Azure AD 'token' with redirect SUCCEDEED",
            authResponse
          );
          super.log("Fetch Graph API 'access_token' in silent mode STARTED");
          this.getGraphAPITokenAndUser({
            scopes,
            withUserData,
            authCallback,
            isRedirect: true,
          });
        } else {
          this.log("Fetch Azure AD 'token' with redirect FAILED", error, true);
          authCallback(error);
        }
      }
    );
  }

  protected loginRedirect({
    scopes,
    prompt,
  }: {
    scopes: string[];
    prompt?: MicrosoftLoginPrompt;
  }) {
    this.getMsalInstance().loginRedirect({ scopes, prompt });
  }

  protected getGraphAPITokenAndUser({
    scopes,
    withUserData,
    authCallback,
    isRedirect,
  }: GraphAPITokenAndUser) {
    const msalInstance = this.getMsalInstance();
    return msalInstance
      .acquireTokenSilent({ scopes })
      .catch((error: any) => {
        this.log(
          "Fetch Graph API 'access_token' in silent mode is FAILED",
          error,
          true
        );
        if (isRedirect) {
          this.log("Fetch Graph API 'access_token' with redirect STARTED");
          msalInstance.acquireTokenRedirect({ scopes });
        } else {
          this.log("Fetch Graph API 'access_token' with popup STARTED");
          msalInstance.acquireTokenPopup({ scopes });
        }
      })
      .then((authResponseWithAccessToken: AuthResponse) => {
        this.log(
          "Fetch Graph API 'access_token' SUCCEDEED",
          authResponseWithAccessToken
        );
        if (withUserData) {
          this.getUserData(authResponseWithAccessToken);
        } else {
          this.log("Login SUCCEDED");
          authCallback(null, { authResponseWithAccessToken });
        }
      })
      .catch((error: AuthError) => {
        this.log("Login FAILED", error, true);
        authCallback(error);
      });
  }

  protected popupLogin({
    scopes,
    withUserData,
    authCallback,
    prompt,
  }: PopupLogin) {
    super.log("Fetch Azure AD 'token' with popup STARTED");
    const msalInstance = this.getMsalInstance();
    msalInstance
      .loginPopup({ scopes, prompt })
      .then((authResponse: AuthResponse) => {
        this.log("Fetch Azure AD 'token' with popup SUCCEDEED", authResponse);
        this.log("Fetch Graph API 'access_token' in silent mode STARTED");
        this.getGraphAPITokenAndUser({
          scopes,
          withUserData,
          authCallback,
          isRedirect: false,
        });
      })
      .catch((error: AuthError) => {
        this.log("Fetch Azure AD 'token' with popup FAILED", error, true);
        authCallback(error);
      });
  }

  protected getUserData(authResponseWithAccessToken: AuthResponse) {
    const { authCallback } = this.props;
    const { accessToken } = authResponseWithAccessToken;
    this.log("Fetch Graph API user data STARTED");
    const options = {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    };

    return fetch("https://graph.microsoft.com/v1.0/me", options)
      .then((response: Response) => response.json())
      .then((userData: GraphAPIUserData) => {
        this.log("Fetch Graph API user data SUCCEDEED", userData);
        this.log("Login SUCCEDED");
        authCallback(undefined, {
          ...userData,
          ...authResponseWithAccessToken,
        });
      });
  }
}
