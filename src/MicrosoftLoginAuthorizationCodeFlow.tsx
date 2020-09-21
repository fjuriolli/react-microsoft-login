import * as React from "react";
import {
  PublicClientApplication,
  AuthenticationResult,
  AuthError,
} from "@azure/msal-browser";
import AbstractMicrosoftLogin, {
  UserAgentApp,
  GraphAPITokenAndUser,
  PopupLogin,
  CLIENT_ID_REGEX,
} from "./AbstractMicrosoftLogin";
import { MicrosoftLoginPrompt } from "../index";

import { GraphAPIUserData } from "../index";

const getUserAgentApp = ({
  clientId,
  tenantUrl,
  redirectUri,
}: UserAgentApp) => {
  if (clientId && CLIENT_ID_REGEX.test(clientId)) {
    return new PublicClientApplication({
      auth: {
        ...(redirectUri && { redirectUri }),
        ...(tenantUrl && { authority: tenantUrl }),
        clientId,
        navigateToLoginRequestUrl: false,
      },
    });
  }
};

export default class MicrosoftLoginAuthorizationCodeFlow extends AbstractMicrosoftLogin {
  constructor(props: any) {
    super(props);
  }

  private getMsalInstance(): PublicClientApplication {
    return super.getState().msalInstance as PublicClientApplication;
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

    msalInstance
      .handleRedirectPromise()
      .then((response) => {
        //handle redirect response

        // In case multiple accounts exist, you can select
        const currentAccounts = msalInstance.getAllAccounts();
        if (currentAccounts === null) {
          // no accounts detected
          console.log("Logged in, but no currentAccounts retrieved");
        } else if (currentAccounts.length > 1) {
          // Add choose account code here
          console.log(
            "Logged in, but many accounts retrieved",
            currentAccounts
          );
        } else if (currentAccounts.length === 1) {
          const username = currentAccounts[0].username;
          console.log(`Logged in with account ${username}.`);
        }
      })
      .catch((err) => {
        console.error("Error while redirecting user to login", err);
      });
  }

  protected loginRedirect({
    scopes,
    prompt,
  }: {
    scopes: string[];
    prompt?: MicrosoftLoginPrompt;
  }) {
    const msalInstance = this.getMsalInstance();
    msalInstance
      .loginPopup({ scopes, prompt })
      .then(function (loginResponse) {
        //login success

        // In case multiple accounts exist, you can select
        const currentAccounts = msalInstance.getAllAccounts();

        if (currentAccounts === null) {
          // no accounts detected
          console.log("no accounts detected");
        } else if (currentAccounts.length > 1) {
          // Add choose account code here
          console.log("Multiple accounts chosen", currentAccounts);
        } else if (currentAccounts.length === 1) {
          const username = currentAccounts[0].username;
          console.log(`Account ${username} chosen.`);
        }
      })
      .catch(function (error) {
        //login failure
        console.log(error);
      });
  }

  protected getGraphAPITokenAndUser({
    scopes,
    withUserData,
    authCallback,
    isRedirect,
  }: GraphAPITokenAndUser) {
    const msalInstance = this.getMsalInstance();
    return (
      msalInstance
        // TODO: https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/login-user.md#silent-login-with-ssosilent
        .ssoSilent({ scopes, loginHint: "get_from_username" })
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
        .then((authResponseWithAccessToken: AuthenticationResult) => {
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
        })
    );
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
      .then((authResult: AuthenticationResult) => {
        this.log("Fetch Azure AD 'token' with popup SUCCEDEED", authResult);
        this.log("Fetch Graph API 'access_token' in silent mode STARTED");
        this.getGraphAPITokenAndUser({
          scopes,
          withUserData,
          authCallback,
          isRedirect: false,
        });
      })
      .catch((error: any) => {
        this.log("Fetch Azure AD 'token' with popup FAILED", error, true);
        authCallback(error);
      });
  }

  protected getUserData(authResponseWithAccessToken: AuthenticationResult) {
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
        // authCallback(undefined, {
        //   ...userData,
        //   ...authResponseWithAccessToken,
        // });
        console.log("userData", userData);
      });
  }
}
