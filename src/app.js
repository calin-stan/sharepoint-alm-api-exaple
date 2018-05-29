import React from "react";
import ReactDOM from "react-dom";
import AuthenticationContext from "adal-angular";
import Axios, * as axios from "axios";

class HelloMessage extends React.Component {
  constructor() {
    super();
    this.keyPrefix = 'SPFxAppInstall-';

    this.state = {
      token: '',
      manualToken: 'Input_token',
      tenantId: localStorage.getItem(`${this.keyPrefix}tenantId`) || '',
      tenantName: localStorage.getItem(`${this.keyPrefix}tenantName`) || '',
      siteCollectionURL: localStorage.getItem(`${this.keyPrefix}siteCollectionURL`) || '',
      appCatalogLocation: localStorage.getItem(`${this.keyPrefix}appCatalogLocation`) || '',
      appId: ''
    }

    const config = {
      clientId: "20f19a01-5f31-405e-9ff7-8034ec867d6f", // Your AAD Application ID
      redirectUri: `http://127.0.0.1:8080`, // Your web app URL
      instance: "https://login.microsoftonline.com/",
      tenant: this.state.tenantId,
      postLogoutRedirectUri: window.location.origin,
      cacheLocation: "localStorage", // enable this for IE, as sessionStorage does not work for localhost.
    };
    this.authContext = new AuthenticationContext(config);

    this.login = () => {
      this.authContext.login();
    }

    this.logout = () => {
      this.authContext.logOut();
    }

    this.getToken = () => {
      const app = this;
      this.authContext.acquireToken(
        `https://${this.state.tenantName}.sharepoint.com`,
        (errorDesc, token, error) => {
          if (error) {
            console.log('Error while obtaining the access token', error, errorDesc);
          } else if (token != null) {
            console.log('Access token:', token);

            app.setState({
              token: token
            });
          } else {
            console.log(error, token, errorDesc);
          }
        }
      );
    }

    let handleAxiosError = (error, errorSource) => {
      console.log(`[${errorSource}] error: `, error.response.data['odata.error'].message.value);
    }

    this.listApps = () => {
      return Axios({
        url: `${this.state.siteCollectionURL}/_api/web/${this.state.appCatalogLocation}/AvailableApps`,
        method: 'get',
        headers: {
          'Authorization': `Bearer ${this.state.token}`,
          'Accept': 'application/json;odata=nometadata'
        }
      })
        .then(response => {
          console.log('[listApps] response:', response);
          var apps = [];
          response.data.value.map(app => {
            apps.push({
              title: app.Title,
              id: app.ID
            });
          });
          console.table(apps);
          return apps;
        })
        .catch(err => handleAxiosError(err, 'listApps'));
    }

    this.getApp = (appId) => {
      Axios({
        url: `${this.state.siteCollectionURL}/_api/web/${this.state.appCatalogLocation}/AvailableApps/GetById('${appId}')`,
        method: 'get',
        headers: {
          'Authorization': `Bearer ${this.state.token}`,
          'Accept': 'application/json;odata=nometadata'
        }
      })
        .then(response => {
          console.log('[getApp] response:', response);
        })
        .catch(err => handleAxiosError(err, 'getApp'));
    }

    var addApp = (appPackageName) => {
      fetch(`/${appPackageName}`).then(f => f.arrayBuffer())
        .then(buffer => {
          Axios({
            url: `${this.state.siteCollectionURL}/_api/web/${this.state.appCatalogLocation}/Add(overwrite=true, url='${appPackageName}')`,
            method: 'post',
            headers: {
              'Authorization': `Bearer ${this.state.token}`,
              'Accept': 'application/json;odata=nometadata',
              'binaryStringRequestBody': true
            },
            data: buffer
          })
            .then(response => {
              console.log('[addApp] response:', response);
              this.deployApp(response.data.UniqueId)
            })
            .catch(err => console.log('[addApp]', err));
        });
    }

    this.addApps = () => {
      var appPackageNames = ['SPFx-App-1.sppkg', 'SPFx-App-2.sppkg'];
      appPackageNames.map(appPackageName => {
        addApp(appPackageName);
      })
    }

    this.deployApp = (appId) => {
      Axios({
        url: `${this.state.siteCollectionURL}/_api/web/${this.state.appCatalogLocation}/AvailableApps/GetById('${appId}')/Deploy`,
        method: 'post',
        headers: {
          'Authorization': `Bearer ${this.state.token}`,
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata;charset=utf-8'
        }
      })
        .then(response => {
          console.log('[deployApp] response:', response);
          this.installApp(appId)
        })
        .catch(err => handleAxiosError(err, 'deployApp'));
    }

    this.installApp = (appId) => {
      Axios({
        url: `${this.state.siteCollectionURL}/_api/web/${this.state.appCatalogLocation}/AvailableApps/GetById('${appId}')/Install`,
        method: 'post',
        headers: {
          'Authorization': `Bearer ${this.state.token}`,
          'Accept': 'application/json;odata=nometadata'
        }
      })
        .then(response => {
          console.log('[installApp] response:', response);
        })
        .catch(err => handleAxiosError(err, 'installApp'));
    }

    var uninstallApp = (appId) => {
      Axios({
        url: `${this.state.siteCollectionURL}/_api/web/${this.state.appCatalogLocation}/AvailableApps/GetById('${appId}')/Uninstall`,
        method: 'post',
        headers: {
          'Authorization': `Bearer ${this.state.token}`,
          'Accept': 'application/json;odata=nometadata'
        }
      })
        .then(response => {
          console.log('[uninstallApp] response:', response);
        })
        .catch(err => handleAxiosError(err, 'uninstallApp'));
    }

    this.uninstallApps = () => {
      this.listApps()
        .then(apps => {
          apps.map(app => {
            uninstallApp(app.id)
          });
        });
    }

    this.getContext = () => {
      Axios({
        url: `${this.state.siteCollectionURL}/_api/contextinfo`,
        method: 'post',
        headers: {
          'Authorization': `Bearer ${this.state.token}`,
          'Accept': 'application/json;odata=nometadata'
        }
      })
        .then(response => {
          console.log('[getContext] response:', response);
        })
        .catch(err => console.error('getContext', err));
    }

    this.confirmOptions = () => {
      localStorage.setItem(`${this.keyPrefix}tenantId`, this.state.tenantId);
      localStorage.setItem(`${this.keyPrefix}tenantName`, this.state.tenantName);
      localStorage.setItem(`${this.keyPrefix}siteCollectionURL`, this.state.siteCollectionURL);
      localStorage.setItem(`${this.keyPrefix}appCatalogLocation`, this.state.appCatalogLocation);
    }
  }

  componentDidMount() {
    this.authContext.handleWindowCallback();
    if (!this.authContext.getCachedUser()) {
      console.log('No logged in user.');
    } else {
      console.log('Logged in user:', this.authContext.getCachedUser());
      this.getToken();
    }
  }

  render() {
    return (
      <div>
        <h1 className='title'>SharePoint Framework app installer</h1>

        <h2>Step 1. Fill in with the appropriate data and press the "Confirm options" button</h2>
        <label htmlFor='tenantId'>Your tenant ID</label>
        <input
          placeholder='4452c95e-45a8-487c-bbf9-4ce7c8977337'
          name='tenantId'
          onChange={(e) => { this.setState({ tenantId: e.target.value }) }}
          value={this.state.tenantId} />
        <br />

        <label htmlFor='tenantName'>Your tenant name</label>
        <input
          placeholder='contoso'
          name='tenantName'
          onChange={(e) => { this.setState({ tenantName: e.target.value }) }}
          value={this.state.tenantName} />
        <br />

        <label htmlFor='siteCollectionURL'>Site collection absolute URL</label>
        <input
          placeholder='https://contoso.sharepoint.com/sites/SPFx-App'
          name='siteCollectionURL'
          onChange={(e) => { this.setState({ siteCollectionURL: e.target.value }) }}
          value={this.state.siteCollectionURL} />
        <br />

        <select onChange={(e) => { this.setState({ appCatalogLocation: e.target.value }) }}>
          <option selected={true} disabled={true} value="">App catalog location</option>
          <option value="sitecollectionappcatalog">Site collection</option>
          <option value="tenantappcatalog">Tenant</option>
        </select>
        <br />

        <button className='button' onClick={this.confirmOptions}>Confirm options</button>
        <br />
        <hr />

        <h2>Step 2. Log in with a tenant admin account</h2>
        <button className='button' onClick={this.login}>Login</button>
        <button className='button' onClick={this.logout}>Logout</button>
        <br />
        <hr />

        <h2>Step 3. Use the API</h2>
        <button className='button' onClick={this.addApps}>Install apps</button>
        <br />

        <button className='button' onClick={this.listApps}>List apps</button>
        <br />

        <button className='button' onClick={this.uninstallApps}>Uninstall apps</button>
        <br />

        <hr />
        <p>Get an app by ID</p>
        <label htmlFor='appId'>App ID</label>
        <input
          name='appId'
          onChange={(e) => { this.setState({ appId: e.target.value }) }}
          value={this.state.appId}
        />
        <br />
        <button className='button' onClick={e => {this.getApp(this.state.appId)}}>Get app</button>

        <hr />
        <p>Token: {this.state.token}</p>
      </div>
    );
  }
}

var mountNode = document.getElementById("app");
ReactDOM.render(<HelloMessage name="there" />, mountNode);
