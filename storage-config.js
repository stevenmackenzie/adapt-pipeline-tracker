/*
  ADAPT Pipeline Tracker — Storage configuration
  ================================================
  STORAGE_MODE is now 'sharepoint'. Data is read from and written to the
  "ADAPT Tracker Data" SharePoint list, via Microsoft Graph, so everyone who
  opens this site with a digital.enterprises / adaptivesupply Microsoft 365
  account sees and edits the same data.

  First time a person opens the site, a Microsoft sign-in popup will appear.
  That's expected — it's how Graph knows who's asking and what they're
  allowed to see.
*/

const STORAGE_MODE = 'sharepoint'; // 'local' | 'sharepoint'

const GRAPH_CONFIG = {
  clientId: 'cf148c37-2f3b-4f93-9c9f-61518d1d369b',
  tenantId: 'fae45676-db94-4c5c-b601-20b526263541',
  siteId: 'adaptivesupply.sharepoint.com,5369a1a6-263c-4c5f-bd98-7eb5a22b0fea,4de1f333-d8d5-4d37-8f80-76a0e6167e2d',
  listId: '0d6b84e5-3f3c-4832-b7b6-eee7bcb3a91e',
  redirectUri: window.location.origin + window.location.pathname
};

// ---------------------------------------------------------------------------
// Local backend — kept here as a fallback / for local testing.
// ---------------------------------------------------------------------------
const LocalBackend = {
  mode: 'local',
  async get(key) {
    const v = localStorage.getItem(key);
    return v !== null ? { key, value: v, shared: false } : null;
  },
  async set(key, value) {
    localStorage.setItem(key, value);
    return { key, value, shared: false };
  },
  async delete(key) {
    localStorage.removeItem(key);
    return { key, deleted: true, shared: false };
  },
  async list(prefix) {
    const keys = Object.keys(localStorage).filter((k) => !prefix || k.indexOf(prefix) === 0);
    return { keys, prefix, shared: false };
  }
};

// ---------------------------------------------------------------------------
// SharePoint backend — reads/writes the "ADAPT Tracker Data" list via Graph.
// ---------------------------------------------------------------------------
let msalInstance = null;
let cachedAccount = null;

function getMsal() {
  if (!msalInstance) {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: GRAPH_CONFIG.clientId,
        authority: 'https://login.microsoftonline.com/' + GRAPH_CONFIG.tenantId,
        redirectUri: GRAPH_CONFIG.redirectUri
      },
      cache: {
        cacheLocation: 'localStorage' // survives page refreshes, so people aren't asked to sign in every time
      }
    });
  }
  return msalInstance;
}

async function getGraphToken() {
  const app = getMsal();
  const request = { scopes: ['Sites.ReadWrite.All'] };

  if (!cachedAccount) {
    const existing = app.getAllAccounts();
    if (existing.length > 0) cachedAccount = existing[0];
  }

  try {
    if (cachedAccount) {
      const result = await app.acquireTokenSilent({ ...request, account: cachedAccount });
      return result.accessToken;
    }
  } catch (e) {
    // fall through to interactive sign-in below
  }

  const result = await app.loginPopup(request);
  cachedAccount = result.account;
  return result.accessToken;
}

async function graphFetch(path, options) {
  const token = await getGraphToken();
  const res = await fetch('https://graph.microsoft.com/v1.0' + path, {
    ...options,
    headers: {
      ...(options && options.headers),
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json'
    }
  });
  if (!res.ok) {
    const errBody = await res.text();
    throw new Error('Graph API error ' + res.status + ': ' + errBody);
  }
  if (res.status === 204) return null;
  return res.json();
}

function listItemsPath() {
  return '/sites/' + GRAPH_CONFIG.siteId + '/lists/' + GRAPH_CONFIG.listId + '/items';
}

async function findItemByKey(key) {
  const data = await graphFetch(listItemsPath() + "?expand=fields&$filter=fields/Key eq '" + key + "'");
  if (data && data.value && data.value.length > 0) return data.value[0];
  return null;
}

const SharePointBackend = {
  mode: 'sharepoint',

  async get(key) {
    const item = await findItemByKey(key);
    if (!item) return null;
    return { key, value: item.fields.Value, shared: true };
  },

  async set(key, value) {
    const existing = await findItemByKey(key);
    if (existing) {
      await graphFetch(listItemsPath() + '/' + existing.id + '/fields', {
        method: 'PATCH',
        body: JSON.stringify({ Value: value })
      });
    } else {
      await graphFetch(listItemsPath(), {
        method: 'POST',
        body: JSON.stringify({ fields: { Title: key, Key: key, Value: value } })
      });
    }
    return { key, value, shared: true };
  },

  async delete(key) {
    const existing = await findItemByKey(key);
    if (existing) {
      await graphFetch(listItemsPath() + '/' + existing.id, { method: 'DELETE' });
    }
    return { key, deleted: true, shared: true };
  },

  async list(prefix) {
    const data = await graphFetch(listItemsPath() + '?expand=fields');
    const keys = (data && data.value ? data.value : [])
      .map((item) => item.fields.Key)
      .filter((k) => k && (!prefix || k.indexOf(prefix) === 0));
    return { keys, prefix, shared: true };
  }
};

window.storage = STORAGE_MODE === 'sharepoint' ? SharePointBackend : LocalBackend;

// Status banner
window.addEventListener('DOMContentLoaded', function () {
  var el = document.getElementById('storage-mode-banner');
  if (!el) return;
  if (STORAGE_MODE === 'local') {
    el.innerHTML =
      '⚠️ Running on local browser storage — what you see here is only saved on this device, not shared with your team yet. ' +
      'See <code>SETUP-GRAPH.md</code> in this repo to connect a shared SharePoint list.';
    el.style.display = 'block';
  } else {
    el.innerHTML =
      '🔗 Connected to shared SharePoint storage — changes here are visible to everyone on the team with access to this site.';
    el.style.display = 'block';
    el.style.background = '#E3F2E5';
    el.style.color = '#1B5E20';
    el.style.borderColor = '#A5D6A7';
  }
});
