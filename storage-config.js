/*
  ADAPT Pipeline Tracker — Storage configuration
  ================================================
  This file controls where the tracker's data actually lives. The tracker
  code itself never changes — it always calls window.storage.get/set/delete/list.
  Swapping backends is just a change in THIS file.

  MODE 'local'
    Browser localStorage. Free, zero setup, works immediately.
    Limitation: data is only visible on this device/browser — not shared
    across the team. Good for testing the site before SharePoint is wired up.

  MODE 'sharepoint'
    Microsoft Graph API against a SharePoint list, so the whole team reads
    and writes the same data. Requires a one-time Entra ID (Azure AD) app
    registration by a Microsoft 365 admin — see SETUP-GRAPH.md for the exact
    steps. Once you have a Client ID, Tenant ID, Site ID and List ID, fill
    them in below and flip STORAGE_MODE to 'sharepoint'.
*/

const STORAGE_MODE = 'local'; // 'local' | 'sharepoint'

const GRAPH_CONFIG = {
  clientId: '',    // from the Entra ID app registration (Application (client) ID)
  tenantId: '',    // your Microsoft 365 tenant ID (or 'organizations')
  siteId: '',      // SharePoint site ID hosting the list (see SETUP-GRAPH.md to find it)
  listId: '',      // SharePoint list ID used as the data store
  redirectUri: window.location.origin + window.location.pathname
};

// ---------------------------------------------------------------------------
// Local backend — active by default, no setup required
// ---------------------------------------------------------------------------
const LocalBackend = {
  mode: 'local',
  async get(key) {
    try {
      const v = localStorage.getItem(key);
      return v !== null ? { key, value: v, shared: false } : null;
    } catch (e) {
      throw e;
    }
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
// SharePoint backend — stub. Complete this once GRAPH_CONFIG is filled in.
// See SETUP-GRAPH.md for:
//   - the Entra ID app registration steps
//   - the exact delegated permissions to request (Sites.ReadWrite.All)
//   - how to find your Site ID and List ID
//   - a worked example of the MSAL.js sign-in + Graph fetch calls that
//     belong in the methods below
// ---------------------------------------------------------------------------
const SharePointBackend = {
  mode: 'sharepoint',
  async get(key) {
    throw new Error('SharePoint backend not yet configured — see SETUP-GRAPH.md, then implement this method.');
  },
  async set(key, value) {
    throw new Error('SharePoint backend not yet configured — see SETUP-GRAPH.md, then implement this method.');
  },
  async delete(key) {
    throw new Error('SharePoint backend not yet configured — see SETUP-GRAPH.md, then implement this method.');
  },
  async list(prefix) {
    throw new Error('SharePoint backend not yet configured — see SETUP-GRAPH.md, then implement this method.');
  }
};

window.storage = STORAGE_MODE === 'sharepoint' ? SharePointBackend : LocalBackend;

// Status banner so nobody mistakes local-only storage for team-shared storage
window.addEventListener('DOMContentLoaded', function () {
  var el = document.getElementById('storage-mode-banner');
  if (!el) return;
  if (STORAGE_MODE === 'local') {
    el.innerHTML =
      '⚠️ Running on local browser storage — what you see here is only saved on this device, not shared with your team yet. ' +
      'See <code>SETUP-GRAPH.md</code> in this repo to connect a shared SharePoint list.';
    el.style.display = 'block';
  } else {
    el.style.display = 'none';
  }
});
