/*
 * MEP dashboards: team defaults.
 *
 * Fill these in once (or use Synchro settings > "Download team config" and save the
 * file over this one). Everyone who opens the dashboards from this folder then starts
 * connected, with no setup of their own. Personal changes made in Synchro settings are
 * kept in that person's browser and sit on top of these values.
 *
 * clientId      Client ID of the SPA app registered at developer.bentley.com > My Apps
 *               (redirect URIs http://localhost:8765/ and https://timecube.github.io/euston-mep-dashboards/). Registered 26 Nov 2026, owner
 *               Yusuf Ali. It is an identifier, not a secret: safe in a shared folder.
 * projectId     The Bentley Infrastructure Cloud project. Pick it by name in Synchro
 *               settings after signing in; there is no ID to look up.
 * scheduleId    The SYNCHRO 4D schedule inside that project. For v6.5 projects this is
 *               usually the same GUID as the project. Pick it in Synchro settings.
 */
window.MEP_DASHBOARD_CONFIG = {
  clientId: 'spa-4lcENJxTe2ZIDKaeNz02Pv9UR',
  webUrl: 'https://timecube.github.io/euston-mep-dashboards/',                    // the team's web address for these dashboards (https://.../). A downloaded copy opened from SharePoint jumps there to sign in.
  projectId: '',
  scheduleId: '',
  projectLabel: 'HS2 Euston',

  // What to read (Synchro settings fills these in; 'auto' makes a sensible guess)
  statusCategoryId: '',          // resource status category that holds Installed / Set Out / ...
  lengthFieldId: 'auto',         // user field holding the element length, or 'none'
  lengthUnit: 'auto',            // 'auto' | 'm' | 'mm' | 'ft'
  objectSource: 'auto',          // 'auto' | 'rules' | 'parent' | 'field:<userFieldId>'
  scopeMode: 'status',           // 'status' | 'descendants' | 'group' | 'name' | 'all'
  scopeValue: '',
  unrecordedStatus: 'Not Started',
  linkProgramme: true,           // read task dates so the insights page can draw the planned curve

  refreshMinutes: 10,            // live refresh of statuses
  fullReloadHours: 6             // full re-read of resources, lengths and programme links
};
