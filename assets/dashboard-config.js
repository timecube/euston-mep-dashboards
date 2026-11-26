/*
 * MEP Installation Progress: team setup.
 *
 * Everyone who opens https://timecube.github.io/euston-mep-dashboards/ gets these settings.
 * They sign in with their own Bentley account and the page finds the project by name, so
 * nobody fills in a form. Advanced options on the page can change things for one person;
 * to change them for everyone, edit this file and run tools/update-website.py.
 *
 * clientId      Bentley SPA app "HS2 Dashboard" (developer.bentley.com > My Apps), owner
 *               Yusuf Ali. An identifier, not a secret. Its redirect URIs include webUrl.
 * projectName   The SYNCHRO project (schedule) name exactly as SYNCHRO shows it.
 * projectId /   Leave empty to find the project by name. Fill them in only to pin a
 * scheduleId    specific schedule.
 * objectSource  'code:4D codes' reads the "4D codes" resource code (the coloured column).
 * includeTypes  Only these 4D codes are counted. Empty shows everything.
 */
window.MEP_DASHBOARD_CONFIG = {
  clientId: 'spa-4lcENJxTe2ZIDKaeNz02Pv9UR',
  webUrl: 'https://timecube.github.io/euston-mep-dashboards/',
  projectName: '4D MEP Test Revit Import',
  projectId: '',
  scheduleId: '',
  projectLabel: 'HS2 Euston',

  objectSource: 'code:4D codes',
  includeTypes: ['Cable Duct', 'Cable Duct Fitting'],
  statusCategoryId: '',          // empty: the category with "Install" in its name (Installation Status)
  lengthFieldId: 'auto',         // a field called Length
  lengthUnit: 'auto',
  scopeMode: 'all',
  scopeValue: '',
  unrecordedStatus: 'Not Started',
  linkProgramme: false,          // this page does not use programme dates; skipping them is faster

  refreshMinutes: 10,
  fullReloadHours: 6
};
