import './theme-toggle.js?v=20260728-2';
import './public-clock-context.js?v=20260802-1';
import './public-load-failure-guard.js?v=20260801-1';

export const firebaseConfig = {
  apiKey: "AIzaSyB4xdaxbkXDRILPe2nGZuGCS-PXf35bk3o",
  authDomain: "qrtimeclock-42764.firebaseapp.com",
  projectId: "qrtimeclock-42764",
  storageBucket: "qrtimeclock-42764.appspot.com",
  messagingSenderId: "232535382723",
  appId: "1:232535382723:web:9fe08f4961d87ba4062076"
};

export const appSettings = {
  companyName: "Chadwell",
  defaultAppUrl: "https://qrtimeclock-42764.web.app"
};

if (typeof window !== 'undefined') {
  queueMicrotask(() => {
    import('./stable-public-clock-handler.js?v=20260802-3').catch((error) => {
      console.warn('Stable public clock handler failed to load:', error.message);
    });
    import('./mobile-punch-editor-actions.js?v=20260802-1').catch((error) => {
      console.warn('Mobile punch editor actions failed to load:', error.message);
    });
    import('./manual-punch-agency-fix.js?v=20260630-1').catch((error) => {
      console.warn('Manual punch agency fix failed to load:', error.message);
    });
    import('./agency-export-saved-timesheet-fallback.js?v=20260706-1').catch((error) => {
      console.warn('Agency export saved-timesheet fallback failed to load:', error.message);
    });
    import('./timeclock-usability-guard.js?v=20260711-1').catch((error) => {
      console.warn('Timeclock usability guard failed to load:', error.message);
    });
    import('./lunch-labels.js?v=20260802-1').catch((error) => {
      console.warn('Lunch label update failed to load:', error.message);
    });
    import('./new-worker-first-punch-hotfix.js?v=20260801-2').catch((error) => {
      console.warn('New-worker first-punch hotfix failed to load:', error.message);
    });
    import('./name-only-worker-resolver.js?v=20260801-1').catch((error) => {
      console.warn('Name-only worker resolver failed to load:', error.message);
    });
    import('./punch-exceptions-dashboard-v2.js?v=20260717-2').catch((error) => {
      console.warn('Punch exception dashboard failed to load:', error.message);
    });
    import('./agency-export-dropdown-dedupe.js?v=20260717-1').catch((error) => {
      console.warn('Agency export dropdown dedupe failed to load:', error.message);
    });
    import('./temp-self-service-compat.js?v=20260726-1').catch((error) => {
      console.warn('Temp worker self-service compatibility failed to load:', error.message);
    });
    import('./name-only-time-lookup.js?v=20260802-3').catch((error) => {
      console.warn('Exact-name time lookup failed to load:', error.message);
    });
    import('./worker-hours-summary.js?v=20260802-1').catch((error) => {
      console.warn('Worker hours summary failed to load:', error.message);
    });
  });
}
