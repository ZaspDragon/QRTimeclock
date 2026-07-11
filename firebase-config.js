export const firebaseConfig = {
  apiKey: "AIzaSyB4xdaxbkXDRILPe2nGZuGCS-PXf35bk3o",
  authDomain: "qrtimeclock-42764.firebaseapp.com",
  projectId: "qrtimeclock-42764",
  storageBucket: "qrtimeclock-42764.appspot.com",
  messagingSenderId: "232535382723",
  appId: "1:232535382723:web:9fe08f4961d87ba4062076"
};

export const appSettings = {
  companyName: "Chadwell",              // fallback if company doc not loaded
  defaultAppUrl: "https://qrtimeclock-42764.web.app"
};

if (typeof window !== 'undefined') {
  queueMicrotask(() => {
    import('./manual-punch-agency-fix.js?v=20260630-1').catch((error) => {
      console.warn('Manual punch agency fix failed to load:', error.message);
    });
    import('./agency-export-saved-timesheet-fallback.js?v=20260706-1').catch((error) => {
      console.warn('Agency export saved-timesheet fallback failed to load:', error.message);
    });
    import('./timeclock-usability-guard.js?v=20260711-1').catch((error) => {
      console.warn('Timeclock usability guard failed to load:', error.message);
    });
  });
}
