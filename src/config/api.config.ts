// Oracle Fusion REST API
export const ORACLE_FUSION_CONFIG = {
  baseUrl:      '',
  defaultLimit: 500,
};

// APEX Database (for auth and user management)
export const APEX_DB_CONFIG = {
  baseUrl: 'https://g15d6279501ae08-buimerc.adb.me-dubai-1.oraclecloudapps.com/ords/bcldifc/reerp',
  endpoints: {
    glBalances: 'gl/trialbalance',
  },
};

// Oracle BI Publisher SOAP (for GL Balances BIP report)
export const ORACLE_SOAP_CONFIG = {
  prod: {
    baseUrl:  '',
    username: 'ratheesh@buimerccorp.com',
    password: 'BCL#261285',
  },
  test: {
    baseUrl:  '',
    username: 'javeedindia@gmail.com',
    password: 'Bumeric2026',
  },
  reports: {
    glBalances: '/Custom/FA_REPORTS/GL_REPORTS/GL_BALANCES_BIP.xdo',
  },
};
