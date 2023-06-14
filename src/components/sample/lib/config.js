const config = {
  initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
  clientId: process.env.REACT_APP_CLIENT_ID,
  apiEndpoint: process.env.REACT_APP_FUNC_ENDPOINT,
  apiName: process.env.REACT_APP_FUNC_NAME,
  summaryEndpoint:process.env.SUMMARY_ENDPOINT,
  textAnalyticsEndpoint:process.env.TEXT_ANALYTICS_ENDPOINT,
  textAnalyticsApiKey:process.env.TEXT_ANALYTICS_API_KEY
};

export default config;
