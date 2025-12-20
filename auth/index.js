/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const { authTools } = require('./tools');

/**
 * Ensures the user is authenticated and returns an access token
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    // Force re-authentication
    throw new Error('Authentication required');
  }
  
  // Use the TokenStorage system that supports automatic refresh
  const TokenStorage = require('./token-storage');
  const config = require('../config');
  
  const tokenStorage = new TokenStorage({
    tokenStorePath: config.AUTH_CONFIG.tokenStorePath,
    clientId: config.AUTH_CONFIG.clientId,
    clientSecret: config.AUTH_CONFIG.clientSecret,
    tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    scopes: config.AUTH_CONFIG.scopes,
    redirectUri: config.AUTH_CONFIG.redirectUri,
  });
  
  console.log('[AUTH] ensureAuthenticated() called - using TokenStorage system');
  const accessToken = await tokenStorage.getValidAccessToken();
  
  if (!accessToken) {
    console.log('[AUTH] No valid access token available from TokenStorage');
    throw new Error('Authentication required');
  }
  
  console.log('[AUTH] Valid access token obtained from TokenStorage');
  return accessToken;
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated
};
