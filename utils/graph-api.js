/**
 * Microsoft Graph API helper functions
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');

const TokenStorage = require('../auth/token-storage'); // adjust path to where TokenStorage is
const tokenStorage = new TokenStorage({
  tokenStorePath: process.env.OUTLOOK_TOKEN_PATH || config.AUTH_CONFIG.tokenStorePath,
  clientId: config.AUTH_CONFIG.clientId,
  clientSecret: config.AUTH_CONFIG.clientSecret,
  tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  scopes: config.AUTH_CONFIG.scopes,
  redirectUri: config.AUTH_CONFIG.redirectUri,
});


/**
 * Makes a request to the Microsoft Graph API with automatic token refresh
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {string} path - API endpoint path
 * @param {object} data - Data to send for POST/PUT requests
 * @param {object} queryParams - Query parameters
 * @param {boolean} _isRetry - Internal flag to prevent infinite recursion
 * @returns {Promise<object>} - The API response
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, _isRetry = false) {
  // For test tokens, we'll simulate the API call
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating ${method} ${path} API call`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  try {
    console.error(`[GRAPH-API] Making real API call: ${method} ${path}`);
    if (data) {
      console.error(`[GRAPH-API] Request payload:`, JSON.stringify(data, null, 2));
    }
    
    // Check if path already contains the full URL (from nextLink)
    let finalUrl;
    if (path.startsWith('http://') || path.startsWith('https://')) {
      // Path is already a full URL (from pagination nextLink)
      finalUrl = path;
      console.error(`Using full URL from nextLink: ${finalUrl}`);
    } else {
      // Build URL from path and queryParams
      // Encode path segments properly
      const encodedPath = path.split('/')
        .map(segment => encodeURIComponent(segment))
        .join('/');
      
      // Build query string from parameters with special handling for OData filters
      let queryString = '';
      if (Object.keys(queryParams).length > 0) {
        // Handle $filter parameter specially to ensure proper URI encoding
        const filter = queryParams.$filter;
        if (filter) {
          delete queryParams.$filter; // Remove from regular params
        }
        
        // Build query string with proper encoding for regular params
        const params = new URLSearchParams();
        for (const [key, value] of Object.entries(queryParams)) {
          params.append(key, value);
        }
        
        queryString = params.toString();
        
        // Add filter parameter separately with proper encoding/
        if (filter) {
          if (queryString) {
            queryString += `&$filter=${encodeURIComponent(filter)}`;
          } else {
            queryString = `$filter=${encodeURIComponent(filter)}`;
          }
        }
        
        if (queryString) {
          queryString = '?' + queryString;
        }
        
        console.error(`Query string: ${queryString}`);
      }
      
      finalUrl = `${config.GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
      console.error(`Full URL: ${finalUrl}`);
    }
    
    return new Promise((resolve, reject) => {
      const options = {
        method: method,
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      };
      
      const req = https.request(finalUrl, options, (res) => {
        let responseData = '';
        
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        
        res.on('end', async () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            try {
              responseData = responseData ? responseData : '{}';
              const jsonResponse = JSON.parse(responseData);
              resolve(jsonResponse);
            } catch (error) {
              reject(new Error(`Error parsing API response: ${error.message}`));
            }
          } else if (res.statusCode === 401 && !_isRetry) {
            // Token expired or invalid - attempt refresh and retry once
            console.error('[AUTH] 401 from Graph, attempting refresh + retry once...');
            console.error("401 body:", responseData);
            
            try {
              console.error('[AUTH] Calling tokenStorage.getValidAccessToken() to refresh...');
              const fresh = await tokenStorage.getValidAccessToken(); // this triggers refresh if expired
              console.error(`[AUTH] getValidAccessToken() result: ${fresh ? 'SUCCESS - got fresh token' : 'FAILED - no token returned'}`);
              
              if (!fresh) {
                console.error('[AUTH] No fresh token available - refresh failed or no refresh_token');
                reject(new Error('UNAUTHORIZED: refresh failed or no refresh_token'));
                return;
              }
              
              console.error('[AUTH] Retrying API call with fresh token...');
              // Retry with fresh token
              const retryResult = await callGraphAPI(fresh, method, path, data, queryParams, true);
              console.error('[AUTH] Retry successful!');
              resolve(retryResult);
            } catch (refreshError) {
              console.error('[AUTH] Token refresh failed with error:', refreshError);
              console.error('[AUTH] Refresh error details:', {
                message: refreshError.message,
                stack: refreshError.stack,
                name: refreshError.name
              });
              const err = new Error(`UNAUTHORIZED: ${refreshError.message}`);
              err.statusCode = 401;
              err.body = responseData;
              reject(err);
            }
          } else {
            reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
          }
        });
      });
      
      req.on('error', (error) => {
        reject(new Error(`Network error during API call: ${error.message}`));
      });
      
      if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
        req.write(JSON.stringify(data));
      }
      
      req.end();
    });
  } catch (error) {
    console.error('Error calling Graph API:', error);
    
    // If this is a 401 error from the try block and we haven't retried yet, attempt refresh
    if (error.statusCode === 401 && !_isRetry) {
      console.error('[AUTH] 401 error caught, attempting refresh + retry once...');
      try {
        const fresh = await tokenStorage.getValidAccessToken();
        if (!fresh) throw new Error('UNAUTHORIZED: refresh failed or no refresh_token');
        return await callGraphAPI(fresh, method, path, data, queryParams, true);
      } catch (refreshError) {
        console.error('[AUTH] Token refresh failed:', refreshError);
        throw new Error('UNAUTHORIZED');
      }
    }
    
    throw error;
  }
}

/**
 * Calls Graph API with pagination support to retrieve all results up to maxCount
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET only for pagination)
 * @param {string} path - API endpoint path
 * @param {object} queryParams - Initial query parameters
 * @param {number} maxCount - Maximum number of items to retrieve (0 = all)
 * @returns {Promise<object>} - Combined API response with all items
 */
async function callGraphAPIPaginated(accessToken, method, path, queryParams = {}, maxCount = 0) {
  if (method !== 'GET') {
    throw new Error('Pagination only supports GET requests');
  }

  const allItems = [];
  let nextLink = null;
  let currentUrl = path;
  let currentParams = queryParams;

  try {
    do {
      // Make API call - now uses callGraphAPI with automatic refresh
      const response = await callGraphAPI(accessToken, method, currentUrl, null, currentParams);
      
      // Add items from this page
      if (response.value && Array.isArray(response.value)) {
        allItems.push(...response.value);
        console.error(`Pagination: Retrieved ${response.value.length} items, total so far: ${allItems.length}`);
      }

      // Check if we've reached the desired count
      if (maxCount > 0 && allItems.length >= maxCount) {
        console.error(`Pagination: Reached max count of ${maxCount}, stopping`);
        break;
      }

      // Get next page URL
      nextLink = response['@odata.nextLink'];
      
      if (nextLink) {
        // Pass the full nextLink URL directly to callGraphAPI
        currentUrl = nextLink;
        currentParams = {}; // nextLink already contains all params
        console.error(`Pagination: Following nextLink, ${allItems.length} items so far`);
      }
    } while (nextLink);

    // Trim to exact count if needed
    const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;

    console.error(`Pagination complete: Retrieved ${finalItems.length} total items`);
    
    return {
      value: finalItems,
      '@odata.count': finalItems.length
    };
  } catch (error) {
    console.error('Error during pagination:', error);
    throw error;
  }
}

module.exports = {
  callGraphAPI,
  callGraphAPIPaginated
};
