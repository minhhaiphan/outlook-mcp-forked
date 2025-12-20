/**
 * Create email draft functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

// Direct TokenStorage import to bypass caching issues
const TokenStorage = require('../auth/token-storage');

/**
 * Create email draft handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateDraft(args) {
  console.log('=== CREATE DRAFT HANDLER ===');
  console.log('Args received:', JSON.stringify(args, null, 2));
  
  const { to, cc, bcc, subject, body, importance = 'normal' } = args || {};
  console.log('Extracted:', { to, cc, bcc, subject, body, importance });
  // At least one of subject or body should be provided for a meaningful draft
  if (!subject && !body) {
    return {
      content: [{ 
        type: "text", 
        text: "Either subject or body content is required to create a draft."
      }]
    };
  }
  
  try {
    // Get access token with fallback to direct TokenStorage
    let accessToken;
    
    try {
      console.log('[CREATE-DRAFT] Attempting ensureAuthenticated()...');
      accessToken = await ensureAuthenticated();
      console.log('[CREATE-DRAFT] ensureAuthenticated() succeeded');
    } catch (authError) {
      console.log('[CREATE-DRAFT] ensureAuthenticated() failed, trying direct TokenStorage...', authError.message);
      
      // Fallback: Use TokenStorage directly (bypasses caching issues)
      const tokenStorage = new TokenStorage({
        tokenStorePath: config.AUTH_CONFIG.tokenStorePath,
        clientId: config.AUTH_CONFIG.clientId,
        clientSecret: config.AUTH_CONFIG.clientSecret,
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        scopes: config.AUTH_CONFIG.scopes,
        redirectUri: config.AUTH_CONFIG.redirectUri,
      });
      
      accessToken = await tokenStorage.getValidAccessToken();
      
      if (!accessToken) {
        throw new Error('Authentication required');
      }
      console.log('[CREATE-DRAFT] Direct TokenStorage succeeded');
    }
    
    // Format recipients (optional for drafts)
    const toRecipients = to ? to.split(',').map(email => {
      email = email.trim();
      return {
        emailAddress: {
          address: email
        }
      };
    }) : [];
    
    const ccRecipients = cc ? cc.split(',').map(email => {
      email = email.trim();
      return {
        emailAddress: {
          address: email
        }
      };
    }) : [];
    
    const bccRecipients = bcc ? bcc.split(',').map(email => {
      email = email.trim();
      return {
        emailAddress: {
          address: email
        }
      };
    }) : [];
    
    // Prepare draft email object
    const draftObject = {
      subject: subject || '',
      body: {
        contentType: (body && body.includes('<html')) ? 'html' : 'Text',
        content: body || ''
      },
      importance,
      // isDraft: true
    };
   
    // Only add recipients if they exist
    if (toRecipients.length > 0) {
      draftObject.toRecipients = toRecipients;
    }
    
    if (ccRecipients.length > 0) {
      draftObject.ccRecipients = ccRecipients;
    }
    
    if (bccRecipients.length > 0) {
      draftObject.bccRecipients = bccRecipients;
    }
    
    // Make API call to create draft in drafts folder
    console.log('About to call Graph API with:', {
      method: 'POST',
      // endpoint: 'me/mailFolders/drafts/messages',
      endpoint: 'me/messages',
      payload: draftObject
    });
    
    // const result = await callGraphAPI(accessToken, 'POST', 'me/mailFolders/drafts/messages', draftObject);
    const result = await callGraphAPI(accessToken, 'POST', 'me/messages', draftObject);
    
    console.log('Graph API call completed successfully, result:', {
      id: result.id,
      subject: result.subject,
      isDraft: result.isDraft
    });
    
    // Format success message
    let recipientInfo = '';
    if (toRecipients.length > 0) {
      recipientInfo = `\nRecipients: ${toRecipients.length}${ccRecipients.length > 0 ? ` + ${ccRecipients.length} CC` : ''}${bccRecipients.length > 0 ? ` + ${bccRecipients.length} BCC` : ''}`;
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Draft created successfully!\n\nSubject: ${subject || '(no subject)'}${recipientInfo}\nDraft ID: ${result.id}\nMessage Length: ${body ? body.length : 0} characters\n\nThe draft has been saved in your Drafts folder and can be edited or sent later.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error creating draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateDraft;
