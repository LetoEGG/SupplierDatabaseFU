import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import * as crypto from "crypto";

interface ResetPasswordRequest {
  userPrincipalName?: string; // Entra UPN (preferred)
  userId?: string; // Entra Object ID (alternative)
  firstName: string;
  lastName: string;
  personalEmail: string;
}

interface ResetPasswordResponse {
  success: boolean;
  newPassword?: string;
  userPrincipalName?: string;
  error?: string;
  details?: string;
  timestamp?: string;
}

// Configuration
const SUPPLIER_PORTAL_URL = "https://egginthenest.sharepoint.com/sites/SupplierPortal/SitePages/MyProfile.aspx";
const HELPDESK_EMAIL = "helpdesk_eu@egg-events.com";
const SERVICE_ACCOUNT_EMAIL = "noreply@egg-events.com";

/**
 * Azure Function: Reset Freelancer Password
 * Resets password for an existing freelancer and sends new welcome email
 */
export async function resetFreelancerPassword(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("🔄 Reset password request received");

  try {
    // Parse request body
    const resetRequest: ResetPasswordRequest = await request.json() as ResetPasswordRequest;

    // Validate required fields
    if (
      (!resetRequest.userPrincipalName && !resetRequest.userId) ||
      !resetRequest.firstName ||
      !resetRequest.lastName ||
      !resetRequest.personalEmail
    ) {
      return {
        status: 400,
        jsonBody: {
          success: false,
          error: "Missing required fields: (userPrincipalName OR userId), firstName, lastName, personalEmail",
        },
      };
    }

    context.log(`🔐 Resetting password for: ${resetRequest.firstName} ${resetRequest.lastName}`);
    context.log(`📧 Identifier: ${resetRequest.userPrincipalName || resetRequest.userId}`);

    // Initialize Microsoft Graph client
    const credential = new ClientSecretCredential(
      process.env.AZURE_TENANT_ID!,
      process.env.AZURE_CLIENT_ID!,
      process.env.AZURE_CLIENT_SECRET!
    );

    const graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await credential.getToken("https://graph.microsoft.com/.default");
          return token.token;
        },
      },
    });

    // Find user (prefer UPN, fallback to user ID)
    const userIdentifier = resetRequest.userPrincipalName || resetRequest.userId;
    context.log(`🔍 Looking up user: ${userIdentifier}`);

    let user: any;
    try {
      user = await graphClient
        .api(`/users/${userIdentifier}`)
        .select("id,userPrincipalName,displayName,givenName,surname")
        .get();
      context.log(`✅ Found user: ${user.userPrincipalName} (ID: ${user.id})`);
    } catch (error: any) {
      if (error.statusCode === 404) {
        return {
          status: 404,
          jsonBody: {
            success: false,
            error: "User not found",
            details: `No user found with identifier: ${userIdentifier}`,
            timestamp: new Date().toISOString(),
          },
        };
      }
      throw error;
    }

    // Generate new secure password
    const newPassword = generateSecurePassword();
    context.log("🔐 Generated new password");

    // Update user password
    context.log("🔄 Updating password in Entra ID...");
    await graphClient.api(`/users/${user.id}`).patch({
      passwordProfile: {
        forceChangePasswordNextSignIn: true,
        password: newPassword,
      },
    });
    context.log("✅ Password updated successfully");

    // Send password reset email
    context.log("📧 Sending password reset email...");
    await sendPasswordResetEmail(
      graphClient,
      resetRequest,
      user.userPrincipalName,
      newPassword,
      context
    );
    context.log("✅ Password reset email sent");

    // Return success response
    return {
      status: 200,
      jsonBody: {
        success: true,
        newPassword: newPassword, // Include for SharePoint update if needed
        userPrincipalName: user.userPrincipalName,
        timestamp: new Date().toISOString(),
      } as ResetPasswordResponse,
    };

  } catch (error: any) {
    context.log(`❌ Error resetting password: ${error.message}`);
    
    // Enhanced error handling
    let statusCode = 500;
    let userMessage = "An unexpected error occurred while resetting the password.";
    let errorMessage = error.message || "Unknown error";

    if (error.code === "Request_ResourceNotFound") {
      statusCode = 404;
      userMessage = "User not found. Please verify the user exists in Entra ID.";
    } else if (error.statusCode === 401 || error.code === "InvalidAuthenticationToken") {
      statusCode = 401;
      userMessage = "Authentication failed. Please check app permissions.";
    } else if (error.statusCode === 403) {
      statusCode = 403;
      userMessage = "Insufficient permissions to reset passwords.";
    } else if (error.statusCode === 429) {
      statusCode = 429;
      userMessage = "Rate limit exceeded. Please try again later.";
    }

    return {
      status: statusCode,
      jsonBody: {
        success: false,
        error: userMessage,
        details: errorMessage,
        timestamp: new Date().toISOString(),
      },
    };
  }
}

/**
 * Generate secure password
 */
function generateSecurePassword(): string {
  const length = 16;
  const uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const lowercase = "abcdefghijklmnopqrstuvwxyz";
  const numbers = "0123456789";
  const symbols = "!@#$%^&*";
  
  const allChars = uppercase + lowercase + numbers + symbols;
  
  // Ensure at least one character from each set
  let password = "";
  password += uppercase[crypto.randomInt(uppercase.length)];
  password += lowercase[crypto.randomInt(lowercase.length)];
  password += numbers[crypto.randomInt(numbers.length)];
  password += symbols[crypto.randomInt(symbols.length)];
  
  // Fill remaining length with random characters
  for (let i = password.length; i < length; i++) {
    password += allChars[crypto.randomInt(allChars.length)];
  }
  
  // Shuffle the password
  return password.split('').sort(() => crypto.randomInt(3) - 1).join('');
}

/**
 * Send password reset email with beautiful HTML template
 */
async function sendPasswordResetEmail(
  graphClient: Client,
  resetRequest: ResetPasswordRequest,
  upn: string,
  newPassword: string,
  context: InvocationContext
): Promise<void> {
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            line-height: 1.6; 
            color: #333; 
            margin: 0;
            padding: 0;
        }
        .container { 
            max-width: 600px; 
            margin: 0 auto; 
            padding: 20px; 
        }
        .header { 
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); 
            color: white; 
            padding: 30px; 
            text-align: center; 
            border-radius: 10px 10px 0 0; 
        }
        .header h1 {
            margin: 0 0 10px 0;
            font-size: 28px;
        }
        .content { 
            background: #f8f9fa; 
            padding: 30px; 
            border-radius: 0 0 10px 10px; 
        }
        .credential-box { 
            background: #fff; 
            border-left: 4px solid #dc3545; 
            padding: 20px; 
            margin: 20px 0; 
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .credential-box h3 {
            margin-top: 0;
            color: #dc3545;
        }
        .credential-box p {
            margin: 10px 0;
        }
        .credential-box code {
            background: #e9ecef;
            padding: 4px 8px;
            border-radius: 3px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            color: #d63384;
        }
        .button { 
            display: inline-block; 
            background: #dc3545; 
            color: white; 
            padding: 14px 28px; 
            text-decoration: none; 
            border-radius: 5px; 
            margin: 15px 0;
            font-weight: 600;
        }
        .button:hover {
            background: #c82333;
        }
        .warning { 
            background: #fff3cd; 
            border: 1px solid #ffeaa7; 
            padding: 15px; 
            border-radius: 5px; 
            margin: 15px 0; 
        }
        .info-box {
            background: #d1ecf1;
            border: 1px solid #bee5eb;
            padding: 15px;
            border-radius: 5px;
            margin: 15px 0;
        }
        .steps { 
            background: #e7f3ff; 
            padding: 20px; 
            border-radius: 5px; 
            margin: 15px 0; 
        }
        .steps h3 {
            margin-top: 0;
            color: #0078d4;
        }
        .steps ol {
            margin: 10px 0;
            padding-left: 20px;
        }
        .steps li {
            margin: 8px 0;
        }
        .footer { 
            text-align: center; 
            color: #666; 
            font-size: 12px; 
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #dee2e6;
        }
        .footer a {
            color: #007bff;
            text-decoration: none;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔑 Password Reset Successful</h1>
            <p>Hello ${resetRequest.firstName}!</p>
        </div>
        
        <div class="content">
            <p>Your EGG Events account password has been reset successfully. You can now access your account with the new credentials below.</p>
            
            <div class="credential-box">
                <h3>🔐 Your New Account Details</h3>
                <p><strong>Username:</strong> <code>${upn}</code></p>
                <p><strong>New Temporary Password:</strong> <code>${newPassword}</code></p>
            </div>

            <div class="warning">
                ⚠️ <strong>Important Security Notice:</strong> You must change this temporary password immediately upon first login. This is required for your account security.
            </div>

            <div class="info-box">
                ℹ️ <strong>Why was my password reset?</strong><br>
                Your password may have been reset because:
                <ul style="margin: 10px 0; padding-left: 20px;">
                    <li>You or an administrator requested a password reset</li>
                    <li>You forgot your password</li>
                    <li>Routine security maintenance</li>
                </ul>
            </div>
            
            <h3>🌐 Access Your Account</h3>
            <p>Click the button below to sign in with your new credentials:</p>
            <a href="https://outlook.office.com/" class="button">📧 Sign In to Outlook</a>
            
            <div class="steps">
                <h3>🔑 How to Change Your Password</h3>
                <ol>
                    <li>Click the "Sign In to Outlook" button above or go to <a href="https://office.com/signin">office.com/signin</a></li>
                    <li>Enter your username: <code>${upn}</code></li>
                    <li>Enter your temporary password: <code>${newPassword}</code></li>
                    <li>You'll be prompted to change your password immediately</li>
                    <li>Create a strong new password with:
                        <ul style="margin-top: 5px;">
                            <li>At least 8 characters</li>
                            <li>Uppercase and lowercase letters</li>
                            <li>Numbers and special symbols</li>
                        </ul>
                    </li>
                    <li>Confirm your new password and submit</li>
                </ol>
            </div>
            
            <h3>📋 Access Your Supplier Portal</h3>
            <p>After changing your password, you can access the Supplier Portal:</p>
            <a href="${SUPPLIER_PORTAL_URL}" class="button">🏢 Access Supplier Portal</a>
            
            <div class="warning">
                <strong>Note:</strong> The Supplier Portal requires your EGG account. If you have trouble accessing it, open a private/incognito browser window and sign in with your EGG credentials.
            </div>
            
            <div class="footer">
                <p>If you didn't request this password reset or have any concerns, please contact us immediately.</p>
                <p>📧 Contact us: <a href="mailto:${HELPDESK_EMAIL}">${HELPDESK_EMAIL}</a></p>
                <hr style="border: 0; border-top: 1px solid #dee2e6; margin: 15px 0;">
                <p>© ${new Date().getFullYear()} EGG Events - Account Security System</p>
            </div>
        </div>
    </div>
</body>
</html>
`;

  const message = {
    message: {
      subject: "🔑 Your EGG Events Password Has Been Reset",
      body: {
        contentType: "HTML",
        content: htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: resetRequest.personalEmail,
            name: `${resetRequest.firstName} ${resetRequest.lastName}`,
          },
        },
      ],
    },
    saveToSentItems: false,
  };

  await graphClient.api(`/users/${SERVICE_ACCOUNT_EMAIL}/sendMail`).post(message);
}

// Register the function
app.http("resetFreelancerPassword", {
  methods: ["POST"],
  authLevel: "function",
  handler: resetFreelancerPassword,
});