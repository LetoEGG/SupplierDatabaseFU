import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import * as crypto from "crypto";

interface CreateFreelancerRequest {
  firstName: string;
  lastName: string;
  personalEmail: string;
  jobTitle?: string;
  mobilePhone?: string;
  city?: string;
  country?: string;
  companyName?: string;
}

interface CreateFreelancerResponse {
  success: boolean;
  userPrincipalName?: string;
  userId?: string;
  temporaryPassword?: string;
  error?: string;
  details?: string;
  timestamp?: string;
}

// Configuration
const TALENT_EMAIL_SUFFIX = ".talent@egg-events.com";
const SUPPLIER_PORTAL_URL = "https://egginthenest.sharepoint.com/sites/SupplierPortal/SitePages/MyProfile.aspx";
const HELPDESK_EMAIL = "helpdesk_eu@egg-events.com";
const SERVICE_ACCOUNT_EMAIL = "noreply@egg-events.com";

/**
 * Azure Function: Create Freelancer
 * Creates a new Entra ID user for a freelancer with all attributes and sends welcome email
 */
export async function createFreelancer(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("🆕 Create freelancer request received");

  try {
    // Parse request body
    const freelancerRequest: CreateFreelancerRequest = await request.json() as CreateFreelancerRequest;

    // Validate required fields
    if (
      !freelancerRequest.firstName ||
      !freelancerRequest.lastName ||
      !freelancerRequest.personalEmail
    ) {
      return {
        status: 400,
        jsonBody: {
          success: false,
          error: "Missing required fields: firstName, lastName, personalEmail",
        },
      };
    }

    context.log(`👤 Creating freelancer: ${freelancerRequest.firstName} ${freelancerRequest.lastName}`);
    context.log(`📧 Personal email: ${freelancerRequest.personalEmail}`);

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

    // Generate Entra UPN
    const upn = generateTalentUPN(freelancerRequest.firstName, freelancerRequest.lastName);
    context.log(`🔑 Generated UPN: ${upn}`);

    // Check if user already exists
    const existingUser = await checkUserExists(graphClient, upn, context);
    if (existingUser) {
      context.log(`⚠️ User already exists: ${upn}`);
      return {
        status: 409,
        jsonBody: {
          success: false,
          error: "User already exists",
          details: `A user with UPN ${upn} already exists in Entra ID`,
          userId: existingUser.id,
          userPrincipalName: existingUser.userPrincipalName,
          timestamp: new Date().toISOString(),
        },
      };
    }

    // Generate secure password
    const temporaryPassword = generateSecurePassword();
    context.log("🔐 Generated temporary password");

    // Create user object
    const userObject = {
      displayName: `${freelancerRequest.firstName} ${freelancerRequest.lastName}`,
      userPrincipalName: upn,
      mailNickname: normalizeName(`${freelancerRequest.firstName}${freelancerRequest.lastName}`),
      accountEnabled: true,
      passwordProfile: {
        forceChangePasswordNextSignIn: true,
        password: temporaryPassword,
      },
      givenName: freelancerRequest.firstName,
      surname: freelancerRequest.lastName,
      department: "Suppliers",
      mail: freelancerRequest.personalEmail, // Store personal email for reference
    };

    // Add optional fields
    if (freelancerRequest.jobTitle) {
      userObject['jobTitle'] = freelancerRequest.jobTitle;
    }
    if (freelancerRequest.mobilePhone) {
      userObject['mobilePhone'] = freelancerRequest.mobilePhone;
    }
    if (freelancerRequest.city) {
      userObject['city'] = freelancerRequest.city;
    }
    if (freelancerRequest.country) {
      userObject['country'] = freelancerRequest.country;
    }
    if (freelancerRequest.companyName) {
      userObject['companyName'] = freelancerRequest.companyName;
    }

    // Create user in Entra ID
    context.log("➕ Creating user in Entra ID...");
    const newUser = await graphClient.api("/users").post(userObject);
    context.log(`✅ User created: ${newUser.userPrincipalName} (ID: ${newUser.id})`);

    // Send welcome email
    context.log("📧 Sending welcome email...");
    await sendWelcomeEmail(
      graphClient,
      freelancerRequest,
      upn,
      temporaryPassword,
      context
    );
    context.log("✅ Welcome email sent");

    // Return success response
    return {
      status: 201,
      jsonBody: {
        success: true,
        userPrincipalName: newUser.userPrincipalName,
        userId: newUser.id,
        temporaryPassword: temporaryPassword, // Include for SharePoint update
        timestamp: new Date().toISOString(),
      } as CreateFreelancerResponse,
    };

  } catch (error: any) {
    context.log(`❌ Error creating freelancer: ${error.message}`);
    
    // Enhanced error handling
    let statusCode = 500;
    let userMessage = "An unexpected error occurred while creating the freelancer account.";
    let errorMessage = error.message || "Unknown error";

    if (error.code === "Request_ResourceNotFound") {
      statusCode = 404;
      userMessage = "Required Azure resource not found. Please check app configuration.";
    } else if (error.statusCode === 401 || error.code === "InvalidAuthenticationToken") {
      statusCode = 401;
      userMessage = "Authentication failed. Please check app permissions.";
    } else if (error.statusCode === 403) {
      statusCode = 403;
      userMessage = "Insufficient permissions to create users.";
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
 * Normalize name for UPN generation
 */
function normalizeName(name: string): string {
  if (!name) return "";

  let normalized = name.toLowerCase();

  // Replace accented characters
  const replacements: { [key: string]: string } = {
    'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
    'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
    'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
    'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
    'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u',
    'ç': 'c', 'ñ': 'n', 'ÿ': 'y'
  };

  for (const [char, replacement] of Object.entries(replacements)) {
    normalized = normalized.replace(new RegExp(char, 'g'), replacement);
  }

  // Remove any remaining non-alphanumeric characters except dots and hyphens
  normalized = normalized.replace(/[^a-z0-9.-]/g, '');

  return normalized;
}

/**
 * Generate talent UPN
 */
function generateTalentUPN(firstName: string, lastName: string): string {
  const firstNormalized = normalizeName(firstName);
  const lastNormalized = normalizeName(lastName);

  if (!firstNormalized || !lastNormalized) {
    throw new Error("Cannot generate UPN: Invalid first or last name");
  }

  return `${firstNormalized}.${lastNormalized}${TALENT_EMAIL_SUFFIX}`;
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
 * Check if user already exists
 */
async function checkUserExists(
  graphClient: Client,
  upn: string,
  context: InvocationContext
): Promise<any> {
  try {
    const user = await graphClient
      .api(`/users/${upn}`)
      .select("id,userPrincipalName,displayName")
      .get();
    return user;
  } catch (error: any) {
    if (error.statusCode === 404) {
      return null; // User doesn't exist
    }
    throw error;
  }
}

/**
 * Send welcome email with beautiful HTML template
 */
async function sendWelcomeEmail(
  graphClient: Client,
  freelancer: CreateFreelancerRequest,
  upn: string,
  password: string,
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
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
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
            border-left: 4px solid #007bff; 
            padding: 20px; 
            margin: 20px 0; 
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .credential-box h3 {
            margin-top: 0;
            color: #007bff;
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
            background: #007bff; 
            color: white; 
            padding: 14px 28px; 
            text-decoration: none; 
            border-radius: 5px; 
            margin: 15px 0;
            font-weight: 600;
        }
        .button:hover {
            background: #0056b3;
        }
        .warning { 
            background: #fff3cd; 
            border: 1px solid #ffeaa7; 
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
            <h1>🎉 Welcome to EGG Events!</h1>
            <p>Hello and welcome, ${freelancer.firstName}!</p>
        </div>
        
        <div class="content">
            <p>Your account as an <strong>EGG supplier</strong> has been created successfully. We're excited to have you on board!</p>
            
            <div class="credential-box">
                <h3>📧 Your Account Details</h3>
                <p><strong>Username:</strong> <code>${upn}</code></p>
                <p><strong>Temporary Password:</strong> <code>${password}</code></p>
            </div>

            <div class="warning">
                ⚠️ <strong>Important:</strong> This password must be changed immediately upon first login for security purposes.
            </div>
            
            <h3>🌐 Access Your Applications</h3>
            <p>We're using the Office 365 Suite. You can access your email and applications online:</p>
            <a href="https://outlook.office.com/" class="button">📧 Access Outlook Online</a>
            
            <p><strong>Note:</strong> Please use the web-based applications to access your EGG email and documents.</p>
            
            <div class="steps">
                <h3>🔑 How to Change Your Password</h3>
                <ol>
                    <li>Go to <a href="https://office.com/signin">office.com/signin</a></li>
                    <li>Sign in with your username: <code>${upn}</code></li>
                    <li>Enter your temporary password: <code>${password}</code></li>
                    <li>You'll be prompted to change your password immediately</li>
                    <li>Create a new strong password (at least 8 characters with uppercase, lowercase, numbers, and symbols)</li>
                    <li>Confirm your new password and submit</li>
                </ol>
            </div>
            
            <h3>📋 Complete Your Supplier Profile</h3>
            <p>To complete your supplier file and access additional resources, please visit:</p>
            <a href="${SUPPLIER_PORTAL_URL}" class="button">🏢 Access Supplier Portal</a>
            
            <div class="warning">
                <strong>Note:</strong> The Supplier Portal is only accessible with your EGG account. If you encounter any access issues, try opening the link in a private/incognito browser window and log in with your EGG credentials.
            </div>
            
            <p>📊 <strong>Monthly Activity Survey:</strong> You'll receive a monthly survey to confirm your continued collaboration with us. Please respond promptly to maintain your active status.</p>
            
            <div class="footer">
                <p>Welcome aboard! If you need any assistance, we're here to help.</p>
                <p>📧 Contact us: <a href="mailto:${HELPDESK_EMAIL}">${HELPDESK_EMAIL}</a></p>
                <hr style="border: 0; border-top: 1px solid #dee2e6; margin: 15px 0;">
                <p>© ${new Date().getFullYear()} EGG Events - Supplier Onboarding System</p>
            </div>
        </div>
    </div>
</body>
</html>
`;

  const message = {
    message: {
      subject: "🎉 Welcome to EGG Events - Your Account is Ready!",
      body: {
        contentType: "HTML",
        content: htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: freelancer.personalEmail,
            name: `${freelancer.firstName} ${freelancer.lastName}`,
          },
        },
      ],
    },
    saveToSentItems: false,
  };

  await graphClient.api(`/users/${SERVICE_ACCOUNT_EMAIL}/sendMail`).post(message);
}

// Register the function
app.http("createFreelancer", {
  methods: ["POST"],
  authLevel: "function",
  handler: createFreelancer,
});