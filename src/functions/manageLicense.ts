import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

interface LicenseRequest {
  userEmail: string; // Partner's personal email (kept for validation/logging)
  entraUPN?: string; // Entra UPN if user already exists (e.g., analia.theillaud.talent@egg-events.com)
  entraObjectId?: string; // Entra Object ID if user already exists
  firstName: string;
  lastName: string;
  isActive: boolean;
}

// Office 365 E1 Group ID
const OFFICE365_E1_GROUP_ID = "1263971a-19e5-42d1-a25e-36d71ec76014";

export async function manageLicense(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("🔐 License management request received");

  try {
    // Parse request body
    const licenseRequest: LicenseRequest = await request.json() as LicenseRequest;

    // Validate required fields
    if (
      !licenseRequest.userEmail ||
      !licenseRequest.firstName ||
      !licenseRequest.lastName ||
      licenseRequest.isActive === undefined
    ) {
      return {
        status: 400,
        jsonBody: {
          success: false,
          error: "Missing required fields: userEmail, firstName, lastName, isActive",
        },
      };
    }

    context.log(`👤 Processing license for: ${licenseRequest.userEmail}`);
    context.log(`📧 Entra UPN: ${licenseRequest.entraUPN || 'Not provided'}`);
    context.log(`🆔 Entra Object ID: ${licenseRequest.entraObjectId || 'Not provided'}`);
    context.log(`📊 Active status: ${licenseRequest.isActive}`);

    // Initialize Microsoft Graph client
    const credential = new ClientSecretCredential(
      process.env.AZURE_TENANT_ID!,
      process.env.AZURE_CLIENT_ID!,
      process.env.AZURE_CLIENT_SECRET!
    );

    const graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await credential.getToken(
            "https://graph.microsoft.com/.default"
          );
          return token.token;
        },
      },
    });

    // Find or create user in Entra ID
    const user = await findOrCreateUser(
      graphClient,
      licenseRequest.entraObjectId, // Try Object ID first
      licenseRequest.entraUPN, // Then try UPN
      licenseRequest.userEmail, // Fallback to personal email
      licenseRequest.firstName,
      licenseRequest.lastName,
      context
    );

    if (!user || !user.id) {
      throw new Error("Failed to find or create user in Entra ID");
    }

    context.log(`✅ User identified: ${user.userPrincipalName} (${user.id})`);

    // Manage group membership based on isActive flag
    if (licenseRequest.isActive) {
      // Add user to group
      await addUserToGroup(graphClient, user.id, context);
    } else {
      // Remove user from group
      await removeUserFromGroup(graphClient, user.id, context);
    }

    return {
      status: 200,
      jsonBody: {
        success: true,
        action: licenseRequest.isActive ? "added" : "removed",
        message: licenseRequest.isActive 
          ? "User added to Office 365 E1 group"
          : "User removed from Office 365 E1 group",
        userPrincipalName: user.userPrincipalName,
        userId: user.id,
        timestamp: new Date().toISOString(),
      },
    };
  } catch (error: any) {
    context.error("❌ Error managing license:", error);

    const errorMessage = error.message || "Unknown error occurred";
    
    let statusCode = 500;
    let userMessage = "Failed to manage license";

    if (error.statusCode === 401) {
      statusCode = 401;
      userMessage = "Authentication failed. Please check app permissions.";
    } else if (error.statusCode === 403) {
      statusCode = 403;
      userMessage = "Insufficient permissions to manage licenses.";
    } else if (error.statusCode === 404) {
      statusCode = 404;
      userMessage = "User or group not found.";
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

// Find or create user in Entra ID
// UPDATED LOGIC: Try Object ID first, then UPN, then search/create by email
async function findOrCreateUser(
  graphClient: Client,
  entraObjectId: string | undefined,
  entraUPN: string | undefined,
  email: string,
  firstName: string,
  lastName: string,
  context: InvocationContext
): Promise<any> {
  try {
    // OPTION 1: If we have the Entra Object ID, use it directly (fastest & most reliable)
    if (entraObjectId) {
      context.log(`🔍 Looking up user by Entra Object ID: ${entraObjectId}`);
      try {
        const user = await graphClient
          .api(`/users/${entraObjectId}`)
          .get();
        
        context.log(`✅ Found user by Object ID: ${user.userPrincipalName}`);
        return user;
      } catch (error: any) {
        context.log(`⚠️ User not found by Object ID, will try other methods`);
      }
    }

    // OPTION 2: If we have the Entra UPN, use it (fast & reliable)
    if (entraUPN) {
      context.log(`🔍 Looking up user by Entra UPN: ${entraUPN}`);
      try {
        const user = await graphClient
          .api(`/users/${entraUPN}`)
          .get();
        
        context.log(`✅ Found user by UPN: ${user.userPrincipalName}`);
        return user;
      } catch (error: any) {
        context.log(`⚠️ User not found by UPN, will try searching`);
      }
    }

    // OPTION 3: Search by email (slower, but works for finding existing users)
    context.log(`🔍 Searching for user by email: ${email}`);
    
    // Try searching by mail or userPrincipalName
    const users = await graphClient
      .api("/users")
      .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
      .get();

    if (users.value && users.value.length > 0) {
      context.log(`✅ Found existing user: ${users.value[0].userPrincipalName}`);
      return users.value[0];
    }

    // OPTION 4: If user doesn't exist, create as guest user
    context.log(`➕ Creating new guest user with email: ${email}`);
    
    const invitation = await graphClient
      .api("/invitations")
      .post({
        invitedUserEmailAddress: email,
        invitedUserDisplayName: `${firstName} ${lastName}`,
        inviteRedirectUrl: "https://portal.azure.com",
        sendInvitationMessage: true,
        invitedUserMessageInfo: {
          customizedMessageBody: "You have been invited to access EGG Events resources. Please accept this invitation to activate your account.",
        },
      });

    context.log(`✅ Guest user created: ${invitation.invitedUser.id}`);
    
    return {
      id: invitation.invitedUser.id,
      userPrincipalName: invitation.invitedUserEmailAddress,
      mail: email,
      displayName: `${firstName} ${lastName}`,
    };
  } catch (error) {
    context.error("❌ Error finding/creating user:", error);
    throw error;
  }
}

// Add user to Office 365 E1 group
async function addUserToGroup(
  graphClient: Client,
  userId: string,
  context: InvocationContext
): Promise<void> {
  try {
    context.log(`🔄 Adding user to group (simplified method)`);
    
    // Try to add user directly without checking if already member
    await graphClient
      .api(`/groups/${OFFICE365_E1_GROUP_ID}/members/$ref`)
      .post({
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
      });

    context.log(`✅ User added to Office 365 E1 group`);
    
  } catch (error: any) {
    // If error is "already exists" - that's fine, user is already in group
    if (error.statusCode === 400 && 
        (error.message?.includes("already exist") || 
         error.message?.includes("already a member") ||
         error.code === "Request_BadRequest")) {
      context.log(`ℹ️ User is already a member of the group`);
      return;
    }
    
    // Any other error - throw it
    context.error("❌ Error adding user to group:", error);
    throw error;
  }
}


async function removeUserFromGroup(
  graphClient: Client,
  userId: string,
  context: InvocationContext
): Promise<void> {
  try {
    context.log(`🔄 Removing user from group (simplified method)`);
    
    // Try to remove user directly
    await graphClient
      .api(`/groups/${OFFICE365_E1_GROUP_ID}/members/${userId}/$ref`)
      .delete();

    context.log(`✅ User removed from Office 365 E1 group`);
    
  } catch (error: any) {
    // If error is "not found" - user wasn't in group anyway
    if (error.statusCode === 404) {
      context.log(`ℹ️ User was not a member of the group`);
      return;
    }
    
    // Any other error - throw it
    context.error("❌ Error removing user from group:", error);
    throw error;
  }
}

// Register the function
app.http('manageLicense', {
  methods: ['POST'],
  authLevel: 'function',
  handler: manageLicense
});