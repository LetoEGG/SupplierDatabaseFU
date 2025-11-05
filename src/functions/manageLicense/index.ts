// azure-functions/manageLicense/index.ts
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

interface LicenseRequest {
  userEmail: string;
  firstName: string;
  lastName: string;
  isActive: boolean;
}

// Office 365 E1 Group ID
const OFFICE365_E1_GROUP_ID = "1263971a-19e5-42d1-a25e-36d71ec76014";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  context.log("🔐 License management request received");

  try {
    // Validate request
    if (!req.body) {
      context.res = {
        status: 400,
        body: { success: false, error: "Request body is required" },
      };
      return;
    }

    const licenseRequest: LicenseRequest = req.body;

    // Validate required fields
    if (
      !licenseRequest.userEmail ||
      !licenseRequest.firstName ||
      !licenseRequest.lastName ||
      licenseRequest.isActive === undefined
    ) {
      context.res = {
        status: 400,
        body: {
          success: false,
          error: "Missing required fields: userEmail, firstName, lastName, isActive",
        },
      };
      return;
    }

    context.log(`👤 Processing license for: ${licenseRequest.userEmail}`);
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
      licenseRequest.userEmail,
      licenseRequest.firstName,
      licenseRequest.lastName,
      context
    );

    if (!user || !user.id) {
      throw new Error("Failed to find or create user in Entra ID");
    }

    context.log(`✅ User found/created: ${user.userPrincipalName} (${user.id})`);

    // Manage group membership based on isActive flag
    if (licenseRequest.isActive) {
      // Add user to group
      await addUserToGroup(graphClient, user.id, context);
    } else {
      // Remove user from group
      await removeUserFromGroup(graphClient, user.id, context);
    }

    context.res = {
      status: 200,
      body: {
        success: true,
        action: licenseRequest.isActive ? "added" : "removed",
        userEmail: user.userPrincipalName,
        userId: user.id,
        groupId: OFFICE365_E1_GROUP_ID,
        timestamp: new Date().toISOString(),
      },
    };
  } catch (error) {
    context.log.error("❌ Error managing license:", error);

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
    }

    context.res = {
      status: statusCode,
      body: {
        success: false,
        error: userMessage,
        details: errorMessage,
        timestamp: new Date().toISOString(),
      },
    };
  }
};

// Find existing user or create new guest user
async function findOrCreateUser(
  graphClient: Client,
  email: string,
  firstName: string,
  lastName: string,
  context: Context
): Promise<any> {
  try {
    // Try to find existing user by email
    context.log(`🔍 Searching for user: ${email}`);
    
    const users = await graphClient
      .api("/users")
      .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
      .select("id,userPrincipalName,mail,displayName")
      .get();

    if (users.value && users.value.length > 0) {
      context.log(`✅ Found existing user: ${users.value[0].userPrincipalName}`);
      return users.value[0];
    }

    // User not found, create as guest user
    context.log(`➕ Creating new guest user: ${email}`);
    
    const invitation = await graphClient
      .api("/invitations")
      .post({
        invitedUserEmailAddress: email,
        invitedUserDisplayName: `${firstName} ${lastName}`,
        inviteRedirectUrl: "https://myapps.microsoft.com",
        sendInvitationMessage: true,
        invitedUserMessageInfo: {
          customizedMessageBody: "You have been granted access to EGG Events systems. Please accept this invitation to activate your account.",
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
    context.log.error("❌ Error finding/creating user:", error);
    throw error;
  }
}

// Add user to Office 365 E1 group
async function addUserToGroup(
  graphClient: Client,
  userId: string,
  context: Context
): Promise<void> {
  try {
    // Check if user is already a member
    const members = await graphClient
      .api(`/groups/${OFFICE365_E1_GROUP_ID}/members`)
      .filter(`id eq '${userId}'`)
      .get();

    if (members.value && members.value.length > 0) {
      context.log(`ℹ️ User is already a member of the group`);
      return;
    }

    // Add user to group
    await graphClient
      .api(`/groups/${OFFICE365_E1_GROUP_ID}/members/$ref`)
      .post({
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
      });

    context.log(`✅ User added to Office 365 E1 group`);
  } catch (error) {
    // If error is "One or more added object references already exist"
    if (error.statusCode === 400 && error.message?.includes("already exist")) {
      context.log(`ℹ️ User is already a member of the group`);
      return;
    }
    
    context.log.error("❌ Error adding user to group:", error);
    throw error;
  }
}

// Remove user from Office 365 E1 group
async function removeUserFromGroup(
  graphClient: Client,
  userId: string,
  context: Context
): Promise<void> {
  try {
    // Check if user is a member
    const members = await graphClient
      .api(`/groups/${OFFICE365_E1_GROUP_ID}/members`)
      .filter(`id eq '${userId}'`)
      .get();

    if (!members.value || members.value.length === 0) {
      context.log(`ℹ️ User is not a member of the group`);
      return;
    }

    // Remove user from group
    await graphClient
      .api(`/groups/${OFFICE365_E1_GROUP_ID}/members/${userId}/$ref`)
      .delete();

    context.log(`✅ User removed from Office 365 E1 group`);
  } catch (error) {
    // If error is "resource not found" (user not in group)
    if (error.statusCode === 404) {
      context.log(`ℹ️ User is not a member of the group`);
      return;
    }
    
    context.log.error("❌ Error removing user from group:", error);
    throw error;
  }
}

export default httpTrigger;