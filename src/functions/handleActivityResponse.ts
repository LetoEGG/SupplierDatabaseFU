import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

/**
 * Request payload from Logic App actionable email webhook
 */
interface ActivityResponseRequest {
  response: string; // "yes" or "no" from email button
  supplierEmail: string; // Personal email
  supplierName: string; // Full name for logging
  firstName: string; // For personalized emails
  entraObjectId: string; // Entra Object ID (most reliable identifier)
  sharePointItemId: string; // SharePoint list item ID
  requestedByEmail?: string; // Manager email for notifications
}

/**
 * Response returned to Logic App
 */
interface ActivityResponseResult {
  success: boolean;
  action?: 'activity_confirmed' | 'license_removed';
  message?: string;
  error?: string;
  details?: string;
  timestamp?: string;
}

// Configuration
const OFFICE365_E1_GROUP_ID = "1263971a-19e5-42d1-a25e-36d71ec76014";
const SHAREPOINT_SITE_URL = "https://egginthenest.sharepoint.com/sites/INTRANET";
const SHAREPOINT_LIST_ID = "ebfc366f-f3ca-4230-8f81-25d4e23d5f0b"; 
const SERVICE_ACCOUNT_EMAIL = "noreply@egg-events.com";

/**
 * Azure Function: Handle Activity Check Response
 * Processes webhook callbacks from actionable email buttons (Yes/No)
 * Updates SharePoint and manages E1 license group membership accordingly
 */
export async function handleActivityResponse(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("📥 Activity response webhook received");

  try {
    // Parse request body
    const activityRequest: ActivityResponseRequest = await request.json() as ActivityResponseRequest;

    // Validate required fields
    if (
      !activityRequest.response ||
      !activityRequest.supplierEmail ||
      !activityRequest.entraObjectId ||
      !activityRequest.sharePointItemId
    ) {
      return {
        status: 400,
        jsonBody: {
          success: false,
          error: "Missing required fields: response, supplierEmail, entraObjectId, sharePointItemId",
        } as ActivityResponseResult,
      };
    }

    const response = activityRequest.response.toLowerCase();
    context.log(`👤 Processing ${response.toUpperCase()} response for: ${activityRequest.supplierName}`);
    context.log(`📧 Email: ${activityRequest.supplierEmail}`);
    context.log(`🆔 Entra Object ID: ${activityRequest.entraObjectId}`);
    context.log(`📄 SharePoint Item ID: ${activityRequest.sharePointItemId}`);

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

    const timestamp = new Date().toISOString();

    // Get SharePoint site ID
    const siteId = await getSharePointSiteId(graphClient, context);

    if (response === 'yes') {
      // User confirmed they're still active
      context.log("✅ User confirmed activity - updating SharePoint");

      await updateSharePointItem(
        graphClient,
        siteId,
        activityRequest.sharePointItemId,
        {
          supplierdb_LastActivityCheck: timestamp,
          supplierdb_ActivityReminderCount: 0,
          supplierdb_ActivityCheckStatus: "Active"
        },
        context
      );

      // Send confirmation email
      await sendConfirmationEmail(
        graphClient,
        activityRequest.supplierEmail,
        activityRequest.firstName,
        context
      );

      context.log("✅ Activity confirmation processed successfully");

      return {
        status: 200,
        jsonBody: {
          success: true,
          action: 'activity_confirmed',
          message: 'Activity confirmed and SharePoint updated',
          timestamp: timestamp
        } as ActivityResponseResult,
      };

    } else if (response === 'no') {
      // User declined - remove E1 license
      context.log("🚫 User declined - removing E1 license");

      // Remove from E1 group
      await removeUserFromGroup(
        graphClient,
        activityRequest.entraObjectId,
        context
      );

      // Update SharePoint
      await updateSharePointItem(
        graphClient,
        siteId,
        activityRequest.sharePointItemId,
        {
          supplierdb_SendToEntra: false,
          supplierdb_ActivityCheckStatus: "License Removed - User Declined",
          supplierdb_LicenseRemovedDate: timestamp
        },
        context
      );

      // Notify manager if email provided
      if (activityRequest.requestedByEmail) {
        await sendManagerNotification(
          graphClient,
          activityRequest.requestedByEmail,
          activityRequest.supplierName,
          activityRequest.supplierEmail,
          context
        );
      }

      context.log("✅ License removal processed successfully");

      return {
        status: 200,
        jsonBody: {
          success: true,
          action: 'license_removed',
          message: 'License removed and manager notified',
          timestamp: timestamp
        } as ActivityResponseResult,
      };

    } else {
      // Invalid response value
      return {
        status: 400,
        jsonBody: {
          success: false,
          error: `Invalid response value: ${activityRequest.response}. Expected 'yes' or 'no'.`,
        } as ActivityResponseResult,
      };
    }

  } catch (error: any) {
    context.error("❌ Error processing activity response:", error);

    return {
      status: 500,
      jsonBody: {
        success: false,
        error: error.message || "Internal server error",
        details: error.toString(),
      } as ActivityResponseResult,
    };
  }
}

/**
 * Get SharePoint site ID from URL
 */
async function getSharePointSiteId(
  graphClient: Client,
  context: InvocationContext
): Promise<string> {
  try {
    const url = new URL(SHAREPOINT_SITE_URL);
    const hostname = url.hostname;
    const sitePath = url.pathname;

    context.log(`🔍 Looking up site: ${hostname}:${sitePath}`);

    const site = await graphClient
      .api(`/sites/${hostname}:${sitePath}`)
      .get();

    context.log(`✅ Found site ID: ${site.id}`);
    return site.id;

  } catch (error) {
    context.error("❌ Error getting SharePoint site ID:", error);
    throw error;
  }
}

/**
 * Update SharePoint list item with new field values
 */
async function updateSharePointItem(
  graphClient: Client,
  siteId: string,
  itemId: string,
  fields: Record<string, any>,
  context: InvocationContext
): Promise<void> {
  try {
    context.log(`📝 Updating SharePoint item ${itemId}`);

    await graphClient
      .api(`/sites/${siteId}/lists/${SHAREPOINT_LIST_ID}/items/${itemId}/fields`)
      .patch(fields);

    context.log(`✅ SharePoint item updated successfully`);

  } catch (error) {
    context.error("❌ Error updating SharePoint item:", error);
    throw error;
  }
}

/**
 * Remove user from Office 365 E1 license group
 */
async function removeUserFromGroup(
  graphClient: Client,
  userId: string,
  context: InvocationContext
): Promise<void> {
  try {
    context.log(`🔄 Removing user from E1 group: ${userId}`);

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

    context.error("❌ Error removing user from group:", error);
    throw error;
  }
}

/**
 * Send confirmation email to supplier after activity confirmation
 */
async function sendConfirmationEmail(
  graphClient: Client,
  recipientEmail: string,
  firstName: string,
  context: InvocationContext
): Promise<void> {
  try {
    context.log(`📧 Sending confirmation email to ${recipientEmail}`);

    const message = {
      subject: "Thank you - Activity Confirmed",
      body: {
        contentType: "HTML",
        content: `
          <p>Dear ${firstName},</p>
          <p>Thank you for confirming your activity status!</p>
          <p>Your access to <strong>egg</strong> resources has been maintained.</p>
          <p>You will receive another check-in request in approximately one month.</p>
          <p>If you have any questions, please contact: OnBoard-Talent@egg-events.com</p>
          <p>Best regards,<br>egg Talent Team</p>
        `
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipientEmail
          }
        }
      ]
    };

    await graphClient
      .api(`/users/${SERVICE_ACCOUNT_EMAIL}/sendMail`)
      .post({
        message: message,
        saveToSentItems: false
      });

    context.log(`✅ Confirmation email sent successfully`);

  } catch (error) {
    context.error("❌ Error sending confirmation email:", error);
    // Don't throw - email failure shouldn't fail the whole operation
  }
}

/**
 * Send notification to manager when license is removed
 */
async function sendManagerNotification(
  graphClient: Client,
  managerEmail: string,
  supplierName: string,
  supplierEmail: string,
  context: InvocationContext
): Promise<void> {
  try {
    context.log(`📧 Sending manager notification to ${managerEmail}`);

    const message = {
      subject: `Freelancer Account Deactivated - ${supplierName}`,
      body: {
        contentType: "HTML",
        content: `
          <p>The account for <strong>${supplierName}</strong> (${supplierEmail}) has been deactivated at their request.</p>
          <p><strong>Reason:</strong> User indicated they are no longer working with egg.</p>
          <p><strong>Action Taken:</strong> Office 365 E1 license has been removed. The user can still log in but will not have access to email or SharePoint.</p>
          <p>If this was done in error or you need to restore access, please contact IT helpdesk: helpdesk_eu@egg-events.com</p>
          <p>Best regards,<br>egg Talent Management System</p>
        `,
      },
      toRecipients: [
        {
          emailAddress: {
            address: managerEmail
          }
        }
      ],
      ccRecipients: [
        {
          emailAddress: {
            address: "OnBoard-Talent@egg-events.com"
          }
        }
      ],
      bccRecipients: [
        {
          emailAddress: {
            address: "dl_freelanceonboarding@egg-events.com"
          }
        }
      ],
      importance: "High"
    };

    await graphClient
      .api(`/users/${SERVICE_ACCOUNT_EMAIL}/sendMail`)
      .post({
        message: message,
        saveToSentItems: false
      });

    context.log(`✅ Manager notification sent successfully`);

  } catch (error) {
    context.error("❌ Error sending manager notification:", error);
    // Don't throw - email failure shouldn't fail the whole operation
  }
}

// Register the function
app.http('handleActivityResponse', {
  methods: ['POST'],
  authLevel: 'function',
  handler: handleActivityResponse
});