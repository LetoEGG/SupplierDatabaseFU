// azure-functions/sendEmail/index.ts
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

interface EmailRequest {
  recipients: string[];
  from: string;
  fromName: string;
  subject: string;
  body: string;
  isHtml: boolean;
}

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  context.log("📧 Email send request received");

  try {
    // Validate request
    if (!req.body) {
      context.res = {
        status: 400,
        body: { success: false, error: "Request body is required" },
      };
      return;
    }

    const emailRequest: EmailRequest = req.body;

    // Validate required fields
    if (
      !emailRequest.recipients ||
      emailRequest.recipients.length === 0 ||
      !emailRequest.subject ||
      !emailRequest.body ||
      !emailRequest.from
    ) {
      context.res = {
        status: 400,
        body: {
          success: false,
          error: "Missing required fields: recipients, from, subject, body",
        },
      };
      return;
    }

    // Validate recipients (max 50)
    if (emailRequest.recipients.length > 50) {
      context.res = {
        status: 400,
        body: {
          success: false,
          error: "Maximum 50 recipients allowed per email",
        },
      };
      return;
    }

    context.log(`📨 Sending email to ${emailRequest.recipients.length} recipients`);
    context.log(`📧 From: ${emailRequest.from} (${emailRequest.fromName})`);
    context.log(`📝 Subject: ${emailRequest.subject}`);

    // Initialize Microsoft Graph client with app-only authentication
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

    // Sanitize HTML body to prevent XSS
    const sanitizedBody = sanitizeHtml(emailRequest.body);

    // Prepare email message
    const message = {
      subject: emailRequest.subject.substring(0, 255), // Limit subject length
      body: {
        contentType: emailRequest.isHtml ? "HTML" : "Text",
        content: sanitizedBody,
      },
      from: {
        emailAddress: {
          address: emailRequest.from,
          name: emailRequest.fromName || "EGG Events",
        },
      },
      toRecipients: emailRequest.recipients.map((email) => ({
        emailAddress: {
          address: email,
        },
      })),
    };

    // Send email using Microsoft Graph API
    // Using the sendMail endpoint which sends from the specified user
    await graphClient
      .api(`/users/${emailRequest.from}/sendMail`)
      .post({
        message: message,
        saveToSentItems: true,
      });

    context.log(`✅ Email sent successfully to ${emailRequest.recipients.length} recipients`);

    context.res = {
      status: 200,
      body: {
        success: true,
        message: `Email sent to ${emailRequest.recipients.length} recipient(s)`,
        recipientCount: emailRequest.recipients.length,
        timestamp: new Date().toISOString(),
      },
    };
  } catch (error) {
    context.log.error("❌ Error sending email:", error);

    const errorMessage = error.message || "Unknown error occurred";
    
    // Check for specific Graph API errors
    let statusCode = 500;
    let userMessage = "Failed to send email";

    if (error.statusCode === 401) {
      statusCode = 401;
      userMessage = "Authentication failed. Please check app permissions.";
    } else if (error.statusCode === 403) {
      statusCode = 403;
      userMessage = "Insufficient permissions to send email.";
    } else if (error.statusCode === 404) {
      statusCode = 404;
      userMessage = "Sender email address not found.";
    } else if (error.statusCode === 429) {
      statusCode = 429;
      userMessage = "Rate limit exceeded. Please try again later.";
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

// Simple HTML sanitization to prevent XSS
function sanitizeHtml(html: string): string {
  if (!html) return "";

  // Remove potentially dangerous tags and attributes
  let sanitized = html
    // Remove script tags and content
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    // Remove event handlers
    .replace(/on\w+\s*=\s*["'][^"']*["']/gi, "")
    .replace(/on\w+\s*=\s*[^\s>]*/gi, "")
    // Remove javascript: protocols
    .replace(/javascript:/gi, "")
    // Remove data: protocols (except for images)
    .replace(/(<(?!img)[^>]*)\sdata:[^>]*>/gi, "$1>");

  return sanitized;
}

export default httpTrigger;