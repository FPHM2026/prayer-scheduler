const { app } = require("@azure/functions");
const crypto = require("crypto");
const { createItem } = require("../graphClient");

app.http("intakeStart", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "intake/start",
  handler: async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const token = crypto.randomUUID();
      const fields = {
        Title: body.recipientName || "New intake",
        Token: token,
        RecipientName: body.recipientName || "",
        RecipientEmail: body.recipientEmail || "",
        Status: "InProgress",
        ResponsesJSON: "{}"
      };
      await createItem(fields);
      return { jsonBody: { token } };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
