const { app } = require("@azure/functions");
const { findItemByToken } = require("../graphClient");

app.http("intakeLoad", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "intake/load",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      if (!body.token) return { status: 400, jsonBody: { error: "token required" } };
      const item = await findItemByToken(body.token);
      if (!item) return { status: 404, jsonBody: { error: "not found" } };
      let responses = {};
      try {
        responses = JSON.parse(item.fields.ResponsesJSON || "{}");
      } catch (e) {}
      return {
        jsonBody: {
          status: item.fields.Status || "InProgress",
          responses,
          recipientName: item.fields.RecipientName || "",
          submittedAt: item.fields.SubmittedAt || null
        }
      };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
