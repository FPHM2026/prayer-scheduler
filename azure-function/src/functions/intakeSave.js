const { app } = require("@azure/functions");
const { findItemByToken, updateItem } = require("../graphClient");

app.http("intakeSave", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "intake/save",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      if (!body.token) return { status: 400, jsonBody: { error: "token required" } };
      const item = await findItemByToken(body.token);
      if (!item) return { status: 404, jsonBody: { error: "not found" } };
      const responses = body.responses || {};
      const fields = { ResponsesJSON: JSON.stringify(responses) };
      if (responses.name) {
        fields.RecipientName = String(responses.name);
        fields.Title = String(responses.name);
      }
      if (responses.email) fields.RecipientEmail = String(responses.email);
      await updateItem(item.id, fields);
      return { jsonBody: { ok: true } };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
