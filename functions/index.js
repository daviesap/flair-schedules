//index.js
import { https } from "firebase-functions/v2";
import { initializeApp } from "firebase-admin/app";
import { getFirestore } from "firebase-admin/firestore";
import { updateDatesHandler } from "./updateDates.js";

// Initialize Firebase Admin globally
initializeApp();
const db = getFirestore();

export const FlairScheduleHelper = https.onRequest(
  { region: "europe-west2", cors: true },
  async (req, res) => {
    const action = req.query.action;

    try {
      if (action === "updateDates") {
        return await updateDatesHandler(req, res, db);
      }

      return res.status(400).json({ error: "Action not recognised" });
    } catch (err) {
      console.error("❌ Uncaught server error:", err?.message || err);
      res.set("Content-Type", "application/json");
      res.status(500).send(JSON.stringify({
        error: "Unexpected server error"
      }));
    }
  }
);