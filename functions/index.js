import { https } from "firebase-functions/v2";
import { initializeApp } from "firebase-admin/app";
import { table } from "@glideapps/tables";
import { DateTime } from "luxon";

initializeApp();

const bfScheduleDatesTable = table({
  token: "c5389e75-ed50-4e6c-b61d-3d94bfe8deaa",
  app: "rS9O2hVbqWGQrmriKHuJ",
  table: "native-table-m86V9FotCksCKNXcgxyx",
  columns: {
    eventId: { type: "string", name: "Name" },
    mainDate: { type: "date-time", name: "VTjRN" },
    mainDayDescription: { type: "string", name: "mgnzm" },
  },
});

export const ensureScheduleDates = https.onRequest(
  { region: "europe-west2", cors: true },
  async (req, res) => {
    const action = req.query.action;

    if (action !== "updateDates") {
      return res.status(400).json({ error: "Action not recognised" });
    }

    try {
      const { eventId, startDate, endDate } = req.body;

      if (!eventId || !startDate || !endDate) {
        return res.status(400).json({ error: "Missing required parameters" });
      }

      const start = DateTime.fromISO(startDate).startOf("day");
      const end = DateTime.fromISO(endDate).startOf("day");

      const dateRange = [];
      for (let dt = start; dt <= end; dt = dt.plus({ days: 1 })) {
        dateRange.push(dt.toISODate());
      }

      const allRows = await bfScheduleDatesTable.get();
      const matchingRows = allRows.filter(row => row.eventId === eventId);
      const existingDates = matchingRows.map(row =>
        DateTime.fromISO(row.mainDate).toISODate()
      );

      const missingDates = dateRange.filter(
        date => !existingDates.includes(date)
      );

      const insertPromises = missingDates.map(date => {
        return bfScheduleDatesTable.addRow({
          eventId,
          mainDate: DateTime.fromISO(date).toISO(),
          mainDayDescription: DateTime.fromISO(date).toFormat("cccc d LLLL yyyy"),
        });
      });

      await Promise.all(insertPromises);

      res.status(200).json({
        message: `Checked ${dateRange.length} days. Added ${missingDates.length} missing dates.`,
      });
    } catch (err) {
      console.error("❌ Error:", err);
      res.status(500).json({ error: err.message });
    }
  }
);