// Force redeploy - minor comment change

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
    date: { type: "date-time", name: "VTjRN" }
  }
});

export const FlairScheduleHelper = https.onRequest(
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

      // Validate date range
      if (start > end) {
        return res.status(400).json({
          message: "Start date must not be after end date.",
          eventId,
          startDate,
          endDate
        });
      }

      // Step 1: Get all rows and filter by eventId
      const allRows = await bfScheduleDatesTable.get();
      const matchingRows = allRows.filter(row => row.eventId === eventId);
      const startCount = matchingRows.length;

      const existingDates = new Set(
        matchingRows.map(row =>
          DateTime.fromISO(row.date).toISODate()
        )
      );

      // Step 2: Delete rows outside the range
      const deletePromises = matchingRows
        .filter(row => {
          const date = DateTime.fromISO(row.date).startOf("day");
          return date < start || date > end;
        })
        .map(row => bfScheduleDatesTable.delete(row.$rowID));

      // Step 3: Create missing rows
      const dateRange = [];
      for (let dt = start; dt <= end; dt = dt.plus({ days: 1 })) {
        dateRange.push(dt.toISODate());
      }

      const addPromises = dateRange
        .filter(date => !existingDates.has(date))
        .map(date =>
          bfScheduleDatesTable.addRow({
            eventId,
            date: DateTime.fromISO(date).toISO()
          })
        );

      await Promise.all([
        Promise.all(deletePromises),
        Promise.all(addPromises)
      ]);

      const rowsFinal = startCount - deletePromises.length + addPromises.length;

      res.status(200).json({
        message: "Sync complete",
        eventId,
        startDate,
        endDate,
        rowsAtStart: startCount,
        rowsDeleted: deletePromises.length,
        rowsAdded: addPromises.length,
        rowsFinal: rowsFinal
      });
    } catch (err) {
      // Always log full error for your reference
      console.error("❌ Uncaught server error:", err?.message || err);

      // Return a clean, predictable payload
      res.set("Content-Type", "application/json");
      res.status(500).send(JSON.stringify({
        error: "Unexpected server error"
      }));
    }
  }
);