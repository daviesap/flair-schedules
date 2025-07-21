//updateDates.js
import { table } from "@glideapps/tables";
import { DateTime } from "luxon";

export async function updateDatesHandler(req, res, db) {
  const {
    appId,
    tableId,
    columnEventId,
    columnDateId,
    eventId,
    startDate,
    endDate
  } = req.body;


  if (!appId || !tableId || !columnEventId || !columnDateId || !eventId || !startDate || !endDate) {
    return res.status(400).json({ error: "Missing required parameters" });
  }

//Get Glide token from Firestore
  let token;
  try {
    const tokenDoc = await db.collection('glideTokens').doc(appId).get();
    if (!tokenDoc.exists) {
      return res.status(404).json({ error: `No token found for appId ${appId}` });
    }
    token = tokenDoc.data().token;
  } catch (err) {
    console.error("Error fetching token:", err);
    return res.status(500).json({ error: "Failed to retrieve token" });
  }

  // Configure the Glide table
  const bfScheduleDatesTable = table({
  token,
  app: appId,
  table: tableId,
  columns: {
    eventId: { type: "string", name: columnEventId },
    date: { type: "date-time", name: columnDateId }
  }
});


  const start = DateTime.fromISO(startDate).startOf("day");
  const end = DateTime.fromISO(endDate).startOf("day");

  if (start > end) {
    return res.status(400).json({
      message: "Start date must not be after end date.",
      eventId,
      startDate,
      endDate
    });
  }

  const allRows = await bfScheduleDatesTable.get();
  const matchingRows = allRows.filter(row => row.eventId === eventId);
  const startCount = matchingRows.length;

  const existingDates = new Set(
    matchingRows.map(row =>
      DateTime.fromISO(row.date).toISODate()
    )
  );

  const deletePromises = matchingRows
    .filter(row => {
      const date = DateTime.fromISO(row.date).startOf("day");
      return date < start || date > end;
    })
    .map(row => bfScheduleDatesTable.delete(row.$rowID));

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
    rowsFinal
  });
}