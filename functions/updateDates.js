import { table } from "@glideapps/tables";
import { DateTime } from "luxon";

// Configure the Glide table
const bfScheduleDatesTable = table({
  token: "c5389e75-ed50-4e6c-b61d-3d94bfe8deaa",
  app: "rS9O2hVbqWGQrmriKHuJ",
  table: "native-table-m86V9FotCksCKNXcgxyx",
  columns: {
    eventId: { type: "string", name: "Name" },
    date: { type: "date-time", name: "VTjRN" }
  }
});

export async function updateDatesHandler(req, res) {
  const { eventId, startDate, endDate } = req.body;

  if (!eventId || !startDate || !endDate) {
    return res.status(400).json({ error: "Missing required parameters" });
  }

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
    message: "Success",
    eventId,
    startDate,
    endDate,
    rowsAtStart: startCount,
    rowsDeleted: deletePromises.length,
    rowsAdded: addPromises.length,
    rowsFinal
  });
}