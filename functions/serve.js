// serve.js — run this locally for testing

import express from "express";
import { mealsPivotHandler } from "./mealsPivot.js";

const app = express();
app.use(express.json());

app.post("/testPivot", mealsPivotHandler);

const port = 5001;
app.listen(port, () => {
  console.log(`🚀 Local server running at http://localhost:${port}`);
});