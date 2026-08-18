import c from "chalk";
import ical from "ical-generator";
import { config as dotenvConfig } from "dotenv";
import { mkdir } from "node:fs/promises";
import { getDateFromValue, getRawValue, isDate, setLkHour } from "./utils";
import { config } from "./config";
import { Workbook } from "exceljs";

dotenvConfig({ quiet: true });

type LastWeek = [Date, Date, Date, Date, Date] | undefined;
interface WeekData {
	startDateTime: Date;
	endDateTime: Date;
	event: string;
}

console.log(c.blue("Script made by @posandu"));

const fileData = await Bun.file(
	"./downloaded/" + config.fileNameWithExt,
).arrayBuffer();

const workbook = new Workbook();
await workbook.xlsx.load(fileData);

const worksheet = workbook.getWorksheet(config.worksheetName);
if (!worksheet) throw new Error("Worksheet not found");

const MODULE_NAME = worksheet
	.getCell(config.summaryCell)
	.text.replace("\n", "")
	.trim();

console.log(c.yellow(MODULE_NAME));

/**
 *
 *
 */

let weeks: WeekData[] = [];
let lastWeek: LastWeek;
let rowIndex = config.dataStartIndex;

while (true) {
	if (rowIndex > worksheet.rowCount) break;

	const isDateRow = config.weekDaysArr.every((col) => {
		const val = isDate(worksheet.getCell(rowIndex, col).value);
		return val;
	});

	if (isDateRow) {
		lastWeek = undefined;

		console.log(c.green("Date row found"));

		lastWeek = config.weekDaysArr.map((col) => {
			const date = getDateFromValue(worksheet.getCell(rowIndex, col).value);

			if (!date) throw new Error("Date not found");

			return date;
		}) as LastWeek;
	}

	/**
	 * current way to check for the time is to check whether the 1st time and 2nd slot are dates
	 */
	const firstTimeSlot = getDateFromValue(worksheet.getCell(rowIndex, 2).value);
	const secondTimeSlot = getDateFromValue(worksheet.getCell(rowIndex, 3).value);

	if (typeof firstTimeSlot === "object" && typeof secondTimeSlot === "object") {
		console.log(c.green("Time slot row found"), c.yellow(rowIndex));

		const startTime = firstTimeSlot.getUTCHours();
		const endTime = secondTimeSlot.getUTCHours();

		console.log(
			c.green("Start time:"),
			c.yellow(startTime),
			c.green("End time:"),
			c.yellow(endTime),
		);

		if (!lastWeek) throw new Error("Last week not found");

		const events = config.weekDaysArr
			.map((col, i) => {
				const event = worksheet.getCell(rowIndex, col);

				if (!event) return;

				return {
					week: i,
					event: getRawValue(event),
				};
			})
			.filter((i) => i !== undefined && i.event !== null);

		events.forEach((event) => {
			if (!event) return;
			if (typeof event.event !== "string")
				throw Error("PARSING ERROR - WRONG TYPE - Got " + event.event);

			const date = lastWeek![event.week];

			if (!date) throw new Error("Date not found");

			const startDate = setLkHour(date, startTime);
			const endDate = setLkHour(date, endTime);

			console.log(
				c.green("Start date:"),
				c.yellow(startDate),
				c.green("End date:"),
				c.yellow(endDate),
			);

			weeks.push({
				startDateTime: startDate,
				endDateTime: endDate,
				event: event.event,
			});
		});
	}

	rowIndex++;
}

weeks = weeks.sort(
	(a, b) => a.startDateTime.getTime() - b.startDateTime.getTime(),
);

type Week = (typeof weeks)[number];

function mergeConsecutiveWeeks(weeks: Week[]): Week[] {
	if (weeks.length <= 1) return weeks;

	const result: Week[] = [];
	let current = weeks[0]!;

	for (let i = 1; i < weeks.length; i++) {
		const next = weeks[i]!;

		if (
			current.endDateTime.toDateString() ===
				next.startDateTime.toDateString() &&
			current.endDateTime.getTime() === next.startDateTime.getTime() &&
			current.event === next.event
		) {
			current = { ...current, endDateTime: next.endDateTime };
		} else {
			result.push(current);
			current = next;
		}
	}

	result.push(current);
	return result;
}

weeks = mergeConsecutiveWeeks(weeks);

function eventId(start: Date, summary: string) {
	const digest = new Bun.CryptoHasher("sha1")
		.update(`${start.toISOString()}|${summary}`)
		.digest("hex");
	return `${digest}@nsbm-timetable`;
}

function eventBelongsToDegree(event: string, modules: readonly string[]) {
	const mentionsKnownModule = Object.keys(config.modules).some((code) =>
		event.includes(code),
	);
	if (!mentionsKnownModule) return true;
	return modules.some((code) => event.includes(code));
}

await mkdir("ics", { recursive: true });

for (const degree of config.degrees) {
	const modules =
		config.degreeModules[degree as keyof typeof config.degreeModules];
	if (!modules) throw new Error(`No modules configured for ${degree}`);

	const degreeWeeks = weeks.filter((week) =>
		eventBelongsToDegree(week.event, modules),
	);

	const calendar = ical({
		name: `${degree} Timetable`,
		timezone: "Asia/Colombo",
	});

	for (const week of degreeWeeks) {
		calendar.createEvent({
			id: eventId(week.startDateTime, week.event),
			start: week.startDateTime,
			end: week.endDateTime,
			summary: week.event,
			stamp: week.startDateTime,
		});
	}

	const path = `ics/${degree}.ics`;
	await Bun.write(path, calendar.toString());
	console.log(
		c.green("Wrote"),
		c.yellow(path),
		c.green(`(${degreeWeeks.length} events)`),
	);
}
