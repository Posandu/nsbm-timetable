import ExcelJS, { type CellSharedFormulaValue, type CellValue } from "exceljs";
import c from "chalk";
import ical from "ical-generator";
import { Octokit } from "@octokit/core";

const config = {
	name: "25.3-degree",
	worksheet: "25.3 Timetable",
	timeTableStart: 25,
	dataIndex: 2,
	summaryCell: "B4",
	timeSlotRegex: /(\d+)\.(\d+) ([a-zA-Z]+)\s?-\s?(\d+)\.(\d+) ([a-zA-Z]+)/,
	id: process.env.GIST_ID,
};

console.log(c.bgBlue.black("NSBM Timetable Converter"));
console.log(c.green("Found"), c.yellow(config.name));

console.log();

function getDateFromValue(val: CellValue): Date | undefined {
	if (val instanceof Date) return val;
	if (typeof val === "object" && val !== null) {
		if ((val as CellSharedFormulaValue)["result"] instanceof Date) {
			return (val as CellSharedFormulaValue)["result"] as Date;
		}
	}
	return undefined;
}

function isDate(val: CellValue): val is Date {
	return getDateFromValue(val) !== undefined;
}

function getHours(hours: number, sign: string) {
	sign = sign.trim().toUpperCase();

	if (sign === "PM") {
		if (hours < 12) return hours + 12;
		return hours;
	} else {
		if (hours === 12) return 0;
		return hours;
	}
}

console.log(c.blue("File:"), c.yellow(config.name));

if (!config.id) {
	console.log(c.red("Gist ID not found"));

	process.exit(1);
}

const timetable = await Bun.file(
	"./downloaded/" + config.name + ".xlsx"
).arrayBuffer();

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(timetable);

const worksheet = workbook.getWorksheet(config.worksheet);

if (!worksheet) throw new Error("Worksheet not found");

const MODULE_NAME = worksheet
	.getCell(config.summaryCell)
	.text.replace("Time Table - ", "")
	.trim();

console.log(c.blue("Module:"), c.yellow(MODULE_NAME));

let weeks: {
	start: Date;
	end: Date;
	name: string;
}[] = [];
let row = config.timeTableStart;
let lastWeek: Date[] = [];

while (true) {
	console.log(c.blue("Row:"), c.yellow(row));

	if (row > worksheet.rowCount) break;

	// check if all the rows are dates
	const isDateRow = [3, 4, 5, 6, 7].every((col) => {
		const val = isDate(worksheet.getCell(row, col).value);
		return val;
	});

	if (isDateRow) {
		lastWeek = [];

		console.log(c.green("Date row found"));

		lastWeek = [3, 4, 5, 6, 7].map((col) => {
			const date = getDateFromValue(worksheet.getCell(row, col).value);

			if (!date) throw new Error("Date not found");

			return date;
		});

		console.log(c.green("Last week:"), c.yellow(lastWeek));
	}

	// check if it's a timeslot
	const isTimeSlotRow = worksheet
		.getCell(row, 2)
		.value?.toString()
		.match(config.timeSlotRegex);

	if (isTimeSlotRow) {
		console.log(c.green("Time slot row found"), c.yellow(isTimeSlotRow));

		if (
			!isTimeSlotRow[1] ||
			!isTimeSlotRow[2] ||
			!isTimeSlotRow[3] ||
			!isTimeSlotRow[6]
		)
			throw new Error("Time slot row not found");

		//  01.00 pm - 02.00 pm,01,00,pm,02,00,pm
		const startTime = getHours(Number(isTimeSlotRow[1]), isTimeSlotRow[3]);
		const endTime = getHours(Number(isTimeSlotRow[4]), isTimeSlotRow[6]);

		console.log(
			c.green("Start time:"),
			c.yellow(startTime),
			c.green("End time:"),
			c.yellow(endTime)
		);

		if (!lastWeek.length) throw new Error("Last week not found");

		const events = [3, 4, 5, 6, 7]
			.map((col, i) => {
				const event = worksheet.getCell(row, col).value;

				if (!event) return;

				return {
					weekIndex: i,
					event: event.toString(),
				};
			})
			.filter((i) => i !== undefined);

		events.forEach((event) => {
			if (!event) return;

			const date = lastWeek[event.weekIndex];

			if (!date) throw new Error("Date not found");

			const startDate = new Date(date);
			startDate.setHours(startTime, 0, 0, 0);
			const endDate = new Date(date);
			endDate.setHours(endTime, 0, 0, 0);

			console.log(
				c.green("Start date:"),
				c.yellow(startDate),
				c.green("End date:"),
				c.yellow(endDate)
			);

			weeks.push({
				start: startDate,
				end: endDate,
				name: event.event,
			});
		});

		console.log(c.green("Events:"), c.yellow(events));
	}

	row++;
}

weeks = weeks.sort((a, b) => a.start.getTime() - b.start.getTime());

type Week = (typeof weeks)[number];

function mergeConsecutiveWeeks(weeks: Week[]): Week[] {
	if (weeks.length <= 1) return weeks;

	const result: Week[] = [];
	let current = weeks[0]!;

	for (let i = 1; i < weeks.length; i++) {
		const next = weeks[i]!;

		// Same date AND consecutive times AND same name
		if (
			current.end.toDateString() === next.start.toDateString() &&
			current.end.getTime() === next.start.getTime() &&
			current.name === next.name
		) {
			current = { ...current, end: next.end };
		} else {
			result.push(current);
			current = next;
		}
	}

	result.push(current);
	return result;
}

// merge events that are consecutive
weeks = mergeConsecutiveWeeks(weeks);

const calendar = ical({ name: MODULE_NAME, timezone: "Asia/Colombo" });

weeks.forEach((week) => {
	calendar.createEvent({
		start: week.start,
		end: week.end,
		summary: week.name,
	});
});

// writeFileSync(config.name + ".ics", calendar.toString());

const octokit = new Octokit({
	auth: process.env.TOKEN,
});

await octokit.request("PATCH /gists/{gist_id}", {
	gist_id: config.id,
	files: {
		[config.name + ".ics"]: {
			content: calendar.toString(),
		},
	},
	headers: {
		"X-GitHub-Api-Version": "2022-11-28",
	},
});
