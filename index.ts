import ExcelJS, { type CellSharedFormulaValue, type CellValue } from "exceljs";
import c from "chalk";
import ical from "ical-generator";
import { Octokit } from "@octokit/core";

const CONFIGS = [
	{
		name: "24.3-fdn",
		worksheet: "24.3 & 24.4 FDN",
		timeTableStart: 7,
		dataIndex: 2,
		summaryCell: "B4",
		timeSlotRegex: /(\d+)\.(\d+) ([a-zA-Z]+)\s?-\s?(\d+)\.(\d+) ([a-zA-Z]+)/,
		id: process.env.GIST_ID_24_3_FDN,
	},
];

console.log(c.bgBlue.black("NSBM Timetable Converter"));
console.log(
	c.green("Found"),
	c.yellow(CONFIGS.length),
	c.green("configurations")
);

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

for (const config of CONFIGS) {
	console.log(c.blue("File:"), c.yellow(config.name));

	if (!config.id) {
		console.log(c.red("Gist ID not found"));
		continue;
	}

	const timetable = await Bun.file(
		"./downloaded/" + config.name + ".xlsx"
	).arrayBuffer();

	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(timetable);

	const worksheet = workbook.getWorksheet(config.worksheet);
	let weeks: {
		start: Date;
		end: Date;
		name: string;
	}[][] = [];
	let row = config.timeTableStart;
	let lastWeek: Date[] = [];

	if (!worksheet) throw new Error("Worksheet not found");

	const MODULE_NAME = worksheet
		.getCell(config.summaryCell)
		.text.replace("Time Table - ", "")
		.trim();

	console.log(c.blue("Module:"), c.yellow(MODULE_NAME));

	while (true) {
		const currentRow = worksheet.getRow(row);

		if (!currentRow.getCell(config.dataIndex).text.match(/\d|\s/g)) break;

		/* date check */
		if (isDate(currentRow.getCell("C").value)) {
			lastWeek = [];
			weeks.push([]);

			currentRow.eachCell((cell, i) => {
				if (isDate(cell.value)) lastWeek.push(getDateFromValue(cell.value)!);
			});
		}

		const match = currentRow.getCell("B").text.match(config.timeSlotRegex) as
			| [string, number, number, string, number, number, string]
			| null;

		if (match) {
			const week = weeks[weeks.length - 1];

			if (!week) throw new Error("Week not found");

			/**
			 * Group - 09.00 am - 10.00 am
			 *         ^  ^  ^     ^ ^  ^
			 *         1  2  3     4 5  6
			 */

			for (let day = 0; day < 5; day++) {
				const cell = currentRow.getCell(config.dataIndex + 1 + day).text.trim();
				const dateObj = lastWeek[day];

				if (!dateObj) throw new Error("No date found");

				// no subject
				if (!cell) continue;

				const start = new Date();
				const end = new Date();

				start.setTime(dateObj.getTime());
				end.setTime(dateObj.getTime());

				start.setHours(getHours(+match[1], match[3]), 0, 0);
				end.setHours(getHours(+match[4], match[6]), 0, 0);

				week.push({
					name: cell,
					start,
					end,
				});
			}
		}

		row++;
	}

	weeks = weeks.map((weekI) => {
		const merge = (week: typeof weekI) => {
			const sorted = week.sort((a, b) => a.start.getTime() - b.start.getTime());
			const merged: typeof week = [];

			for (let i = 0; i < sorted.length; i++) {
				const curr = sorted[i]!;
				const next = sorted[i + 1];

				/**
				 * if all the events in the same day have the same title, merge em
				 */
				const sameday = week.filter(
					(e) => e.start.getDay() === curr.start.getDay()
				);

				if (sameday.length > 2 && sameday.every((e) => e.name === curr.name)) {
					i += sameday.length - 1;

					const end = new Date(curr.end.getTime());

					end.setHours(17, 0, 0);

					merged.push({
						name: curr.name,
						start: curr.start,
						end,
					});

					continue;
				}

				if (
					next &&
					curr.name === next.name &&
					curr.end.getTime() === next.start.getTime()
				) {
					curr.end = next.end;
				}

				merged.push(curr);

				i++;
			}

			return merged;
		};

		return merge(weekI);
	});

	console.log(c.blue("Weeks:"), c.yellow(weeks.length));

	const calendar = ical({ name: "NSBM Timetable", timezone: "Asia/Colombo" });

	for (const week of weeks) {
		for (const subject of week) {
			calendar.createEvent({
				start: subject.start,
				end: subject.end,
				summary: subject.name,
			});
		}
	}

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
}
