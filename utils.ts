import type { CellSharedFormulaValue, CellValue } from "exceljs";
import type {
	Cell,
	CellErrorValue,
	CellRichTextValue,
	CellHyperlinkValue,
	CellFormulaValue,
} from "exceljs";
import { config, LK_TZ } from "./config";
import { fromZonedTime } from "date-fns-tz";

export function getDateFromValue(val: CellValue): Date | undefined {
	if (val instanceof Date) return val;
	if (typeof val === "object" && val !== null) {
		if ((val as CellSharedFormulaValue)["result"] instanceof Date) {
			return (val as CellSharedFormulaValue)["result"] as Date;
		}
	}
	return undefined;
}

export function isDate(val: CellValue): val is Date {
	return getDateFromValue(val) !== undefined;
}

export function normalizeDate(date: Date) {
	return;
}

export function normalizeName(name: string) {
	const modules = Object.keys(config.modules);

	modules.forEach((i) => {
		name = name.replace(i, config.modules[i as keyof typeof config.modules]);
	});

	return name.replace(/\n/g, " ").split(" ").filter(Boolean).join(" ");
}

export function toLkDate(date: Date): Date {
	const year = date.getUTCFullYear();
	const month = date.getUTCMonth();
	const day = date.getUTCDate();

	const wallClock = new Date(year, month, day, 0, 0, 0, 0);

	return fromZonedTime(wallClock, LK_TZ);
}

export function setLkHour(date: Date, hour: number) {
	const year = date.getUTCFullYear();
	const month = date.getUTCMonth();
	const day = date.getUTCDate();

	const wallClock = new Date(year, month, day, hour, 0, 0, 0);

	return fromZonedTime(wallClock, LK_TZ);
}

export function getRawValue(
	cell: Cell,
): string | number | boolean | Date | null {
	return unwrap(cell.value);
}

function unwrap(value: unknown): string | number | boolean | Date | null {
	if (value === null || value === undefined) return null;

	if (
		typeof value === "number" ||
		typeof value === "boolean" ||
		typeof value === "string" ||
		value instanceof Date
	) {
		return value;
	}

	if (isFormula(value)) {
		return unwrap(value.result);
	}

	return String(value);
}

function isFormula(value: unknown): value is CellFormulaValue {
	return typeof value === "object" && value !== null && "formula" in value;
}
