import { google } from "googleapis";

export type Status = "P" | "A" | "L" | "R" | "";

export type SessionData = {
  name: string;
  speaker: string;
  week: number;
  day: string;
  status: Status;
  isFireside?: boolean;
};

export type StudentData = {
  name: string;
  email: string;
  school: string;
  attendance: number;
  stats: {
    present: number;
    reflection: number;
    late: number;
    absent: number;
    presentPercentage: number;
  };
  sessions: SessionData[];
  project: {
    status: Status;
  };
  recordingUrl: string;
};

const getSheets = () => {
  return google.sheets({
    version: "v4",
    auth: process.env.GOOGLE_API_KEY,
  });
};

/**
 * Collect all configured sheet IDs from env vars.
 * Supports GOOGLE_SHEETS_SHEET_ID (legacy/default) and
 * GOOGLE_SHEETS_SHEET_ID_19, GOOGLE_SHEETS_SHEET_ID_20, etc.
 * Returns newest cohort first so recent students are found faster.
 */
function getAllSheetIds(): string[] {
  const ids: { num: number; id: string }[] = [];

  // Numbered per-cohort sheets (GOOGLE_SHEETS_SHEET_ID_19, etc.)
  for (const [key, value] of Object.entries(process.env)) {
    const match = key.match(/^GOOGLE_SHEETS_SHEET_ID_(\d+)$/);
    if (match && value) ids.push({ num: parseInt(match[1]), id: value });
  }

  // Sort newest cohort first
  ids.sort((a, b) => b.num - a.num);

  // Default/legacy sheet last (fallback)
  const defaultId = process.env.GOOGLE_SHEETS_SHEET_ID;
  if (defaultId) {
    // Avoid duplicates if default is also listed as a numbered var
    if (!ids.some((e) => e.id === defaultId)) {
      ids.push({ num: 0, id: defaultId });
    }
  }

  return ids.map((e) => e.id);
}

async function getFiresideData(name: string, sheetId: string) {
  try {
    const sheets = getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "'Fireside Chats'!A:ZZ",
    });

    const rows = response.data.values;
    if (!rows) return [];

    const daysRow = rows[0]; // Row with Tuesday/Thursday
    const weekInfoRow = rows[1]; // Row with W1,F1 etc
    const studentRow = rows.find(
      (row) => row[0]?.trim().toLowerCase() === name?.trim().toLowerCase(),
    );

    if (!studentRow || !daysRow || !weekInfoRow) return [];

    const sessions = daysRow
      .slice(6)
      .map((day, index) => {
        if (!day) return null; // Skip empty columns

        const status = studentRow[index + 6];
        if (!status || status === "") return null;

        const weekInfo = weekInfoRow[index + 6]; // e.g., "W1, F1"
        if (!weekInfo) return null;

        const weekMatch = weekInfo.match(/W(\d+)/);
        const week = weekMatch ? parseInt(weekMatch[1]) : 0;

        return {
          name: "Fireside Chat",
          speaker: weekInfo,
          week,
          day,
          status,
          isFireside: true,
        };
      })
      .filter(Boolean);

    return sessions;
  } catch (error) {
    console.error("Error fetching fireside data:", error);
    return [];
  }
}

async function getProjectData(name: string, sheetId: string) {
  try {
    const sheets = getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "'Project'!A:ZZ",
    });

    const rows = response.data.values;
    if (!rows) return { status: "" };

    const studentRow = rows.find(
      (row) => row[0]?.trim().toLowerCase() === name?.trim().toLowerCase(),
    );

    if (!studentRow) return { status: "" };

    return { status: studentRow[2] || "" };
  } catch (error) {
    console.error("Error fetching project data:", error);
    return { status: "" };
  }
}

export async function getStudentData(
  email: string,
): Promise<StudentData | null> {
  const sheetIds = getAllSheetIds();
  if (sheetIds.length === 0) {
    console.error("No sheet IDs configured. Set GOOGLE_SHEETS_SHEET_ID or GOOGLE_SHEETS_SHEET_ID_<N>.");
    return null;
  }

  // Search each configured sheet for the student's email
  for (const sheetId of sheetIds) {
    try {
      const result = await getStudentDataFromSheet(email, sheetId);
      if (result) return result;
    } catch (error) {
      console.warn(`[sheets] Failed to read sheet ${sheetId}:`, error);
      // Continue to next sheet
    }
  }

  return null;
}

async function getStudentDataFromSheet(
  email: string,
  sheetId: string,
): Promise<StudentData | null> {
  const sheets = getSheets();
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'Events'!A:ZZ",
  });

  const rows = response.data.values;
  if (!rows) return null;

  const studentRow = rows.find((row) => row[1] === email);
  if (!studentRow) return null;

  const weekDays = rows[0].slice(10);
  const sessionNames = rows[1].slice(10);
  const speakers = rows[2].slice(10);
  let weekNumber = 0;

  const sessions = sessionNames.map((name, index) => {
    const weekDay = weekDays[index] || "";
    const weekMatch = weekDay.match(/Wk (\d+)/);
    const week = weekMatch ? parseInt(weekMatch[1]) : 0;
    weekNumber = week || weekNumber;
    const dayMatch = weekDay.match(
      /(Monday|Tuesday|Wednesday|Thursday|Friday)/,
    );
    const day = dayMatch ? dayMatch[1] : "";

    return {
      name: name,
      speaker: speakers[index],
      week: weekNumber,
      day: day,
      status: studentRow[index + 10] || "",
      isFireside: false,
    };
  });

  const firesideData = await getFiresideData(studentRow[0], sheetId);

  const allStatuses = [
    ...studentRow.slice(10),
    ...firesideData.map((session) => session!.status),
  ];

  const stats = {
    present: allStatuses.filter((val) => val === "P").length,
    reflection: allStatuses.filter((val) => val === "R").length,
    late: allStatuses.filter((val) => val === "L").length,
    absent: allStatuses.filter((val) => val === "A").length,
    presentPercentage: parseFloat(studentRow[3]),
  };

  const allSessions = [...sessions, ...firesideData].sort((a, b) => {
    if (!a || !b) return 0;
    if (a.week !== b.week) return a.week - b.week;

    const getDayOrder = (day: string) => {
      switch (day) {
        case "Monday":
          return 1;
        case "Tuesday":
          return 2;
        case "Wednesday":
          return 3;
        case "Thursday":
          return 4;
        case "Friday":
          return 5;
        default:
          return 0;
      }
    };

    return getDayOrder(a.day) - getDayOrder(b.day);
  });

  const project = await getProjectData(studentRow[0], sheetId);

  return {
    name: studentRow[0],
    email: studentRow[1],
    school: studentRow[2],
    attendance: parseFloat(studentRow[3]),
    stats,
    sessions: allSessions as SessionData[],
    project,
    recordingUrl: rows[2]?.[0] || "#",
  };
}
