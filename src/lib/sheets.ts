import { google } from "googleapis";

export type SessionData = {
  name: string;
  speaker: string;
  week: number;
  day: string;
  status: "P" | "A" | "L" | "C" | "R" | "";
  isFireside?: boolean;
};

export type StudentData = {
  name: string;
  email: string;
  school: string;
  attendance: number;
  totalEvents: number;
  stats: {
    present: number;
    reflection: number;
    late: number;
    conflict: number;
    absent: number;
    presentPercentage: number;
  };
  sessions: SessionData[];
};

const getSheets = () => {
  return google.sheets({
    version: "v4",
    auth: process.env.GOOGLE_API_KEY,
  });
};

async function getFiresideData(name: string) {
  try {
    const sheets = getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEETS_SHEET_ID,
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

export async function getStudentData(
  email: string,
): Promise<StudentData | null> {
  try {
    const sheets = getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEETS_SHEET_ID,
      range: "A:ZZ",
    });

    const rows = response.data.values;
    if (!rows) return null;

    const studentRow = rows.find((row) => row[1] === email);
    if (!studentRow) return null;

    const weekDays = rows[0].slice(11);
    const sessionNames = rows[1].slice(11);
    const speakers = rows[2].slice(11);
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
        status: studentRow[index + 11] || "",
        isFireside: false,
      };
    });

    const firesideData = await getFiresideData(studentRow[0]);

    const allStatuses = [
      ...studentRow.slice(11),
      ...firesideData.map((session) => session!.status),
    ];

    const stats = {
      present: allStatuses.filter((val) => val === "P").length,
      reflection: allStatuses.filter((val) => val === "R").length,
      late: allStatuses.filter((val) => val === "L").length,
      conflict: allStatuses.filter((val) => val === "C").length,
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

    return {
      name: studentRow[0],
      email: studentRow[1],
      school: studentRow[2],
      attendance: parseFloat(studentRow[3]),
      totalEvents: parseInt(studentRow[4]),
      stats,
      sessions: allSessions as SessionData[],
    };
  } catch (error) {
    console.error("Error fetching student data:", error);
    return null;
  }
}
