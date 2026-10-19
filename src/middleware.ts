import { NextResponse } from "next/server";
import { auth } from "@/lib/auth";

// AUTH_URL sends Google's callback to this host, so a sign-in started on any
// other production host (e.g. the old possible-engagement-rates.vercel.app
// link) loses its PKCE cookie and fails with error=Configuration.
const CANONICAL_HOST = "engagement.heypossible.com";

export const middleware = auth((req) => {
  const host = req.headers.get("host");
  if (
    process.env.VERCEL_ENV === "production" &&
    host &&
    host !== CANONICAL_HOST
  ) {
    const url = new URL(
      req.nextUrl.pathname + req.nextUrl.search,
      `https://${CANONICAL_HOST}`,
    );
    return NextResponse.redirect(url, 308);
  }
});
