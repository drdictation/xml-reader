import { NextResponse } from "next/server";
import type { NextRequest } from "next/server";

export function middleware(request: NextRequest) {
  const password = process.env.APP_PASSWORD;

  // If no password is set, we assume it's in a safe environment (or not yet configured)
  // For production, this should definitely be set.
  if (!password) {
    return NextResponse.next();
  }

  const { pathname } = request.nextUrl;

  // Allow access to login page and public assets
  if (pathname === "/login" || pathname.startsWith("/_next") || pathname.includes(".")) {
    return NextResponse.next();
  }

  const session = request.cookies.get("session")?.value;

  if (session !== password) {
    return NextResponse.redirect(new URL("/login", request.url));
  }

  return NextResponse.next();
}

export const config = {
  matcher: "/((?!api|_next/static|_next/image|favicon.ico).*)",
};
