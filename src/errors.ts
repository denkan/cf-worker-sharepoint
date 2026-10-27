import type { ContentfulStatusCode } from "hono/utils/http-status";

export type ErrorCode =
  | "BAD_REQUEST"
  | "NOT_FOUND"
  | "NOT_A_FOLDER"
  | "NOT_A_FILE"
  | "LIMIT_EXCEEDED"
  | "INDEX_NOT_READY"
  | "UPSTREAM_ERROR"
  | "UPSTREAM_AUTH_ERROR"
  | "INTERNAL_ERROR";

export class ApiError extends Error {
  constructor(
    readonly status: ContentfulStatusCode,
    readonly code: ErrorCode,
    message: string,
    readonly details?: Record<string, unknown>,
  ) {
    super(message);
  }
}

export const notFound = (message = "Not found within the configured root folder") =>
  new ApiError(404, "NOT_FOUND", message);

export const badRequest = (message: string, details?: Record<string, unknown>) =>
  new ApiError(400, "BAD_REQUEST", message, details);

export function errorBody(code: ErrorCode, message: string, details?: Record<string, unknown>) {
  return { error: { code, message, ...(details ? { details } : {}) } };
}
