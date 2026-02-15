import { handleDashboardApiRequest } from "../backend/dashboard-api.mjs";

export default async function handler(req, res) {
  return handleDashboardApiRequest(req, res);
}
