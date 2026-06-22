import axios from "axios";

interface TokenCache {
  token: string;
  expiresAt: number; // epoch ms
}

let cache: TokenCache | null = null;

const TENANT_ID = process.env.D365_TENANT_ID!;
const CLIENT_ID = process.env.D365_CLIENT_ID!;
const CLIENT_SECRET = process.env.D365_CLIENT_SECRET!;
const RESOURCE_URL = process.env.D365_RESOURCE_URL!; // e.g. https://yourorg.crm.dynamics.com

function assertEnv() {
  const missing = ["D365_TENANT_ID", "D365_CLIENT_ID", "D365_CLIENT_SECRET", "D365_RESOURCE_URL"].filter(
    (k) => !process.env[k]
  );
  if (missing.length) {
    throw new Error(`Missing required environment variables: ${missing.join(", ")}. Check your .env file.`);
  }
}

async function getAccessToken(): Promise<string> {
  assertEnv();

  if (cache && cache.expiresAt > Date.now() + 30_000) {
    return cache.token;
  }

  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    grant_type: "client_credentials",
    scope: `${RESOURCE_URL}/.default`,
  });

  const resp = await axios.post(url, body, {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  const { access_token, expires_in } = resp.data;
  cache = {
    token: access_token,
    expiresAt: Date.now() + expires_in * 1000,
  };
  return access_token;
}

export async function dataverseRequest(
  method: "POST" | "PATCH" | "GET",
  path: string,
  data?: unknown
) {
  const token = await getAccessToken();
  const url = `${RESOURCE_URL}/api/data/v9.2/${path}`;

  try {
    const resp = await axios({
      method,
      url,
      data,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Prefer: "return=representation",
      },
    });
    return resp.data;
  } catch (err: any) {
    const detail = err?.response?.data?.error?.message || err.message;
    throw new Error(`Dataverse request failed (${method} ${path}): ${detail}`);
  }
}
