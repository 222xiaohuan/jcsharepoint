const controlUrl = process.env.SYNC_CONTROL_URL || "http://127.0.0.1:8765/api/schedule-due";

function pad(value) {
  return String(value).padStart(2, "0");
}

export function currentScheduleSlot(now = new Date()) {
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`;
}

export async function notifyScheduleDue(now = new Date()) {
  const slotId = currentScheduleSlot(now);
  const response = await fetch(controlUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ slotId, scheduledAt: now.toISOString() }),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`控制服务未接受定时触发: HTTP ${response.status} ${body}`);
  return JSON.parse(body);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  try {
    const result = await notifyScheduleDue();
    console.log(JSON.stringify(result));
  } catch (error) {
    console.error(error.stack || error);
    process.exitCode = 1;
  }
}
