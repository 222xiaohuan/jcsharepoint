const MAX_HANDLED_SLOTS = 32;

export function normalizeScheduleState(value = {}) {
  return {
    enabled: value.enabled !== false,
    pending: Array.isArray(value.pending) ? value.pending.filter((item) => item?.slotId) : [],
    handled: Array.isArray(value.handled) ? value.handled.filter(Boolean).slice(-MAX_HANDLED_SLOTS) : [],
  };
}

export function markScheduleDue(value, slotId, scheduledAt = new Date().toISOString()) {
  const state = normalizeScheduleState(value);
  if (!state.enabled) return { state, status: "paused" };
  if (state.handled.includes(slotId)) return { state, status: "handled" };
  if (state.pending.some((item) => item.slotId === slotId)) return { state, status: "already-pending" };

  state.pending.push({ slotId, scheduledAt });
  return { state, status: "pending" };
}

export function decideSchedule(value, slotId, action) {
  const state = normalizeScheduleState(value);
  if (!['confirm', 'skip'].includes(action)) throw new Error("无效的定时同步操作");
  const index = state.pending.findIndex((item) => item.slotId === slotId);
  if (index === -1) throw new Error("定时同步确认项不存在或已处理");

  state.pending.splice(index, 1);
  state.handled = [...state.handled.filter((id) => id !== slotId), slotId].slice(-MAX_HANDLED_SLOTS);
  return { state, status: action === "confirm" ? "confirmed" : "skipped" };
}

export function setScheduleEnabled(value, enabled) {
  const state = normalizeScheduleState(value);
  state.enabled = Boolean(enabled);
  if (!state.enabled) state.pending = [];
  return state;
}
