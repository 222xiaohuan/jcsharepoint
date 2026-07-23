import test from "node:test";
import assert from "node:assert/strict";
import { markScheduleDue, decideSchedule, setScheduleEnabled } from "../schedule-state.mjs";

test("启用同步服务时，定时触发会创建待确认项", () => {
  const result = markScheduleDue({ enabled: true, pending: [], handled: [] }, "2026-07-23T08:00", "2026-07-23T08:00:00.000Z");

  assert.equal(result.status, "pending");
  assert.deepEqual(result.state.pending, [{
    slotId: "2026-07-23T08:00",
    scheduledAt: "2026-07-23T08:00:00.000Z",
  }]);
});

test("用户确认继续或忽略后，待确认项会被标记为已处理", () => {
  const result = decideSchedule({
    enabled: true,
    pending: [{ slotId: "2026-07-23T08:00", scheduledAt: "2026-07-23T08:00:00.000Z" }],
    handled: [],
  }, "2026-07-23T08:00", "skip");

  assert.equal(result.status, "skipped");
  assert.deepEqual(result.state.pending, []);
  assert.deepEqual(result.state.handled, ["2026-07-23T08:00"]);
});

test("暂停同步服务会停止定时触发并清空待确认项", () => {
  const result = setScheduleEnabled({
    enabled: true,
    pending: [{ slotId: "2026-07-23T08:00", scheduledAt: "2026-07-23T08:00:00.000Z" }],
    handled: [],
  }, false);

  assert.equal(result.enabled, false);
  assert.deepEqual(result.pending, []);
  assert.equal(markScheduleDue(result, "2026-07-23T12:00").status, "paused");
});
