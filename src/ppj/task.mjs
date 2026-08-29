import { randomUUID } from "node:crypto";
import process from "node:process";

import {
  acquireTaskLock,
  openTask,
  recordTaskPpjRevision,
} from "../cli/task-store.mjs";

export async function recordPpjTask({
  taskId,
  cwd = process.cwd(),
  stage,
  workspace,
  receipt,
  candidate,
  review,
}) {
  if (taskId == null) return null;
  const opened = await openTask({ workspaceRoot: cwd, taskId });
  const sessionId = `ppj_${randomUUID().replaceAll("-", "").slice(0, 16)}`;
  const lock = await acquireTaskLock(opened.taskRoot, { sessionId });
  try {
    const current = await openTask({ workspaceRoot: opened.workspaceRoot, taskId });
    return await recordTaskPpjRevision(current, workspace, {
      stage,
      receipt,
      candidate,
      review,
    });
  } finally {
    await lock.release();
  }
}
