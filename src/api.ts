// SPDX-FileCopyrightText: 2026 Cisco Systems, Inc.
// SPDX-License-Identifier: Apache-2.0

import type { AppSnapshot, RecordKind, RosterLabel } from './types';

async function request<T>(input: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers ?? {});
  if (!(init?.body instanceof FormData) && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  let response: Response;
  try {
    response = await fetch(input, {
      headers,
      ...init,
    });
  } catch (error) {
    if (error instanceof TypeError) {
      throw new Error('Cannot reach the SOC Cockpit API. Start the API server with npm.cmd run dev:api, or run npm.cmd run dev for the full stack.');
    }
    throw error;
  }

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error ?? `Request failed with status ${response.status}`);
  }

  return response.json() as Promise<T>;
}

export const api = {
  getSnapshot: () => request<AppSnapshot>('/api/snapshot'),
  createRecord: (kind: RecordKind, payload: Record<string, unknown>) => request<AppSnapshot>(`/api/${kind}`, { method: 'POST', body: JSON.stringify(payload) }),
  deleteRecord: (kind: RecordKind, id: number) => request<AppSnapshot>(`/api/${kind}/${id}`, { method: 'DELETE' }),
  updateRecord: (kind: RecordKind, id: number, payload: Record<string, unknown>) => request<AppSnapshot>(`/api/${kind}/${id}`, { method: 'PATCH', body: JSON.stringify(payload) }),
  updateClockZones: (zones: string[]) => request<AppSnapshot>('/api/settings/clocks', { method: 'POST', body: JSON.stringify({ zones }) }),
  updateRosterPlannerSettings: (payload: { staff: string[]; labels: RosterLabel[]; channelOrder?: string[] }) => request<AppSnapshot>('/api/settings/roster-planner', { method: 'POST', body: JSON.stringify(payload) }),
  renameRosterStaff: (name: string, nextName: string) => request<AppSnapshot>(`/api/roster/staff/${encodeURIComponent(name)}`, { method: 'PATCH', body: JSON.stringify({ nextName }) }),
  updateRosterLabelDefinition: (code: string, payload: RosterLabel) => request<AppSnapshot>(`/api/roster/label/${encodeURIComponent(code)}`, { method: 'PATCH', body: JSON.stringify(payload) }),
  upsertRosterPlan: (changes: Array<Record<string, unknown>>) => request<AppSnapshot>('/api/roster/planner', { method: 'POST', body: JSON.stringify({ changes }) }),
  updateMatrixTask: (id: number, payload: Record<string, unknown>) => request<AppSnapshot>(`/api/matrixTask/${id}`, { method: 'PATCH', body: JSON.stringify(payload) }),
  carryOverMatrixTasks: (toDate: string, fromDate?: string) => request<AppSnapshot>('/api/matrix/carry-over', { method: 'POST', body: JSON.stringify({ toDate, fromDate }) }),
  downloadStandupTemplate: async () => {
    const response = await fetch('/api/standup/template');
    if (!response.ok) {
      throw new Error(`Could not download the template (status ${response.status}).`);
    }
    return response.blob();
  },
  importStandupWorkbook: (file: File) => {
    const formData = new FormData();
    formData.append('workbook', file);
    return request<{ snapshot: AppSnapshot; added: number; updated: number; skipped: number; objectivesCreated: number; tasksLinked: number; createdObjectives: string[]; ambiguousObjectives: string[]; errors: string[] }>('/api/standup/import', { method: 'POST', body: formData });
  },
  uploadObjectiveDocument: (id: number, file: File) => {
    const formData = new FormData();
    formData.append('document', file);
    return request<AppSnapshot>(`/api/objective/${id}/document`, { method: 'POST', body: formData });
  },
  uploadProcessDocument: (id: number, file: File) => {
    const formData = new FormData();
    formData.append('document', file);
    return request<AppSnapshot>(`/api/process/${id}/document`, { method: 'POST', body: formData });
  },
  generateAdvisorBrief: (date?: string) => request<AppSnapshot>('/api/advisor/generate', { method: 'POST', body: JSON.stringify({ date }) }),
  answerQuestion: (messages: Array<{ role: 'user' | 'assistant'; content: string }>) =>
    request<{ answer: string }>('/api/assistant/chat', { method: 'POST', body: JSON.stringify({ messages }) }),
};
