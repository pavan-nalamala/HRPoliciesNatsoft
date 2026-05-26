import type { SPFI } from '@pnp/sp';

export type WorkflowStatus = 'Pending' | 'Completed';

export async function getLatestItemByCanId(
  sp: SPFI,
  listTitle: string,
  canId: string | number
): Promise<Record<string, unknown> | null> {
  const items = await sp.web.lists
    .getByTitle(listTitle)
    .items.filter(`can_id eq '${canId}'`)
    .top(1)
    .orderBy('Id', false)();

  return (items[0] as Record<string, unknown>) ?? null;
}

/** Remove null/undefined so SharePoint is not sent invalid date placeholders. */
export function stripNullishFields(
  payload: Record<string, unknown>
): Record<string, unknown> {
  const result: Record<string, unknown> = {};

  for (const key of Object.keys(payload)) {
    const value = payload[key];
    if (value !== null && value !== undefined && value !== '') {
      result[key] = value;
    }
  }

  return result;
}

/** Create or update the latest list item for a candidate (matched by can_id). */
export async function upsertByCanId(
  sp: SPFI,
  listTitle: string,
  canId: string | number,
  payload: Record<string, unknown>
): Promise<number> {
  const existing = await getLatestItemByCanId(sp, listTitle, canId);
  const body = stripNullishFields({ ...payload, can_id: String(canId) });

  if (existing?.Id) {
    await sp.web.lists
      .getByTitle(listTitle)
      .items.getById(Number(existing.Id))
      .update(body);
    return Number(existing.Id);
  }

  const created = await sp.web.lists.getByTitle(listTitle).items.add(body);
  return created.data?.Id ?? (created as { Id?: number }).Id ?? 0;
}

/**
 * Replace a list item attachment (delete same file name first).
 * SharePoint returns 400 if you add a file that already exists on the item.
 */
export async function replaceListItemAttachment(
  sp: SPFI,
  listTitle: string,
  itemId: number,
  fileName: string,
  file: File | string | null | undefined
): Promise<void> {
  if (!(file instanceof File) || file.size <= 0) {
    return;
  }

  const item = sp.web.lists.getByTitle(listTitle).items.getById(itemId);
  const attachments = await item.attachmentFiles();

  if (attachments.some((a) => a.FileName === fileName)) {
    await item.attachmentFiles.getByName(fileName).delete();
  }

  await item.attachmentFiles.add(fileName, file);
}

/** Load child list rows linked to a parent item (tries common ParentID column names). */
export async function getListItemsByParentId(
  sp: SPFI,
  listTitle: string,
  parentId: number
): Promise<Record<string, unknown>[]> {
  const filters = [
    `ParentID eq ${parentId}`,
    `ParentID eq '${parentId}'`,
    `ParentId eq ${parentId}`,
    `Parent_x0020_ID eq ${parentId}`
  ];

  for (const filter of filters) {
    try {
      const items = await sp.web.lists
        .getByTitle(listTitle)
        .items.filter(filter)();

      if (items.length > 0) {
        return items as Record<string, unknown>[];
      }
    } catch {
      // Try next filter shape.
    }
  }

  return [];
}

export function toAttachmentPreviewUrl(serverRelativeUrl?: string): string {
  if (!serverRelativeUrl) {
    return '';
  }
  if (serverRelativeUrl.startsWith('http')) {
    return serverRelativeUrl;
  }
  return `${window.location.origin}${serverRelativeUrl}`;
}
