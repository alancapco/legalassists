/** Parse Capco Wayne field assignments from assistant replies. */

export type AssignableField = {
  key: string;
  label: string;
  placeholder: string;
  sample_context?: string | null;
};

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function cleanValue(raw: string) {
  return raw
    .replace(/^["'`]+|["'`]+$/g, "")
    .replace(/\s*\*\*\s*$/g, "")
    .replace(/\s+$/g, "")
    .trim();
}

function isPlaceholderEcho(val: string, field: AssignableField): boolean {
  const v = val.trim();
  if (!v) return true;
  if (/^(\(empty\)|\(not provided\)|\(unknown\)|tbd|to be determined|null|undefined|<empty>)$/i.test(v)) {
    return true;
  }
  if (
    v.toLowerCase() === (field.placeholder || "").trim().toLowerCase() ||
    v.toLowerCase() === `[${(field.label || "").trim().toLowerCase()}]`
  ) {
    return true;
  }
  return false;
}

/** Prefer fenced ```fields blocks; supports single-line and multi-line bullet values. */
export function parseAssignedValues(
  reply: string,
  fields: AssignableField[],
): Record<string, string> {
  const next: Record<string, string> = {};
  if (!reply.trim() || !fields.length) return next;

  const fencedMatches = Array.from(reply.matchAll(/```fields\s*([\s\S]*?)```/gi));
  const hasFenced = fencedMatches.length > 0;
  const body = hasFenced ? fencedMatches.map((m) => m[1]).join("\n") : reply;

  // Map normalized alias -> field
  const aliasEntries: Array<{ regex: RegExp; field: AssignableField }> = [];
  for (const field of fields) {
    const aliases = Array.from(
      new Set([field.key, field.label, field.placeholder].filter(Boolean)),
    );
    for (const alias of aliases) {
      aliasEntries.push({
        regex: new RegExp(
          `^(?:[-*•]\\s*)?(?:\\*\\*|\`)?${escapeRegExp(alias)}(?:\\*\\*|\`)?\\s*[:=]\\s*(.*)$`,
          "i",
        ),
        field,
      });
    }
  }

  const lines = body.split(/\r?\n/);
  let currentField: AssignableField | null = null;
  let currentParts: string[] = [];

  const flush = () => {
    if (!currentField || !currentParts.length) {
      currentField = null;
      currentParts = [];
      return;
    }
    const joined = cleanValue(currentParts.join(" "));
    if (joined && !isPlaceholderEcho(joined, currentField) && joined !== "NA") {
      next[currentField.key] = joined;
    }
    currentField = null;
    currentParts = [];
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      flush();
      continue;
    }

    let matchedField: AssignableField | null = null;
    let remainder = "";
    for (const entry of aliasEntries) {
      const m = line.match(entry.regex);
      if (m) {
        matchedField = entry.field;
        remainder = m[1] || "";
        break;
      }
    }

    if (matchedField) {
      flush();
      currentField = matchedField;
      const cleaned = cleanValue(remainder);
      if (cleaned) currentParts.push(cleaned);
    } else if (hasFenced && currentField) {
      // Multi-line continuation inside ```fields block (e.g. bullet items)
      const cont = cleanValue(line);
      if (cont) currentParts.push(cont);
    }
  }
  flush();

  return next;
}

export function followUpPrompt(
  fields: AssignableField[],
  values: Record<string, string>,
  appliedCount: number,
  options?: { assignedCount?: number },
): string {
  const empty = fields.filter((f) => !(values[f.key] || "").trim());
  const assignedCount = options?.assignedCount ?? 0;
  const syncFailed = assignedCount > 0 && appliedCount === 0;

  if (syncFailed) {
    const next = empty[0] || fields[0];
    return (
      `I captured ${assignedCount} value(s) in Variables, but they did **not** appear in the Word canvas yet ` +
      `(placeholder text was not found for replace). Try the Variables panel → edit a field and blur, or restate the value.\n\n` +
      (next
        ? `**Next:** Confirm the value for **${next.label || next.placeholder || next.key}**.`
        : "")
    );
  }

  if (!empty.length) {
    return appliedCount
      ? `I've written ${appliedCount} value(s) into the Word document. All yellow fields are filled — say if you want a playbook review next.`
      : "All yellow fields are filled. Want a Capco playbook review?";
  }
  const next = empty[0];
  const remaining = empty
    .slice(0, 6)
    .map((f) => f.label || f.placeholder || f.key)
    .join(", ");
  const wrote = appliedCount
    ? `I've updated the document with ${appliedCount} value(s). `
    : "";
  return (
    `${wrote}` +
    `Still need ${empty.length} field(s): ${remaining}.\n\n` +
    `**Next:** What should I use for **${next.label || next.placeholder || next.key}**?`
  );
}
