const PLACEHOLDER_PATTERN = /(?:\[[A-Z0-9_]+\]|\<[A-Z0-9_]+\>)/g;

export type DeanonymiseResult = {
  text: string;
  restoredCount: number;
  unresolvedCount: number;
  unusedCount: number;
};

export function deanonymiseText(
  mapping: Record<string, string>,
  placeholderText: string
): DeanonymiseResult {
  let restoredCount = 0;
  let unresolvedCount = 0;

  const text = placeholderText.replace(PLACEHOLDER_PATTERN, (placeholder) => {
    const value = resolvePlaceholderValue(mapping, placeholder);
    if (!value) {
      unresolvedCount += 1;
      return placeholder;
    }

    restoredCount += 1;
    return value;
  });

  const unusedCount = Object.keys(mapping).filter(
    (placeholder) => !placeholderText.includes(placeholder)
  ).length;

  return {
    text,
    restoredCount,
    unresolvedCount,
    unusedCount
  };
}

function resolvePlaceholderValue(mapping: Record<string, string>, placeholder: string): string | null {
  const direct = mapping[placeholder];
  if (direct) {
    return direct;
  }

  const match = placeholder.match(/^[<\[]([A-Z_]+)_(PREFERRED_NAME|FIRST_NAME|SURNAME)_(\d{3})[>\]]$/);
  if (!match) {
    return null;
  }

  const [, category, part, index] = match;
  const baseValue = mapping[`[${category}_${index}]`] ?? mapping[`<${category}_${index}>`];
  if (!baseValue) {
    return null;
  }

  const { firstName, surname } = splitPersonName(baseValue);
  if (part === "SURNAME") {
    return surname;
  }

  const preferredName =
    mapping[`[${category}_PREFERRED_NAME_${index}]`] ??
    mapping[`<${category}_PREFERRED_NAME_${index}>`];
  return preferredName ?? firstName;
}

function splitPersonName(value: string): { firstName: string | null; surname: string | null } {
  const cleaned = value
    .trim()
    .replace(/\b(?:Mr|Mrs|Ms|Miss|Dr)\.?\s+/gi, "")
    .replace(/\s+/g, " ");
  const parts = cleaned.split(" ").filter(Boolean);
  if (parts.length === 0) {
    return { firstName: null, surname: null };
  }
  if (parts.length === 1) {
    return { firstName: parts[0], surname: null };
  }
  return { firstName: parts[0], surname: parts[parts.length - 1] };
}
