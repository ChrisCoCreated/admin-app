import type { PseudonymisationResponse, SafetyCheckResponse } from "./pseudonymiserTypes";

export async function pseudonymiseNote(
  apiBaseUrl: string,
  text: string
): Promise<PseudonymisationResponse> {
  const response = await fetch(`${trimTrailingSlash(apiBaseUrl)}/pseudonymise`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ text })
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(detail || `Pseudonymiser returned HTTP ${response.status}`);
  }

  return normalizePseudonymisationResponse(await response.json());
}

export async function safetyCheckText(
  apiBaseUrl: string,
  text: string
): Promise<SafetyCheckResponse> {
  const response = await fetch(`${trimTrailingSlash(apiBaseUrl)}/safety-check`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ text })
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(detail || `Pseudonymiser returned HTTP ${response.status}`);
  }

  return response.json();
}

function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, "");
}

function normalizePseudonymisationResponse(payload: unknown): PseudonymisationResponse {
  const source = isRecord(payload) ? payload : {};
  const findings = isRecord(source.findings) ? source.findings : {};

  return {
    pseudonymised_text: typeof source.pseudonymised_text === "string" ? source.pseudonymised_text : "",
    mapping: isStringRecord(source.mapping) ? source.mapping : {},
    findings: {
      entities_detected: stringArray(findings.entities_detected),
      risks_preserved: stringArray(findings.risks_preserved),
      clinical_details_preserved: stringArray(findings.clinical_details_preserved),
      warnings: stringArray(findings.warnings),
      riskLevel: isRiskLevel(findings.riskLevel) ? findings.riskLevel : "LOW",
      safeToSend: typeof findings.safeToSend === "boolean" ? findings.safeToSend : true,
      directIdentifiers: stringArray(findings.directIdentifiers),
      indirectIdentifiers: stringArray(findings.indirectIdentifiers),
      residualSpans: residualSpanArray(findings.residualSpans),
      reason:
        typeof findings.reason === "string"
          ? findings.reason
          : "No unresolved direct identifiers found."
    },
    entities: findingArray(source.entities),
    counts: countArray(source.counts)
  };
}

function isRiskLevel(value: unknown): value is PseudonymisationResponse["findings"]["riskLevel"] {
  return value === "HIGH" || value === "MEDIUM" || value === "LOW";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isStringRecord(value: unknown): value is Record<string, string> {
  return isRecord(value) && Object.values(value).every((item) => typeof item === "string");
}

function stringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string") : [];
}

function findingArray(value: unknown): PseudonymisationResponse["entities"] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.flatMap((item) => {
    if (!isRecord(item)) {
      return [];
    }

    const entityType = typeof item.entity_type === "string" ? item.entity_type : null;
    const replacement = typeof item.replacement === "string" ? item.replacement : null;
    if (!entityType || !replacement) {
      return [];
    }

    return [
      {
        entity_type: entityType,
        start: typeof item.start === "number" ? item.start : 0,
        end: typeof item.end === "number" ? item.end : 0,
        score: typeof item.score === "number" ? item.score : 0,
        replacement
      }
    ];
  });
}

function residualSpanArray(value: unknown): PseudonymisationResponse["findings"]["residualSpans"] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.flatMap((item) => {
    if (!isRecord(item)) {
      return [];
    }

    const start = typeof item.start === "number" ? item.start : null;
    const end = typeof item.end === "number" ? item.end : null;
    const severity = item.severity === "direct" || item.severity === "indirect" ? item.severity : null;
    const category = typeof item.category === "string" ? item.category : null;
    const reason = typeof item.reason === "string" ? item.reason : null;
    if (start === null || end === null || !severity || !category || !reason || start >= end) {
      return [];
    }

    return [{ start, end, severity, category, reason }];
  });
}

function countArray(value: unknown): PseudonymisationResponse["counts"] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.flatMap((item) => {
    if (!isRecord(item) || typeof item.entity_type !== "string" || typeof item.count !== "number") {
      return [];
    }

    return [{ entity_type: item.entity_type, count: item.count }];
  });
}
