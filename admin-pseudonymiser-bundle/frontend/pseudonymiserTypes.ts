export type EntityCount = {
  entity_type: string;
  count: number;
};

export type Finding = {
  entity_type: string;
  start: number;
  end: number;
  score: number;
  replacement: string;
};

export type ResidualSpan = {
  start: number;
  end: number;
  severity: "direct" | "indirect";
  category: string;
  reason: string;
};

export type PseudonymisationFindings = {
  entities_detected: string[];
  risks_preserved: string[];
  clinical_details_preserved: string[];
  warnings: string[];
  riskLevel: "HIGH" | "MEDIUM" | "LOW";
  safeToSend: boolean;
  directIdentifiers: string[];
  indirectIdentifiers: string[];
  residualSpans: ResidualSpan[];
  reason: string;
};

export type PseudonymisationResponse = {
  pseudonymised_text: string;
  mapping: Record<string, string>;
  findings: PseudonymisationFindings;
  entities: Finding[];
  counts: EntityCount[];
};

export type SafetyCheckResponse = {
  warnings: string[];
  riskLevel: "HIGH" | "MEDIUM" | "LOW";
  safeToSend: boolean;
  directIdentifiers: string[];
  indirectIdentifiers: string[];
  reason: string;
};
