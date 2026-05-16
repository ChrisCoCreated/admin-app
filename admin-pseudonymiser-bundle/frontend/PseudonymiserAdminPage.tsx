import { AlertTriangle, Clipboard, Eraser, Loader2, RotateCcw, ShieldCheck, Wand2 } from "lucide-react";
import { useMemo, useState } from "react";
import { deanonymiseText } from "./deanonymise";
import { pseudonymiseNote } from "./pseudonymiserApi";
import type { PseudonymisationResponse, ResidualSpan } from "./pseudonymiserTypes";
import "./pseudonymiser-admin.css";

const EXAMPLE_NOTE =
  "Paulette Crawley feels safe at Lindau. Carrie is brilliant. Discussed Martin Tyrell and selling her properties.";

const LLM_COPY_PREFIX = [
  "Instructions for the LLM:",
  "Keep every placeholder tag exactly as written, including the square brackets.",
  "Do not rename, renumber, remove, expand, or merge any text inside brackets such as [CLIENT_001] or [CLIENT_PREFERRED_NAME_001].",
  "In body text, use [CLIENT_PREFERRED_NAME_001] instead of [CLIENT_001] whenever the client is mentioned naturally in a sentence.",
  "Use [CLIENT_001] only for headings, labels, tables, or when the full placeholder is explicitly needed.",
  "If a preferred-name or surname variant is needed, convert [CLIENT_001] to [CLIENT_PREFERRED_NAME_001] or [CLIENT_SURNAME_001] with the same number.",
  "Do not invent new numbering. Keep the numeric suffix aligned with the original placeholder.",
  "You may rewrite the surrounding narrative, but preserve all bracketed tags verbatim.",
  "",
  "Text:"
].join("\n");

const PLACEHOLDER_CATEGORY_OPTIONS = [
  "CLIENT",
  "STAFF",
  "RELATIVE",
  "FRIEND",
  "PROFESSIONAL",
  "CARE_HOME",
  "LOCATION",
  "ORGANISATION",
  "DATE",
  "PHONE",
  "EMAIL",
  "ADDRESS",
  "IDENTIFIER"
] as const;

type ReviewMark =
  | {
      kind: "replaced";
      start: number;
      end: number;
      placeholder: string;
      original: string;
    }
  | {
      kind: "likely" | "possible";
      start: number;
      end: number;
      category: string;
      reason: string;
      value: string;
      replacement?: string;
    };

export type PseudonymiserAdminPageProps = {
  apiBaseUrl?: string;
  initialText?: string;
  title?: string;
};

export function PseudonymiserAdminPage({
  apiBaseUrl = "/api/pseudonymiser",
  initialText = EXAMPLE_NOTE,
  title = "Care note pseudonymiser"
}: PseudonymiserAdminPageProps) {
  const [input, setInput] = useState(initialText);
  const [placeholderText, setPlaceholderText] = useState("");
  const [result, setResult] = useState<PseudonymisationResponse | null>(null);
  const [reviewText, setReviewText] = useState("");
  const [manualMapping, setManualMapping] = useState<Record<string, string>>({});
  const [clientPreferredName, setClientPreferredName] = useState("");
  const [manualIdentifierText, setManualIdentifierText] = useState("");
  const [manualIdentifierCategory, setManualIdentifierCategory] = useState("PROFESSIONAL");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  const effectiveMapping = useMemo(
    () => buildEffectiveMapping(result?.mapping ?? {}, manualMapping, clientPreferredName),
    [clientPreferredName, manualMapping, result]
  );
  const residualSpans = useMemo(
    () => normaliseResidualSpans(result?.findings.residualSpans ?? [], result?.pseudonymised_text.length ?? 0),
    [result]
  );
  const reviewMarks = useMemo(
    () => buildReviewMarks(reviewText, result?.pseudonymised_text ?? "", effectiveMapping, residualSpans),
    [effectiveMapping, residualSpans, result, reviewText]
  );
  const identifierOptions = useMemo(() => buildIdentifierOptions(effectiveMapping), [effectiveMapping]);
  const pendingSuggestionCount = reviewMarks.filter((mark) => mark.kind !== "replaced").length;
  const restorePreview = useMemo(() => {
    if (!placeholderText.trim() || Object.keys(effectiveMapping).length === 0) {
      return { text: "", restoredCount: 0, unresolvedCount: 0 };
    }
    return deanonymiseText(effectiveMapping, placeholderText);
  }, [effectiveMapping, placeholderText]);
  const canPseudonymise = Boolean(input.trim()) && !busy;

  async function onPseudonymise() {
    if (!input.trim()) {
      return;
    }

    setBusy(true);
    setError(null);
    setCopied(false);
    try {
      const response = await pseudonymiseNote(apiBaseUrl, input);
      setResult(response);
      setReviewText(response.pseudonymised_text);
      setManualMapping({});
      setClientPreferredName((current) => current.trim() || defaultPreferredName(response.mapping));
    } catch (err) {
      setError(err instanceof Error ? err.message : "Pseudonymisation failed.");
    } finally {
      setBusy(false);
    }
  }

  async function onCopyForLlm() {
    if (!reviewText) {
      return;
    }
    await navigator.clipboard.writeText(`${LLM_COPY_PREFIX}\n${reviewText}`);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1400);
  }

  async function onCopyRestored() {
    if (!restorePreview.text) {
      return;
    }
    await navigator.clipboard.writeText(restorePreview.text);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1400);
  }

  function onClear() {
    setInput("");
    setPlaceholderText("");
    setResult(null);
    setReviewText("");
    setManualMapping({});
    setClientPreferredName("");
    setManualIdentifierText("");
    setError(null);
  }

  function replaceReviewRange(start: number, end: number, value: string) {
    setReviewText((current) => `${current.slice(0, start)}${value}${current.slice(end)}`);
    setCopied(false);
  }

  function onRevertReplacement(mark: Extract<ReviewMark, { kind: "replaced" }>) {
    replaceReviewRange(mark.start, mark.end, mark.original);
  }

  function onApplySuggestion(mark: Extract<ReviewMark, { kind: "likely" | "possible" }>) {
    const placeholder = mark.replacement ?? nextManualPlaceholder(mark.category, effectiveMapping);
    setManualMapping((current) => ({ ...current, [placeholder]: mark.value }));
    replaceReviewRange(mark.start, mark.end, placeholder);
  }

  function onChooseReplacementIdentifier(mark: Extract<ReviewMark, { kind: "replaced" }>, value: string) {
    const placeholder = placeholderFromIdentifierOption(value, effectiveMapping);
    if (!placeholder || placeholder === mark.placeholder) {
      return;
    }

    setManualMapping((current) => ({ ...current, [placeholder]: mark.original }));
    setReviewText((current) => current.split(mark.placeholder).join(placeholder));
    setCopied(false);
  }

  function onAddManualIdentifier() {
    const value = manualIdentifierText.trim();
    if (!value || !reviewText.includes(value)) {
      return;
    }

    const placeholder = nextManualPlaceholder(manualIdentifierCategory, effectiveMapping);
    setManualMapping((current) => ({ ...current, [placeholder]: value }));
    setReviewText((current) => current.split(value).join(placeholder));
    setManualIdentifierText("");
    setCopied(false);
  }

  return (
    <main className="admin-pseudonymiser">
      <header className="admin-pseudonymiser__topbar">
        <div>
          <h1>{title}</h1>
          <p>Deterministic pseudonymisation for adult social care notes before LLM use.</p>
        </div>
        <div className={`admin-pseudonymiser__risk admin-pseudonymiser__risk--${result?.findings.riskLevel.toLowerCase() ?? "low"}`}>
          {result?.findings.safeToSend === false ? <AlertTriangle size={18} /> : <ShieldCheck size={18} />}
          <span>{result ? result.findings.reason : "Ready"}</span>
        </div>
      </header>

      {error && (
        <div className="admin-pseudonymiser__error" role="alert">
          <AlertTriangle size={18} />
          <span>{error}</span>
        </div>
      )}

      <section className="admin-pseudonymiser__grid">
        <div className="admin-pseudonymiser__pane">
          <div className="admin-pseudonymiser__pane-header">
            <h2>Original text</h2>
            <span>{input.length.toLocaleString()} chars</span>
          </div>
          <textarea
            value={input}
            onChange={(event) => setInput(event.target.value)}
            spellCheck={false}
            aria-label="Original care text"
          />
          <div className="admin-pseudonymiser__actions">
            <button className="admin-pseudonymiser__button admin-pseudonymiser__button--primary" onClick={onPseudonymise} disabled={!canPseudonymise}>
              {busy ? <Loader2 className="admin-pseudonymiser__spin" size={18} /> : <Wand2 size={18} />}
              <span>Pseudonymise</span>
            </button>
            <button className="admin-pseudonymiser__button" onClick={onClear} type="button">
              <Eraser size={18} />
              <span>Clear</span>
            </button>
          </div>
        </div>

        <div className="admin-pseudonymiser__pane">
          <div className="admin-pseudonymiser__pane-header">
            <h2>Pseudonymised text</h2>
            <span>{reviewText.length.toLocaleString()} chars</span>
          </div>
          <HighlightedOutput
            text={reviewText}
            marks={reviewMarks}
            placeholder="Pseudonymised output will appear here."
            onApplySuggestion={onApplySuggestion}
            onRevertReplacement={onRevertReplacement}
            onChooseReplacementIdentifier={onChooseReplacementIdentifier}
            identifierOptions={identifierOptions}
          />
          {reviewText && (
            <>
              <div className="admin-pseudonymiser__manual" aria-label="Add missed identifier">
                <div className="admin-pseudonymiser__inline-heading">
                  <strong>Missed text</strong>
                  <span>Type an identifier the recogniser missed, then assign a placeholder.</span>
                </div>
                <div className="admin-pseudonymiser__manual-row">
                  <input
                    value={manualIdentifierText}
                    onChange={(event) => setManualIdentifierText(event.target.value)}
                    placeholder="Type missed text"
                    aria-label="Missed identifier text"
                  />
                  <select
                    value={manualIdentifierCategory}
                    onChange={(event) => setManualIdentifierCategory(event.target.value)}
                    aria-label="Identifier category"
                  >
                    {PLACEHOLDER_CATEGORY_OPTIONS.map((category) => (
                      <option key={category} value={category}>
                        {category.replace(/_/g, " ")}
                      </option>
                    ))}
                  </select>
                  <button
                    className="admin-pseudonymiser__button"
                    type="button"
                    onClick={onAddManualIdentifier}
                    disabled={!manualIdentifierText.trim() || !reviewText.includes(manualIdentifierText.trim())}
                  >
                    <Wand2 size={16} />
                    <span>Add identifier</span>
                  </button>
                </div>
              </div>

              <label className="admin-pseudonymiser__field">
                <span>Preferred client name</span>
                <input
                  value={clientPreferredName}
                  onChange={(event) => setClientPreferredName(event.target.value)}
                  placeholder="Preferred name"
                  aria-label="Client preferred name"
                />
              </label>

              <div className="admin-pseudonymiser__actions">
                <button className="admin-pseudonymiser__button admin-pseudonymiser__button--primary" onClick={onCopyForLlm} disabled={!reviewText}>
                  <Clipboard size={18} />
                  <span>{copied ? "Copied for LLM" : "Copy for LLM"}</span>
                </button>
              </div>
            </>
          )}
        </div>
      </section>

      {result && (
        <section className="admin-pseudonymiser__summary" aria-label="Detection summary">
          <div>
            <h2>Detection summary</h2>
            <p>{pendingSuggestionCount > 0 ? "Review highlighted suggestions before copying." : "Only reversible placeholders remain highlighted."}</p>
          </div>
          <div className="admin-pseudonymiser__counts">
            {(result.counts.length ? result.counts : [{ entity_type: "No findings yet", count: 0 }]).map((item) => (
              <div className="admin-pseudonymiser__count" key={item.entity_type}>
                <span>{item.entity_type.replace(/_/g, " ")}</span>
                <strong>{item.count}</strong>
              </div>
            ))}
          </div>
        </section>
      )}

      <section className="admin-pseudonymiser__restore" aria-label="Restore placeholders">
        <div className="admin-pseudonymiser__section-heading">
          <div>
            <h2>Restore</h2>
            <p>Paste LLM output with placeholders on the left and the restored text appears on the right.</p>
          </div>
          {placeholderText.trim() && (
            <span>
              {restorePreview.restoredCount.toLocaleString()} restored
              {restorePreview.unresolvedCount > 0 ? `, ${restorePreview.unresolvedCount.toLocaleString()} unresolved` : ""}
            </span>
          )}
        </div>
        <div className="admin-pseudonymiser__grid admin-pseudonymiser__grid--restore">
          <div className="admin-pseudonymiser__pane">
            <div className="admin-pseudonymiser__pane-header">
              <h2>Text with placeholders</h2>
              <span>{placeholderText.length.toLocaleString()} chars</span>
            </div>
            <textarea
              value={placeholderText}
              onChange={(event) => setPlaceholderText(event.target.value)}
              spellCheck={false}
              aria-label="Text with pseudonymised placeholders"
              placeholder="Paste text that uses the same placeholders."
            />
          </div>

          <div className="admin-pseudonymiser__pane">
            <div className="admin-pseudonymiser__pane-header">
              <h2>Restored text</h2>
              <span>{restorePreview.restoredCount.toLocaleString()} restored</span>
            </div>
            <textarea
              value={restorePreview.text}
              readOnly
              spellCheck={false}
              aria-label="Restored care text"
              placeholder="Restored output will appear here."
            />
            <div className="admin-pseudonymiser__actions">
              <button className="admin-pseudonymiser__button" onClick={onCopyRestored} disabled={!restorePreview.text} type="button">
                <Clipboard size={18} />
                <span>Copy restored</span>
              </button>
              <button className="admin-pseudonymiser__button" onClick={() => setPlaceholderText("")} disabled={!placeholderText} type="button">
                <RotateCcw size={18} />
                <span>Reset restore</span>
              </button>
            </div>
          </div>
        </div>
      </section>
    </main>
  );
}

function HighlightedOutput({
  text,
  marks,
  placeholder,
  onApplySuggestion,
  onRevertReplacement,
  onChooseReplacementIdentifier,
  identifierOptions
}: {
  text: string;
  marks: ReviewMark[];
  placeholder: string;
  onApplySuggestion: (mark: Extract<ReviewMark, { kind: "likely" | "possible" }>) => void;
  onRevertReplacement: (mark: Extract<ReviewMark, { kind: "replaced" }>) => void;
  onChooseReplacementIdentifier: (mark: Extract<ReviewMark, { kind: "replaced" }>, value: string) => void;
  identifierOptions: string[];
}) {
  if (!text) {
    return (
      <div className="admin-pseudonymiser__output admin-pseudonymiser__output--placeholder" role="textbox" aria-label="Pseudonymised care note" aria-readonly="true">
        {placeholder}
      </div>
    );
  }

  let cursor = 0;
  const parts: Array<JSX.Element | string> = [];

  marks.forEach((mark) => {
    if (mark.start > cursor) {
      parts.push(text.slice(cursor, mark.start));
    }
    parts.push(
      renderReviewMark(
        text,
        mark,
        onApplySuggestion,
        onRevertReplacement,
        onChooseReplacementIdentifier,
        identifierOptions
      )
    );
    cursor = mark.end;
  });

  if (cursor < text.length) {
    parts.push(text.slice(cursor));
  }

  return (
    <div className="admin-pseudonymiser__output" role="textbox" aria-label="Pseudonymised care note" aria-readonly="true">
      {parts}
    </div>
  );
}

function renderReviewMark(
  text: string,
  mark: ReviewMark,
  onApplySuggestion: (mark: Extract<ReviewMark, { kind: "likely" | "possible" }>) => void,
  onRevertReplacement: (mark: Extract<ReviewMark, { kind: "replaced" }>) => void,
  onChooseReplacementIdentifier: (mark: Extract<ReviewMark, { kind: "replaced" }>, value: string) => void,
  identifierOptions: string[]
): JSX.Element {
  const content = text.slice(mark.start, mark.end);

  if (mark.kind === "replaced") {
    return (
      <span className="admin-pseudonymiser__replacement" key={`${mark.start}-${mark.end}-${mark.placeholder}`}>
        <button
          className="admin-pseudonymiser__highlight admin-pseudonymiser__highlight--replaced"
          onClick={() => onRevertReplacement(mark)}
          title={`Revert: ${mark.original}`}
          type="button"
        >
          {content}
        </button>
        <select
          value={mark.placeholder}
          onChange={(event) => onChooseReplacementIdentifier(mark, event.target.value)}
          onClick={(event) => event.stopPropagation()}
          aria-label={`Change identifier for ${mark.placeholder}`}
        >
          {identifierOptions.map((option) => (
            <option key={option} value={option}>
              {identifierOptionLabel(option)}
            </option>
          ))}
        </select>
      </span>
    );
  }

  return (
    <button
      className={`admin-pseudonymiser__highlight admin-pseudonymiser__highlight--${mark.kind}`}
      key={`${mark.start}-${mark.end}-${mark.reason}`}
      onClick={() => onApplySuggestion(mark)}
      title={`Pseudonymise: ${mark.reason}`}
      type="button"
    >
      {content}
    </button>
  );
}

function buildReviewMarks(
  text: string,
  originalPseudonymisedText: string,
  mapping: Record<string, string>,
  residualSpans: ResidualSpan[]
): ReviewMark[] {
  if (!text) {
    return [];
  }

  const replacedMarks = Object.entries(mapping).flatMap(([placeholder, original]) =>
    findAllOccurrences(text, placeholder).map((start) => ({
      kind: "replaced" as const,
      start,
      end: start + placeholder.length,
      placeholder,
      original
    }))
  );
  const revertedMarks = Object.entries(mapping).flatMap(([placeholder, original]) =>
    findAllOccurrences(text, original).map((start) => ({
      kind: "likely" as const,
      start,
      end: start + original.length,
      category: placeholderCategoryValue(placeholder),
      reason: "Previously pseudonymised value was restored",
      value: original,
      replacement: placeholder
    }))
  );
  const suggestionMarks = residualSpans.flatMap((span) => {
    const value = originalPseudonymisedText.slice(span.start, span.end);
    if (!value || mapping[value]) {
      return [];
    }

    return findAllOccurrences(text, value).map((start) => ({
      kind: span.severity === "direct" ? ("likely" as const) : ("possible" as const),
      start,
      end: start + value.length,
      category: span.category,
      reason: span.reason,
      value
    }));
  });

  return normaliseReviewMarks([...replacedMarks, ...revertedMarks, ...suggestionMarks], text.length);
}

function findAllOccurrences(text: string, value: string): number[] {
  const starts: number[] = [];
  let cursor = 0;
  while (value && cursor < text.length) {
    const start = text.indexOf(value, cursor);
    if (start === -1) {
      break;
    }
    starts.push(start);
    cursor = start + value.length;
  }
  return starts;
}

function normaliseReviewMarks(marks: ReviewMark[], textLength: number): ReviewMark[] {
  const ordered = marks
    .filter((mark) => mark.start >= 0 && mark.end <= textLength && mark.start < mark.end)
    .sort(
      (left, right) =>
        left.start - right.start ||
        reviewMarkRank(left.kind) - reviewMarkRank(right.kind) ||
        right.end - left.end
    );
  const accepted: ReviewMark[] = [];
  let cursor = 0;

  ordered.forEach((mark) => {
    if (mark.start >= cursor) {
      accepted.push(mark);
      cursor = mark.end;
    }
  });

  return accepted;
}

function reviewMarkRank(kind: ReviewMark["kind"]): number {
  if (kind === "replaced") {
    return 0;
  }
  return kind === "likely" ? 1 : 2;
}

function buildEffectiveMapping(
  baseMapping: Record<string, string>,
  manualMapping: Record<string, string>,
  clientPreferredName: string
): Record<string, string> {
  const merged = { ...baseMapping, ...manualMapping };

  Object.entries(merged).forEach(([placeholder, original]) => {
    const match = placeholder.match(/^\[([A-Z_]+)_(\d{3})\]$/);
    if (!match) {
      return;
    }

    const [, category, index] = match;
    const { firstName, surname } = splitPersonName(original);
    const preferredName =
      category === "CLIENT" ? normalisePreferredName(clientPreferredName) ?? firstName : firstName;

    if (preferredName) {
      merged[`[${category}_PREFERRED_NAME_${index}]`] = preferredName;
      merged[`[${category}_FIRST_NAME_${index}]`] = preferredName;
    }
    if (surname) {
      merged[`[${category}_SURNAME_${index}]`] = surname;
    }
  });

  return merged;
}

function defaultPreferredName(mapping: Record<string, string>): string {
  const clientName = mapping["[CLIENT_001]"];
  if (!clientName) {
    return "";
  }
  const { firstName } = splitPersonName(clientName);
  return firstName ?? "";
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

function normalisePreferredName(value: string): string | null {
  const trimmed = value.trim().replace(/\s+/g, " ");
  return trimmed || null;
}

function nextManualPlaceholder(category: string, mapping: Record<string, string>): string {
  const cleanCategory = placeholderSafeCategory(category);
  const highest = Object.keys(mapping).reduce((max, placeholder) => {
    const match = placeholder.match(/^\[([A-Z_]+)_(\d{3})\]$/);
    if (!match || match[1] !== cleanCategory) {
      return max;
    }
    return Math.max(max, Number(match[2]));
  }, 0);
  return `[${cleanCategory}_${String(highest + 1).padStart(3, "0")}]`;
}

function buildIdentifierOptions(mapping: Record<string, string>): string[] {
  const existing = Object.keys(mapping).filter((placeholder) =>
    /^\[[A-Z_]+_\d{3}\]$/.test(placeholder) &&
    !placeholder.includes("_PREFERRED_NAME_") &&
    !placeholder.includes("_FIRST_NAME_") &&
    !placeholder.includes("_SURNAME_")
  );
  const newOptions = PLACEHOLDER_CATEGORY_OPTIONS.map((category) => `${category}_NEW`);
  return [...existing, ...newOptions].sort((left, right) =>
    identifierOptionLabel(left).localeCompare(identifierOptionLabel(right))
  );
}

function placeholderFromIdentifierOption(value: string, mapping: Record<string, string>): string | null {
  if (value.endsWith("_NEW")) {
    return nextManualPlaceholder(value.slice(0, -4), mapping);
  }
  return value.startsWith("[") ? value : null;
}

function identifierOptionLabel(value: string): string {
  return value.startsWith("[") ? value.slice(1, -1) : value;
}

function placeholderSafeCategory(category: string): string {
  const clean = category.toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  return clean || "IDENTIFIER";
}

function placeholderCategoryValue(placeholder: string): string {
  const match = placeholder.match(/^\[([A-Z_]+)_\d{3}\]$/);
  return match?.[1] ?? "IDENTIFIER";
}

function normaliseResidualSpans(spans: ResidualSpan[], textLength: number): ResidualSpan[] {
  const ordered = spans
    .filter((span) => span.start >= 0 && span.end <= textLength && span.start < span.end)
    .sort(
      (left, right) =>
        left.start - right.start ||
        severityRank(left.severity) - severityRank(right.severity) ||
        right.end - left.end
    );
  const accepted: ResidualSpan[] = [];
  let cursor = 0;

  ordered.forEach((span) => {
    if (span.start >= cursor) {
      accepted.push(span);
      cursor = span.end;
    }
  });

  return accepted;
}

function severityRank(severity: ResidualSpan["severity"]): number {
  return severity === "direct" ? 0 : 1;
}

export default PseudonymiserAdminPage;
