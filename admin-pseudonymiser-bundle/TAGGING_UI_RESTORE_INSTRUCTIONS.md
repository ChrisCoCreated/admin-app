# Restore Interactive Tagging UI

These instructions are for Codex in the new admin web-app workspace. The current new implementation has lost the interactive tagging layer: blue/yellow/orange highlighting, inline dropdown recategorisation, manual missed-identifier tagging, and the mapping-aware restore flow. Restore those behaviours without adding Ollama, Tauri, desktop-sidecar code, or any local-model setup UI.

## Source Of Truth

Use the original portable bundle as the reference implementation:

- Frontend page: `admin-pseudonymiser-bundle/frontend/PseudonymiserAdminPage.tsx`
- Styles: `admin-pseudonymiser-bundle/frontend/pseudonymiser-admin.css`
- API client: `admin-pseudonymiser-bundle/frontend/pseudonymiserApi.ts`
- Types: `admin-pseudonymiser-bundle/frontend/pseudonymiserTypes.ts`
- Restore helper: `admin-pseudonymiser-bundle/frontend/deanonymise.ts`
- Backend router: `admin-pseudonymiser-bundle/backend/pseudonymiser_core/router.py`

If those files are available in the target workspace, diff the new implementation against them. The important thing is not visual exactness; the important thing is preserving the state model and interactions below.

## What Is Missing

The stripped implementation likely renders the pseudonymised output as plain text or a plain textarea. That loses the review model.

The page must render pseudonymised output through an interactive component equivalent to `HighlightedOutput`, not a simple textarea. This output must show:

- Blue highlights for already pseudonymised placeholder tags such as `[CLIENT_001]`.
- Orange highlights for likely direct identifiers that remain, based on `findings.residualSpans` where `severity === "direct"`.
- Yellow highlights for possible indirect identifiers that remain, based on `findings.residualSpans` where `severity === "indirect"`.
- A tiny inline dropdown beside every blue placeholder so the user can recategorise it to an existing placeholder or create a new category placeholder.
- Click-to-apply behaviour for orange/yellow suggestions so missed text can be converted into a placeholder.
- Click-to-revert behaviour for blue placeholders so a user can temporarily restore the original value for review.
- Manual missed-text tagging controls under the output.
- Preferred client-name aliases such as `[CLIENT_PREFERRED_NAME_001]`, `[CLIENT_FIRST_NAME_001]`, and `[CLIENT_SURNAME_001]`.
- Restore mode that uses the local mapping, including those aliases, to put original values back into LLM output.

## Backend Contract

The frontend must call:

```http
POST /api/pseudonymiser/pseudonymise
Content-Type: application/json

{ "text": "..." }
```

Expected response shape:

```ts
type PseudonymisationResponse = {
  pseudonymised_text: string;
  mapping: Record<string, string>;
  findings: {
    entities_detected: string[];
    risks_preserved: string[];
    clinical_details_preserved: string[];
    warnings: string[];
    riskLevel: "HIGH" | "MEDIUM" | "LOW";
    safeToSend: boolean;
    directIdentifiers: string[];
    indirectIdentifiers: string[];
    residualSpans: Array<{
      start: number;
      end: number;
      severity: "direct" | "indirect";
      category: string;
      reason: string;
    }>;
    reason: string;
  };
  entities: Array<{
    entity_type: string;
    start: number;
    end: number;
    score: number;
    replacement: string;
  }>;
  counts: Array<{ entity_type: string; count: number }>;
};
```

Do not add `use_ollama`, `ollama_model`, `/ollama/check`, `/ollama/status`, or any external LLM call to this feature.

## Required Frontend State

Reintroduce these state values or direct equivalents:

```ts
const [input, setInput] = useState(initialText);
const [placeholderText, setPlaceholderText] = useState("");
const [result, setResult] = useState<PseudonymisationResponse | null>(null);
const [reviewText, setReviewText] = useState("");
const [manualMapping, setManualMapping] = useState<Record<string, string>>({});
const [clientPreferredName, setClientPreferredName] = useState("");
const [manualIdentifierText, setManualIdentifierText] = useState("");
const [manualIdentifierCategory, setManualIdentifierCategory] = useState("PROFESSIONAL");
```

Important distinctions:

- `result.mapping` is the engine-produced mapping.
- `manualMapping` is the user-produced mapping for missed identifiers and recategorisation.
- `effectiveMapping` is `{ ...result.mapping, ...manualMapping }` plus generated first-name/preferred-name/surname aliases.
- `reviewText` is editable only through controlled actions: pseudonymise, revert, apply suggestion, recategorise, and manual missed-identifier replacement.
- Do not overwrite `reviewText` from `result.pseudonymised_text` after the user starts reviewing.

## Review Mark Model

Implement this union or equivalent:

```ts
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
```

Build marks from three sources:

1. Replaced placeholders:
   - For every `[PLACEHOLDER_001] -> original` entry in `effectiveMapping`, find every occurrence of the placeholder in `reviewText`.
   - Create `kind: "replaced"` marks.

2. Reverted original values:
   - For every mapping entry, find occurrences of the original value in `reviewText`.
   - Create `kind: "likely"` marks with `replacement` set to the original placeholder.
   - This means if the user clicks a blue placeholder to revert it, the restored raw identifier becomes orange and can be clicked to pseudonymise again.

3. Residual spans from the backend:
   - Read `result.findings.residualSpans`.
   - Slice the value from `result.pseudonymised_text` using `span.start` and `span.end`.
   - Find that value in current `reviewText`.
   - Create `kind: "likely"` for `severity === "direct"` and `kind: "possible"` for `severity === "indirect"`.

Sort and de-overlap marks by:

```ts
left.start - right.start ||
reviewMarkRank(left.kind) - reviewMarkRank(right.kind) ||
right.end - left.end
```

Where `replaced` ranks before `likely`, and `likely` ranks before `possible`. Accept only non-overlapping marks in order.

## Required Interactions

### Pseudonymise

When the user clicks Pseudonymise:

1. POST the original text to `/api/pseudonymiser/pseudonymise`.
2. Set `result` to the response.
3. Set `reviewText` to `response.pseudonymised_text`.
4. Clear `manualMapping`.
5. Set `clientPreferredName` to the first name from `[CLIENT_001]` unless already set.

### Blue Highlight: Revert

For every blue placeholder highlight:

- Render a small button containing the placeholder text.
- On click, replace only that mark range with the original raw value from the mapping.
- After replacing it, mark-building should detect that raw value and show it as orange because it matches a known original value.

### Blue Dropdown: Recategorise

Render a select immediately beside every blue placeholder.

Options must include:

- Existing base placeholders in the effective mapping, for example `[CLIENT_001]`, `[STAFF_001]`, `[PROFESSIONAL_001]`.
- New category options using the sentinel format `${CATEGORY}_NEW`, for every category in:

```ts
[
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
]
```

When the user selects:

- Existing placeholder: replace all occurrences of the old placeholder in `reviewText` with the chosen placeholder and map the chosen placeholder to the original value.
- New category: compute the next stable number for that category using existing mapping keys, for example `[RELATIVE_002]`; replace all occurrences of the old placeholder with it; add it to `manualMapping`.

The dropdown can be visually narrow, but it must remain keyboard-accessible and have an accessible label.

### Orange/Yellow Highlight: Apply Suggestion

For orange/yellow highlights:

- Render as clickable inline buttons.
- On click, create a placeholder:
  - If the mark has `replacement`, use that replacement.
  - Otherwise generate the next placeholder for `mark.category`.
- Add `{ [placeholder]: mark.value }` to `manualMapping`.
- Replace that exact range in `reviewText` with the placeholder.

### Manual Missed Identifier

Under the output, include:

- Text input for missed text.
- Category dropdown using the same category list.
- `Add identifier` button.

On click:

- Trim the missed text.
- Require that the text exists in `reviewText`.
- Generate the next placeholder for the chosen category.
- Add it to `manualMapping`.
- Replace all occurrences of that exact text in `reviewText` with the new placeholder.
- Clear the input.

### Preferred Client Name

Include a preferred client-name input. `effectiveMapping` must add aliases for base placeholders:

- `[CATEGORY_PREFERRED_NAME_001]`
- `[CATEGORY_FIRST_NAME_001]`
- `[CATEGORY_SURNAME_001]`

For client placeholders, preferred name comes from the preferred-name input if populated; otherwise first name from the original value. For non-client placeholders, first-name aliases can use the first token from the original value.

### Copy For LLM

Copy only:

1. The placeholder-preservation instructions.
2. The current `reviewText`.

Never copy or send `mapping`, `manualMapping`, or raw source text. Preserve these rules exactly:

```text
Instructions for the LLM:
Keep every placeholder tag exactly as written, including the square brackets.
Do not rename, renumber, remove, expand, or merge any text inside brackets such as [CLIENT_001] or [CLIENT_PREFERRED_NAME_001].
In body text, use [CLIENT_PREFERRED_NAME_001] instead of [CLIENT_001] whenever the client is mentioned naturally in a sentence.
Use [CLIENT_001] only for headings, labels, tables, or when the full placeholder is explicitly needed.
If a preferred-name or surname variant is needed, convert [CLIENT_001] to [CLIENT_PREFERRED_NAME_001] or [CLIENT_SURNAME_001] with the same number.
Do not invent new numbering. Keep the numeric suffix aligned with the original placeholder.
You may rewrite the surrounding narrative, but preserve all bracketed tags verbatim.

Text:
```

### Restore

Restore mode must:

- Accept placeholder-bearing LLM output.
- Replace bracketed placeholders from `effectiveMapping`.
- Support aliases such as `[CLIENT_PREFERRED_NAME_001]`, `[CLIENT_FIRST_NAME_001]`, and `[CLIENT_SURNAME_001]`.
- Show restored count and unresolved count.
- Leave unknown placeholders unchanged.

Use `deanonymiseText` from the bundle if available.

## Required CSS / Visual Contract

The implementation must visibly distinguish the three mark types:

```css
.review-highlight.replaced,
.admin-pseudonymiser__highlight--replaced {
  outline: 1px solid #4b86c5;
  background: #dcedff;
}

.review-highlight.likely,
.admin-pseudonymiser__highlight--likely {
  outline: 1px solid #d26b23;
  background: #ffe0c4;
}

.review-highlight.possible,
.admin-pseudonymiser__highlight--possible {
  outline: 1px solid #b28a0c;
  background: #fff1b8;
}
```

Required semantics:

- Blue = already replaced placeholder; click to revert; dropdown to recategorise.
- Orange = likely direct identifier; click to pseudonymise.
- Yellow = possible indirect identifier; click to pseudonymise.

Do not render the reviewed output in a plain textarea because inline buttons and dropdowns cannot be embedded inside textarea content. Use a scrollable `div` with `white-space: pre-wrap` and `role="textbox"`/`aria-readonly="true"` or an equivalent rich text implementation.

## Summary / Review Panel

Add a detection summary that shows:

- Entity counts from `result.counts`.
- Review status:
  - `needs review` when orange/yellow marks exist.
  - `ready` when only blue placeholders remain.
- A short list of review marks with category and reason.

This is not decorative. It is a safety feature because it tells the user whether direct/indirect residual identifiers still need attention.

## Tests To Add

Add frontend tests using the target app’s test stack. If none exists, use React Testing Library plus the project’s existing test runner.

Minimum tests:

1. Pseudonymise response renders blue placeholders:
   - Mock `/api/pseudonymiser/pseudonymise`.
   - Response includes `[CLIENT_001]`.
   - Assert a blue/replaced highlight button exists with text `[CLIENT_001]`.

2. Blue placeholder can be reverted:
   - Click `[CLIENT_001]`.
   - Assert original raw value appears.
   - Assert it is now represented as a likely/orange mark.

3. Orange residual can be applied:
   - Mock `residualSpans` for a direct identifier left in `pseudonymised_text`.
   - Click the orange mark.
   - Assert it is replaced with a generated placeholder.
   - Assert manual mapping is used by restore.

4. Dropdown recategorises a placeholder:
   - Open/change the inline select for `[CLIENT_001]`.
   - Choose `RELATIVE_NEW` or an existing placeholder.
   - Assert all occurrences in `reviewText` update to the new placeholder.

5. Manual missed identifier:
   - Enter missed text that exists in review text.
   - Select category.
   - Click `Add identifier`.
   - Assert all occurrences are replaced with the next stable placeholder.

6. Copy for LLM:
   - Mock clipboard.
   - Assert copied text includes placeholder instructions and `reviewText`.
   - Assert copied text does not include raw mapping values.

7. Restore:
   - Given mapping `{ "[CLIENT_001]": "Paulette Crawley" }`.
   - Paste `[CLIENT_PREFERRED_NAME_001] met [CLIENT_SURNAME_001]`.
   - Assert first/preferred name and surname resolve correctly.

Also add a browser smoke test if the app uses Playwright/Cypress:

- Load the admin pseudonymiser route.
- Enter: `Paulette Crawley feels safe at Lindau. Carrie is brilliant. Discussed Martin Tyrell and selling her properties.`
- Click Pseudonymise.
- Confirm visible blue highlights and inline dropdowns.
- Confirm no Ollama UI appears anywhere on the page.

## Implementation Order

1. Locate the stripped pseudonymiser page/component in the target admin app.
2. Locate the API route it calls and verify the response contains `mapping`, `findings.residualSpans`, and `counts`.
3. Restore the `ReviewMark` model, `buildReviewMarks`, `normaliseReviewMarks`, `HighlightedOutput`, and `renderReviewMark`.
4. Restore state and handlers for:
   - `effectiveMapping`
   - `manualMapping`
   - `clientPreferredName`
   - `onRevertReplacement`
   - `onApplySuggestion`
   - `onChooseReplacementIdentifier`
   - `onAddManualIdentifier`
5. Replace the plain pseudonymised output textarea with the rich highlighted output component.
6. Restore CSS for blue/orange/yellow highlights and inline select.
7. Restore detection summary and restore mode.
8. Add tests.
9. Run type-check, formatter, unit tests, and browser smoke test.

## Acceptance Criteria

The work is complete only when:

- Pseudonymised placeholders render blue.
- Every blue placeholder has an inline dropdown to change identifier/category.
- Clicking a blue placeholder reverts it to the original value.
- Reverted original values become orange and can be pseudonymised again.
- Direct residual identifiers render orange.
- Indirect residual identifiers render yellow.
- Clicking orange/yellow marks replaces the highlighted text with a stable placeholder.
- Manual missed-identifier tagging works and uses stable numbering.
- Preferred client-name aliases are generated and restored.
- Copy for LLM includes instructions plus pseudonymised text only.
- Restore mode resolves base placeholders and preferred-name/first-name/surname aliases.
- The page has no Ollama, Tauri, local model, or desktop-engine setup UI.
- Tests cover the tagging behaviours above.
