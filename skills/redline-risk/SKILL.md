---
name: redline-risk
description: 'Contract redline analysis with directional risk scoring. Extracts tracked changes from Word documents and produces a color-coded risk assessment showing which changes favor or hurt your position.'
---

# Redline Risk

Analyze contract redlines with directional risk assessment. When activated, Redline Risk extracts tracked changes from a Word document, assesses each change's impact relative to your party's position, and produces a color-coded output document showing which edits help or hurt you.

## Activation

When the user invokes this skill (e.g., "use redline-risk", "analyze this redline", "score these changes", "what did they change in this contract"), respond with:

> **Redline Risk is active.** I'll analyze the tracked changes and produce a color-coded Word document showing which edits favor or hurt your position.

Then follow the workflow below.

## Supported Sources

- **Tracked changes documents:** Word documents (.docx) with tracked changes/revisions
- **Before/after comparison:** Two versions of the same contract for semantic diffing

## Tool Location

The Redline Risk Python utility is located at:
```
<plugin_dir>/skills/redline-risk/tools/redline_risk.py
```

To find the actual path, run:
```bash
find ~/.copilot/installed-plugins -name "redline_risk.py" -path "*/redline-risk/*" 2>/dev/null
```

If not found there, check the project directory for the redline-risk repo.

## First-Run Setup

Before first use, check that dependencies are installed:

```bash
python3 <path-to>/redline_risk.py setup-check
```

If anything is missing, run the setup script from the redline-risk plugin directory:
```bash
bash <path-to>/setup.sh
```

Or install manually:
```bash
pip3 install lxml python-docx pillow
```

## Workflow

Follow these steps exactly. The order matters.

### Step 1: Identify the party

**CRITICAL:** Before any analysis, determine which party the user represents.

If the user hasn't specified, ask explicitly:
> "Which party do you represent in this contract?"

Once identified, determine the party mapping: which entity name in the contract corresponds to the user's party and which corresponds to the counterparty?

Restate this mapping clearly before proceeding:
> "You = [Entity A], Counterparty = [Entity B]"

This step is mandatory. The assessment is directional -- every judgment depends on knowing whose side you're on.

### Step 2: Extract changes

Extract all tracked changes from the document:

```bash
python3 redline_risk.py extract-changes --source "<path-to-contract.docx>"
```

**Or**, if the user provides two separate files (original and modified versions):

```bash
python3 redline_risk.py extract-changes --before "<original.docx>" --after "<modified.docx>"
```

Read the output JSON carefully. Understand every change: what was inserted, deleted, or modified, and where it appears in the contract.

**Error handling:**
- If no changes are found, tell the user: "No tracked changes found in this document. This usually means changes were already accepted. If you have the original version, I can compare the two files -- provide both using '--before original.docx --after revised.docx'."

### Step 3: Filter and group changes

Call the filter-changes command to remove formatting-only edits and group related changes:

```bash
python3 redline_risk.py filter-changes --changes <(echo '<json-from-step-2>')
```

Or save the JSON from step 2 to a temporary file and pass the path.

The filter command applies deterministic grouping:
1. **Proximity grouping**: Insert/delete pairs in the same paragraph, by the same author, within the same timestamp window are grouped as replacements
2. **Move pairing**: moveFrom/moveTo pairs matched by revision ID are grouped as single move operations
3. **Formatting isolation**: Changes with only formatting differences (no text change) are separated
4. **Cosmetic detection**: Changes where normalized text is identical are flagged as cosmetic

Review the grouped output. The deterministic algorithm may miss connections or incorrectly group unrelated changes. Refine as needed:
- **Merge groups** that are part of the same logical change but span paragraph boundaries (e.g., deleting a clause in Section 5 and inserting a replacement in Section 5.1)
- **Split groups** that were proximity-grouped but are actually independent changes
- **Reclassify** changes that the deterministic step misidentified

### Step 4: Read the full contract

Extract and read the full text of the contract to understand the overall agreement context. You can use the same source document or extract text with python-docx.

**CRITICAL:** Individual changes can only be assessed in context of the whole deal. Read the full contract once for overall structure, key provisions, and party roles.

For very long contracts (100+ pages) that exceed comfortable context:
1. Read the full contract once for overall structure and key provisions
2. Take notes on the major sections, party obligations, risk allocation, and any unusual provisions
3. Assess changes section by section with the relevant section text plus your notes from the full read

### Step 5: Classify and assess each grouped change

For each grouped change, determine:

#### 1. Category (classify first)

Classify the change before assessing its direction. This forces decomposed judgment.

Categories:
1. **scope_obligation**: Changes to what a party must do, how much, or to what standard (effort levels, performance metrics, deliverables, timelines)
2. **risk_allocation**: Changes to indemnification, liability caps, liability exclusions, insurance, representations and warranties
3. **procedural_administrative**: Changes to notice requirements, cure periods, reporting obligations, approval processes, governing law, dispute resolution
4. **definitional**: Changes to defined terms, including scope narrowing/broadening of key definitions
5. **remedial_enforcement**: Changes to termination rights, remedies, damages, specific performance, injunctive relief

#### 2. Direction

Within that category, does this favor the user's party, the counterparty, or neither?

- **for**: Favors the user's party
- **against**: Favors the counterparty
- **neutral**: No material impact on either party

#### 3. Impact (0.0 to 1.0)

How significant is this change? Consider both the legal effect and the practical effect given the rest of the contract.

- **0.7-1.0**: High impact (deal-critical, major risk shift, fundamental obligation change)
- **0.4-0.69**: Medium impact (meaningful change that affects rights or obligations)
- **0.0-0.39**: Low impact (minor adjustment, clarification, administrative change)

#### 4. Confidence

How confident are you in this assessment?

- **high**: Clear legal effect, standard pattern, no ambiguity
- **medium**: Some uncertainty about interaction with other provisions or practical effect
- **low**: Depends on business context you don't have, ambiguous language, or multiple plausible interpretations

**Mark low confidence when:**
- The change's effect depends on business context you don't have (e.g., whether the cap amount is meaningful depends on deal size)
- The change interacts with other provisions in ways you can't fully trace (note which provisions and why)
- The language is ambiguous and could be interpreted multiple ways (state the interpretations)
- The change appears in a highly negotiated or jurisdiction-specific area where standard assessments may not apply
- The change could favor either party depending on how a dispute plays out
- You identified multiple plausible legal interpretations and can't determine which is more likely

#### 5. Explanation

What is the legal effect of this specific edit? Be concrete and cite the actual language.

**Good example:**
> "Narrows the indemnification scope from 'all losses' to 'direct damages,' excluding consequential and incidental damages -- see 'direct damages arising under this Agreement' replacing 'all claims arising under or related to this Agreement' in Section 11.1."

**Bad example:**
> "May affect indemnification obligations."

#### 6. Conditional note (optional)

If the effect depends on another provision or on business context not in the document, say so explicitly.

**Example:**
> "The practical effect depends on whether the $5M cap in Section 11.3 is meaningful relative to the deal size, which is not stated in the agreement."

## AI Assessment Guidelines

These guidelines shape how you evaluate each change:

### Common Patterns

Recognize these standard negotiation moves and their typical impact:

- **"Best efforts" → "commercially reasonable efforts"**: Standard weakening of obligation (medium impact, against the obligated party). "Commercially reasonable" permits balancing performance against cost.
- **"Reasonable efforts" → "best efforts"**: Strengthening of obligation (medium impact, favors the non-obligated party)
- **"Shall" → "may"**: Converting obligation to discretion (high impact)
- **Adding "material" before "breach"**: Raising the threshold for breach (medium impact)
- **Adding a carve-out or exception**: Narrowing a protection (impact depends on breadth)
- **Shortening a time period** (notice, cure): Reducing the other party's protection (medium impact)
- **Adding a cap on liability**: Limiting exposure (high impact if uncapped before)
- **Changing "and" → "or" in conditions**: Lowering the threshold (can be high impact)
- **Adding "sole discretion" qualifier**: Removing objective standards (medium-high impact)
- **Broadening force majeure definition**: Weakening the non-declaring party's position (medium impact)
- **Adding "including but not limited to"**: Expanding scope of a provision (impact varies)
- **Deleting "survive termination" clause**: Limiting post-termination obligations (can be high impact)

### Avoid These Mistakes

1. **Don't score a change as "against" just because the clause it's in is unfavorable.** The question is whether *this edit* made it more or less unfavorable.

2. **Don't treat cosmetic edits as meaningful.** Reformatting, renumbering, restating without substantive change should be caught by filter-changes, but if any slip through, score them as neutral/zero impact.

3. **Don't over-rate boilerplate changes.** "Governed by the laws of [State X]" changing to "[State Y]" matters, but it's a known negotiation point, not a hidden risk.

4. **Don't under-rate definitional changes.** If the counterparty modified a defined term used 40 times in the agreement, that single edit has compounding effects throughout. Count the usages and note the ripple effect.

5. **Don't confuse clause risk with edit risk.** The assessment baseline is the previous draft, not a market-standard template. If the contract is already heavily one-sided against you, a neutral edit isn't a win -- but it's still neutral in terms of what this redline changed.

### Step 6: Write the summary

After assessing all changes, characterize the overall redline posture:

1. **Classify** as aggressive, defensive, balanced, or cleanup
2. **Explain** the classification in 2-3 sentences
3. **List** the top 3 areas of concern

**Do not produce a single "overall balance" number.** A handful of small favorable edits do not offset one liability cap change. Instead, call out the most significant changes and what they mean for the user's position.

### Step 7: Build the output document

Call the build command with the complete assessment data:

```bash
python3 redline_risk.py build \
  --source "<original-contract.docx>" \
  --output ~/Desktop/redline-risk-analysis.docx \
  --title "Redline Risk Analysis" \
  --subtitle "Vendor Services Agreement -- Changes from CounterpartyCo" \
  --party "Client" \
  --party-mapping '{"Client": "Licensee", "CounterpartyCo": "Licensor"}' \
  --changes '[
    {
      "group_id": "g1",
      "section": "4.2",
      "section_title": "Performance Obligations",
      "before_text": "best efforts",
      "after_text": "commercially reasonable efforts",
      "category": "scope_obligation",
      "impact": 0.6,
      "direction": "against",
      "confidence": "high",
      "explanation": "Replaces \"best efforts\" with \"commercially reasonable efforts.\" Commercially reasonable is a lower standard -- it permits the provider to balance performance against cost, giving more latitude to underperform without breach.",
      "conditional_note": null
    }
  ]'
```

The output document includes:
1. **Title page** with document title, subtitle, date, party perspective, and party mapping
2. **Summary section** with total change count, priority review list, posture assessment, and low-confidence changes
3. **Detailed analysis by severity** (sorted by impact score)
4. **Detailed analysis by document order** (sorted by section number)

Save to the user's Desktop and tell them the filename.

## Error Handling

- **No tracked changes found**: Tell the user the document appears to have no tracked changes (already accepted) and offer to compare two separate versions.
- **Ambiguous party labels**: Stop and ask. Do not guess which entity corresponds to the user's party.
- **Incomprehensible changes**: Skip garbled text or corrupted revision markup and note in the summary as "N changes could not be parsed."
- **Missing dependencies**: If setup-check fails, guide the user to run setup.sh or install manually.

## Notes

- The tool parses raw OOXML because python-docx silently discards all tracked change markup
- Move operations (moveFrom/moveTo) are paired by revision ID and assessed based on whether the relocation changes legal effect
- Formatting-only changes are separated and listed in a cleanup section
- Low-confidence assessments are marked with lighter color shading and a [LOW CONFIDENCE] indicator
- The output uses color-coded backgrounds: green (favorable), gray (neutral), red (unfavorable)
