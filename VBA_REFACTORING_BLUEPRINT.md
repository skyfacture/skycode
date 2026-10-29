# Skycode VBA Refactoring Blueprint (Access + SolidWorks Composer)

## Purpose
This document defines the required approach for safely refactoring the Skycode VBA automation stack for SolidWorks Composer.

This codebase is behavior-critical and environment-sensitive. The baseline assumption is:
- Automation host must remain Access VBA.
- Composer API behavior is partly undocumented and empirically discovered.
- Existing scripts encode months of trial-and-error and must be treated as production knowledge.

## Non-Negotiable Constraints
- Do not migrate runtime automation from Access VBA to Python as part of refactoring work.
- Preserve COM call order unless a controlled experiment proves equivalence.
- Preserve string path handling conventions used for actor addressing and selection.
- Preserve existing logging behavior and message content unless explicitly versioned.
- Do not combine functional and stylistic changes in the same change set.

## Refactoring Mission
- Improve readability, modularity, and maintainability.
- Keep all external behavior identical by default.
- Produce repeatable, low-risk changes with rollback points.

## Required Inputs Before Any Refactor
- Current semantic inventory in CODE_SEMANTIC_INDEX.md and CODE_SEMANTIC_INDEX.json.
- A baseline run log from Access for each script being changed.
- A target script list and exact expected outcomes.
- A change journal entry template for recording hypotheses and results.

## Canonical Workflow
1. Select one script as the unit of change.
2. Capture baseline behavior in Access.
3. Identify one refactor objective only.
4. Apply the smallest possible code transformation.
5. Re-run in Access using the same test data and sequence.
6. Diff logs and outcomes against baseline.
7. Accept only if behavior is unchanged or intentionally improved with evidence.

## Access-First Execution Standard
- All verification runs are executed from Access VBA.
- All COM objects are created and released using the same lifecycle pattern as baseline scripts.
- Global and module state must be initialized exactly once per run path.
- Error handling must always log enough context to reproduce failures.

## Safe Change Types (Allowed)
- Extracting duplicated logic into local helper procedures without changing call order.
- Normalizing constant values and magic strings into named constants.
- Splitting large procedures into deterministic private helpers.
- Clarifying variable names where scope and semantics remain unchanged.
- Standardizing log wrappers while preserving log text semantics.

## High-Risk Change Types (Require Explicit Experiment Plan)
- Reordering Composer API calls.
- Replacing selection/filter logic used for actor traversal.
- Altering path normalization or token parsing rules.
- Changing view naming, numbering, or creation sequencing.
- Introducing new host/runtime or replacing Access orchestration.

## Refactor Design Rules
- Rule 1: One behavior axis per change. Do not refactor naming and selection logic together.
- Rule 2: Keep procedure signatures stable unless callers are updated in the same change set.
- Rule 3: Preserve defensive checks around null/missing Composer entities.
- Rule 4: Maintain deterministic traversal order for assemblies, subassemblies, and parts.
- Rule 5: Keep metadata sanitization pure and side-effect free.

## Code Organization Standard
- Entry procedure at top of module.
- Validation and configuration helpers next.
- Traversal and transformation helpers next.
- Logging and utility helpers last.
- Shared constants grouped in a single section.

## Logging Standard
- Start message includes script name and timestamp.
- Each major phase logs begin and end.
- Each failure logs actor/view identifier and operation name.
- End message includes summary counts and failure counts.
- Log wording must remain stable to allow baseline diffing.

## Error Handling Standard
- Use structured error boundaries around Composer calls.
- Never swallow errors silently.
- Convert recoverable errors into warnings with context.
- Bubble unrecoverable errors after logging operation and inputs.
- Ensure cleanup executes even when errors occur.

## Test Strategy
- Use an Access harness that can run a chosen module repeatably.
- Keep at least one stable regression model for baseline comparison.
- Validate these dimensions for each run:
  - View count and naming
  - Actor selection set
  - Metadata write outcomes
  - Pivot/frame transform outcomes
  - Log signature consistency

## Definition Of Done For A Refactor
- Code compiles and runs in Access.
- Baseline and refactored outputs are equivalent for agreed scenarios.
- Logs show no new warnings/errors unless intentionally introduced and documented.
- Semantic index has been regenerated.
- Change journal contains rationale, risks, and validation evidence.

## Required Deliverables Per Change
- Before/after diff focused on one module objective.
- Access run logs for baseline and refactored runs.
- Validation checklist with pass/fail per scenario.
- Rollback note describing exact revert path.

## Recommended Module Partitioning Pattern
- modSkycode_Entry: orchestrates main flow.
- modSkycode_Traversal: assembly/part iteration and filtering.
- modSkycode_Metadata: read/write/sanitize property operations.
- modSkycode_ViewOps: create/rename/select view operations.
- modSkycode_TransformOps: pivot/frame/local coordinate operations.
- modSkycode_Logging: log sink and formatting helpers.

## Change Journal Template
- Script/module:
- Refactor objective:
- Baseline behavior summary:
- Change summary:
- Risks considered:
- Validation scenarios:
- Result:
- Rollback plan:
- Follow-up tasks:

## Practical Guidance For Agents
- Read CODE_SEMANTIC_INDEX.md first to locate candidate modules and helper reuse.
- Prefer edits that reduce duplication without changing orchestration sequence.
- Do not infer undocumented Composer behavior from docs alone.
- Treat historical scripts in archive as empirical references when behavior is unclear.

## Operating Principle
Refactor for clarity, not novelty. Preserve proven behavior first, then improve structure incrementally with measurable evidence.
