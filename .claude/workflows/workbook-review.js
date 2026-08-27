export const meta = {
  name: 'workbook-review',
  description: 'Review a workbook and its VBA, fix findings, approve each fix independently, then test',
  whenToUse: 'When an Excel workbook (or its exported VBA under src/) needs a full review-fix-approve-test pass rather than a single edit. Pass {workbook, dimensions, approvers, maxFiles} as args.',
  phases: [
    { title: 'Inventory', detail: 'map sheets, formulas, VBA modules and the current lint baseline' },
    { title: 'Review', detail: 'one read-only reviewer per dimension, in parallel' },
    { title: 'Fix', detail: 'one fixer per file, minimal diffs, lint before returning' },
    { title: 'Approve', detail: 'independent approvers vote on each fix diff' },
    { title: 'Rework', detail: 'one round to address blocking approver findings' },
    { title: 'Test', detail: 'static, structural and behavioural gates' },
  ],
}

// ---------------------------------------------------------------- configuration
const workbook = (args && args.workbook) || ''
const APPROVERS = (args && args.approvers) || 2
const MAX_FILES = (args && args.maxFiles) || 4
const DIMENSIONS = (args && args.dimensions) || [
  {
    key: 'vba-correctness',
    prompt: 'Review dimension: VBA correctness. Wrong results, range/array off-by-one, Variant coercion, integer overflow, missing Set, functions that swallow failure and return a plausible value.',
  },
  {
    key: 'error-handling',
    prompt: 'Review dimension: error handling and application state. Open On Error Resume Next, ScreenUpdating/EnableEvents/Calculation/DisplayAlerts not restored on the failure path, handlers that hide the original Err.',
  },
  {
    key: 'structure-formulas',
    prompt: 'Review dimension: workbook structure and formulas. Cached error cells, volatile and fragile functions, external links, hidden sheets holding live logic, defined names pointing at nothing, hardcoded ranges.',
  },
  {
    key: 'maintainability-safety',
    prompt: 'Review dimension: maintainability and macro safety. Untyped variables, duplicated logic, magic numbers, sheets by index, plus Shell/WScript.Shell/Kill/SendKeys and anything auto-running on open with unvalidated input.',
  },
]

const FINDINGS_SCHEMA = {
  type: 'object',
  required: ['dimension', 'findings'],
  properties: {
    dimension: { type: 'string' },
    findings: {
      type: 'array',
      items: {
        type: 'object',
        required: ['id', 'file', 'line', 'severity', 'title', 'detail', 'failure_scenario'],
        properties: {
          id: { type: 'string' },
          file: { type: 'string' },
          line: { type: 'integer' },
          severity: { type: 'string', enum: ['error', 'warn', 'info'] },
          title: { type: 'string' },
          detail: { type: 'string' },
          failure_scenario: { type: 'string' },
          proposed_fix: { type: 'string' },
        },
      },
    },
  },
}

const FIX_SCHEMA = {
  type: 'object',
  required: ['file', 'fixed', 'deferred'],
  properties: {
    file: { type: 'string' },
    fixed: {
      type: 'array',
      items: {
        type: 'object',
        required: ['id', 'what_changed', 'why_this_is_the_cause'],
        properties: {
          id: { type: 'string' },
          what_changed: { type: 'string' },
          why_this_is_the_cause: { type: 'string' },
          test_added: { type: 'string' },
        },
      },
    },
    deferred: {
      type: 'array',
      items: {
        type: 'object',
        required: ['id', 'reason'],
        properties: { id: { type: 'string' }, reason: { type: 'string' } },
      },
    },
    lint_before: { type: 'integer' },
    lint_after: { type: 'integer' },
    notes: { type: 'string' },
  },
}

const VERDICT_SCHEMA = {
  type: 'object',
  required: ['file', 'verdict', 'confidence', 'reasoning'],
  properties: {
    file: { type: 'string' },
    verdict: { type: 'string', enum: ['approve', 'needs_changes', 'reject'] },
    confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
    reasoning: { type: 'string' },
    blocking_issues: { type: 'array', items: { type: 'string' } },
    required_changes: { type: 'array', items: { type: 'string' } },
    collateral_risk: { type: 'string' },
  },
}

const TEST_SCHEMA = {
  type: 'object',
  required: ['verdict', 'summary'],
  properties: {
    static_gate: { type: 'object' },
    structural_gate: { type: 'object' },
    behavioural_gate: { type: 'object' },
    tests_added: { type: 'array', items: { type: 'string' } },
    uncovered_fixes: { type: 'array', items: { type: 'string' } },
    failing_files: { type: 'array', items: { type: 'string' } },
    verdict: { type: 'string', enum: ['green', 'green_static_only', 'red', 'blocked'] },
    summary: { type: 'string' },
  },
}

const target = workbook ? `Workbook: ${workbook}. ` : 'No workbook file given - review the exported VBA source under src/ and tests/. '

// ------------------------------------------------------------------- inventory
phase('Inventory')
const inventory = await agent(
  `${target}Produce the baseline for a workbook review.
` +
    `Run: python3 tools/wb_inventory.py ${workbook || '<no workbook - skip>'} (skip if no workbook file exists)
` +
    `Run: python3 tools/vba_lint.py src tests --json
` +
    `List the modules under src/ and tests/ with their procedure names.
` +
    `Return a compact map: sheets and their state, formula and error-cell totals, external links, defined names, ` +
    `module -> procedures, and the current lint counts by severity. Facts only, no recommendations.`,
  { label: 'inventory', phase: 'Inventory', agentType: 'wb-reviewer' },
)

// ---------------------------------------------------------------------- review
phase('Review')
const reviews = await parallel(
  DIMENSIONS.map((d) => () =>
    agent(
      `${target}${d.prompt}

Baseline from the inventory pass:
${typeof inventory === 'string' ? inventory : JSON.stringify(inventory)}

Report only findings in your dimension. Set dimension to "${d.key}" and give each finding an id prefixed "${d.key}-".`,
      { label: `review:${d.key}`, phase: 'Review', schema: FINDINGS_SCHEMA, agentType: 'wb-reviewer' },
    ),
  ),
)

// Barrier is deliberate: dedup needs every reviewer's findings before any fixer
// starts, otherwise two fixers touch the same line for the same defect.
const seen = new Set()
const findings = []
for (const r of reviews.filter(Boolean)) {
  for (const f of r.findings || []) {
    const key = `${f.file}:${f.line}:${(f.title || '').slice(0, 40).toLowerCase()}`
    if (seen.has(key)) continue
    seen.add(key)
    findings.push(f)
  }
}

const rank = { error: 0, warn: 1, info: 2 }
findings.sort((a, b) => (rank[a.severity] ?? 3) - (rank[b.severity] ?? 3))
log(`${findings.length} unique finding(s) after dedup: ` +
  `${findings.filter((f) => f.severity === 'error').length} error, ` +
  `${findings.filter((f) => f.severity === 'warn').length} warn, ` +
  `${findings.filter((f) => f.severity === 'info').length} info`)

if (!findings.length) {
  return { findings: [], fixes: [], verdict: 'no_findings', inventory }
}

// Group by file so parallel fixers never touch the same file.
const byFile = new Map()
for (const f of findings) {
  if (!byFile.has(f.file)) byFile.set(f.file, [])
  byFile.get(f.file).push(f)
}
let batches = [...byFile.entries()].map(([file, items]) => ({ file, items }))
batches.sort((a, b) => (rank[a.items[0].severity] ?? 3) - (rank[b.items[0].severity] ?? 3))
if (batches.length > MAX_FILES) {
  const dropped = batches.slice(MAX_FILES)
  log(`CAP: fixing ${MAX_FILES} of ${batches.length} files this run. Not fixed: ` +
    dropped.map((b) => `${b.file} (${b.items.length} finding(s))`).join(', '))
  batches = batches.slice(0, MAX_FILES)
}

// ------------------------------------------------- fix -> approve -> rework
const describe = (items) =>
  items.map((f) => `- [${f.id}] line ${f.line} (${f.severity}) ${f.title}\n  ${f.detail}\n  fails when: ${f.failure_scenario}\n  suggested: ${f.proposed_fix || 'your call'}`).join('\n')

const results = await pipeline(
  batches,
  (batch) =>
    agent(
      `${target}Fix these findings in ${batch.file} and nothing else:

${describe(batch.items)}

Minimal diff. Add a regression test per behavioural fix and register it in modTestRunner.RunAllTests. ` +
        `Run python3 tools/vba_lint.py ${batch.file} tests before returning and report the counts.`,
      { label: `fix:${batch.file}`, phase: 'Fix', schema: FIX_SCHEMA, agentType: 'wb-fixer' },
    ),

  // Independent approvers, then one rework round if any of them blocks.
  async (fix, batch) => {
    if (!fix) return { batch, fix: null, verdicts: [], approved: false }

    const review = async (label) =>
      (await parallel(
        Array.from({ length: APPROVERS }, (unused, i) => () =>
          agent(
            `Approve or reject the fix to ${batch.file}.

Findings it claims to fix:
${describe(batch.items)}

Fixer's own account (verify it against the diff, do not trust it):
${JSON.stringify(fix)}

Read the real diff with: git diff -- ${batch.file} tests/
You are approver ${i + 1} of ${APPROVERS} and must reach your verdict independently.`,
            { label: `${label}:${batch.file}#${i + 1}`, phase: 'Approve', schema: VERDICT_SCHEMA, agentType: 'wb-approver' },
          ),
        ),
      )).filter(Boolean)

    let verdicts = await review('approve')
    let blocking = verdicts.filter((v) => v.verdict !== 'approve')

    if (blocking.length) {
      const asks = blocking.flatMap((v) => [...(v.blocking_issues || []), ...(v.required_changes || []), v.reasoning])
      log(`${batch.file}: ${blocking.length}/${verdicts.length} approver(s) blocked - one rework round`)
      const reworked = await agent(
        `Your fix to ${batch.file} was blocked in review. Address every point below, ` +
          `keeping the diff minimal. If a point requires a decision you cannot make, revert that part ` +
          `and report it as deferred rather than guessing.

${asks.map((a) => `- ${a}`).join('\n')}

Original findings:
${describe(batch.items)}`,
        { label: `rework:${batch.file}`, phase: 'Rework', schema: FIX_SCHEMA, agentType: 'wb-fixer' },
      )
      verdicts = await review('reapprove')
      blocking = verdicts.filter((v) => v.verdict !== 'approve')
      return {
        batch,
        fix: reworked || fix,
        verdicts,
        approved: blocking.length === 0,
        reworked: true,
      }
    }

    return { batch, fix, verdicts, approved: true, reworked: false }
  },
)

const settled = results.filter(Boolean)
const approved = settled.filter((r) => r.approved)
const rejected = settled.filter((r) => !r.approved)
log(`${approved.length} file(s) approved, ${rejected.length} still blocked`)

if (rejected.length) {
  log(`Blocked (left unfixed in the working tree - revert or decide manually): ` +
    rejected.map((r) => r.batch.file).join(', '))
}

// ------------------------------------------------------------------------ test
phase('Test')
const fixSummary = approved
  .flatMap((r) => (r.fix.fixed || []).map((f) => `${r.batch.file}: [${f.id}] ${f.what_changed} (test: ${f.test_added || 'none'})`))
  .join('\n')

let testReport = await agent(
  `${target}Run every applicable gate and report honestly.

Approved fixes in this round:
${fixSummary || '(none)'}

Baseline inventory for comparison:
${typeof inventory === 'string' ? inventory : JSON.stringify(inventory)}

If a fix has no test that would fail against the pre-fix code, write one under tests/ and re-run. ` +
    `Set failing_files to the src/ files whose gates failed.`,
  { label: 'test-gates', phase: 'Test', schema: TEST_SCHEMA, agentType: 'wb-tester' },
)

// One repair round if the gates are red - fixers own src/, the tester never does.
if (testReport && testReport.verdict === 'red' && (testReport.failing_files || []).length) {
  const repairs = (testReport.failing_files || []).slice(0, MAX_FILES)
  log(`Gates red - one repair round on: ${repairs.join(', ')}`)
  await parallel(
    repairs.map((file) => () =>
      agent(
        `The quality gates failed after this round's fixes. Repair ${file}.

Tester report:
${JSON.stringify(testReport)}

Fix the cause. Do not weaken, skip, or delete a test, and do not touch files other than ${file}.`,
        { label: `repair:${file}`, phase: 'Test', schema: FIX_SCHEMA, agentType: 'wb-fixer' },
      ),
    ),
  )
  testReport = await agent(
    `${target}Re-run every applicable gate after the repair round and report the result.`,
    { label: 're-test-gates', phase: 'Test', schema: TEST_SCHEMA, agentType: 'wb-tester' },
  )
}

return {
  inventory,
  findings_total: findings.length,
  findings_by_severity: {
    error: findings.filter((f) => f.severity === 'error').length,
    warn: findings.filter((f) => f.severity === 'warn').length,
    info: findings.filter((f) => f.severity === 'info').length,
  },
  files_fixed: approved.map((r) => r.batch.file),
  files_blocked: rejected.map((r) => ({
    file: r.batch.file,
    reasons: r.verdicts.filter((v) => v.verdict !== 'approve').map((v) => v.reasoning),
  })),
  deferred: settled.flatMap((r) => (r.fix && r.fix.deferred) || []),
  reworked: settled.filter((r) => r.reworked).map((r) => r.batch.file),
  test_report: testReport,
}
