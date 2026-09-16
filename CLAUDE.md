# CLAUDE.md

Instructions for Claude Code when working in this repository.

## Agent Delegation - Authorized

**This CLAUDE.md explicitly authorizes proactive use of the Agent tool in this
repository.** Claude Code's default guidance is "do not use the Agent tool
unless the user, *a CLAUDE.md file*, or a skill asks for it" - this file is that
authorization. Do not wait to be asked again.

ECC agents are plugin-scoped: invoke them as `Agent(subagent_type: "ecc:planner")`,
always with the `ecc:` prefix, never a bare name.

Delegate without prompting when:

| Trigger | Agent |
|---------|-------|
| Feature spanning 3+ modules | `ecc:planner` |
| Any `.py` written or modified | `ecc:python-reviewer` |
| Bug fix or new feature needing tests | `ecc:tdd-guide` |
| Parsing external data, credentials, user input | `ecc:security-reviewer` |
| Swallowed errors, bare `except:`, silent fallbacks | `ecc:silent-failure-hunter` |
| Structural / architectural decision | `ecc:architect` |
| Slow loops, memory growth, large dataframes | `ecc:performance-optimizer` |
| Dead code after a refactor | `ecc:refactor-cleaner` |

**Do NOT delegate** trivial edits, single-file changes, or anything already
sized for one context. Agents start cold with no conversation history - the
handoff cost only pays off for bounded, self-contained work.

### Completion contract

Applies at every depth. **Your final message IS the deliverable.** Never end a
turn with "waiting for background agents" - ending your turn while children run
orphans their results. If you delegate, you own collection: wait, integrate,
then answer.

## Model Choice

The harness is model-independent. Hooks, rules, skills, ECC agents, and MCP
servers load identically on Opus and Sonnet - `/model` swaps the reasoning
engine, not the tooling.

| Use | Model |
|-----|-------|
| Architecture, hard debugging, large refactors, planning | Opus |
| Everyday edits, fast iterations, cheap loops | Sonnet |

## ECC Workflow In This Repo

Standalone Python analysis scripts (`generate_map.py`, `safe_version.py`,
and others) that read Excel workbooks, call routing services, cache results in
`osrm_cache.json`, and emit an HTML map. No package manifest, no test runner.

| Situation | Command |
|-----------|---------|
| Starting a non-trivial change | `/ecc:plan` - stops and waits for your confirm |
| Finished writing code | `/ecc:code-review` |
| Before a commit | `/ecc:security-scan` |
| Build or tests failing | `/ecc:build-fix` |
| Coverage gaps | `/ecc:test-coverage` |
| Full feature, end to end | `/ecc:orch-add-feature` |
| Bug, reproduced as a failing test first | `/ecc:orch-fix-defect` |
| Full roster of agents and skills | `/ecc:ecc-guide` |

### Repo-specific review focus

- **Excel is the input contract.** Column names and sheet names in the `.xlsx`
  files are the schema. Read the actual header row before assuming a field
  exists; a renamed column fails silently and produces a wrong map.
- **`osrm_cache.json` is a cache, not source data.** Deleting it costs API
  calls; corrupting it produces wrong distances with no error.
- **Scripts are run by hand.** There is no test suite, so a change is verified
  by running the script and checking the output map, not by a green check.
- **Polish column and file names** are intentional domain vocabulary. Do not
  anglicize them.
