---
name: distill
description: End-of-session routine for any project (code repo or skill) - consolidate what this session learned into the project's instruction file with targeted edits that replace rather than append, respecting files that only change through their own command. Use when the user says "distill", "/distill", "wrap up", "end of session" or asks to capture this session's standards or lessons.
---

# distill - consolidate this session's lessons into the instructions

Goal: an upgraded instruction file, not a longer one. Lessons sharpen or replace the rules they touch;
they are never appended, and rules they do not touch are left exactly as they are. Some lessons may already
be in the file from earlier in the session: check before adding, never repeat.

## 1. Find the instruction file(s)

Whatever instructs future work on this project:
- Code repo: `AGENTS.md` / `CLAUDE.md` (and any docs they point to).
- A skill (one built or used this session): its `SKILL.md` and its reference files.
- Other projects: the project instructions, style guide or README that sessions follow.
If it is unclear which file, ask once.

## 2. Respect protected files

Read each target for an update rule: a line starting `Updates:` (e.g. `Updates: only via /awp-learn`), or any
statement that the file may only change through a named command or process (e.g. "never edit SKILL.md by hand").
- Protected: do not edit it. List its lessons as a short block ready to pass to that command, and tell the
  user to run it.
- Not protected: edit as below.

## 3. Collect the lessons

From the whole conversation:
- The owner's corrections (names, wording, formats, behaviour rejected or changed).
- Conventions introduced or changed.
- Mistakes or quirks found in testing that a future session could repeat.
- Rules now wrong or out of step with the actual work (check the work, not memory).
Skip one-off details unless they set a pattern. Drop lessons the file already states.
No lessons left: change nothing and say so.

## 4. Edit only what the lessons touch

- Find the rule each lesson belongs to and change it in place: replace the wording, do not add a second rule
  beside it. Add a new line only when no rule covers the lesson, in the section a reader would look in.
- Remove or merge a rule only when this session's lessons make it wrong, superseded or a duplicate.
- Every other rule keeps its wording, order and structure. Do not reword, reorder or tidy untouched sections.
- One line per rule where possible; no history, no "previously", no changelog or "lessons learned" section.
- Size is a warning sign, not a target: never cut a correct, non-redundant rule to stay short. If the file
  grew, each added line must carry a new rule; say which.
- Whole-file de-bloat (rewording, restructuring) only when the user asks for it.

## 5. Enforce only with what exists

- The project has a checker (lint script, CI, tests, validator): add checks for new rules that can be
  checked mechanically, and run it.
- It has none (documents, skills, writing projects): do not create one, and do not add lint or CI steps
  to its instructions.

## 6. Ship it the way the project ships

- Git repo: branch from the latest main, commit, open a PR; merge only if the user allowed merging in this
  session, otherwise give the link.
- A user skill in Claude: propose the change to its SKILL.md with the skill-proposal tool.
- Anything else: update the file in place and send it.

## 7. Report (short)

- Lead with: done or not, where (PR / file), lines before -> after.
- Each change: the lesson, and the rule it added, replaced or removed (max 5 bullets).
- Each lesson left out because the file already covers it, with where (file and section), so a "no changes" result shows what was checked.
- Protected files skipped, with the lessons to pass to their own command.
- Outstanding actions for the user carried over from the session.
