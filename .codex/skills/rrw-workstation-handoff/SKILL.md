---
name: rrw-workstation-handoff
description: Safely hand off the steel-structure RLP/RAW/RRW research project between two computers through the private GitHub repository HJQ0216/rrw. Use when the user says they are switching computers, leaving the current computer, continuing on another computer, taking over the project again, synchronizing the RRW project, updating HANDOFF.md before a transfer, or checking whether the local main branch is ready to hand off or resume.
---

# RRW Workstation Handoff

Use the private repository `https://github.com/HJQ0216/rrw.git`, branch `main`, and root-level `HANDOFF.md` as the handoff sources of truth.

## Determine the direction

Classify the request before changing Git state:

- **Leave this computer**: the user finished or paused work here and wants the other computer to continue.
- **Resume on this computer**: the other computer pushed work and the user wants this computer to continue.
- **First setup**: this computer does not yet contain the repository.

If the direction is not clear, ask one concise question. Do not push or pull until it is clear.

## Common preflight

1. Locate the repository with `git rev-parse --show-toplevel`.
2. Confirm the remote exactly matches `https://github.com/HJQ0216/rrw.git` or the equivalent SSH URL owned by `HJQ0216/rrw`.
3. Confirm the intended branch is `main`.
4. Run `git status --short --branch` and inspect the result before any mutation.
5. Never use force push, `git reset --hard`, destructive checkout, automatic conflict resolution, or deletion as part of a handoff.

## Leave this computer

1. Inspect the work completed in the current session and the changed files.
2. Update `HANDOFF.md` so it accurately records:
   - project goal and current research stage;
   - work completed in this round;
   - files changed;
   - tests or checks performed;
   - unresolved questions and decisions awaiting confirmation;
   - the exact next action for the receiving computer;
   - current branch, remote, and synchronization state.
3. Preserve existing handoff details that remain valid. Do not rewrite verified thesis parameters or formulas from memory.
4. Check candidate files for obvious secrets such as `.env`, private keys, access tokens, credentials, or password exports. Stop and ask before staging any suspected secret.
5. Run `git add .`, then review `git diff --cached --stat` and `git status --short`.
6. If there are no staged changes, do not create an empty commit. Continue to synchronization verification.
7. Create a concise commit message describing the completed work. Do not use a vague message unless the user supplied it.
8. Push `main` to `origin`.
9. Verify all of the following:
   - `git status --short --branch` shows no uncommitted change;
   - local `main` tracks `origin/main`;
   - local `HEAD` and `origin/main` resolve to the same commit;
   - the push command succeeded.
10. Report the final commit hash, remote URL, and the first instruction the receiving Codex should follow.

Treat an explicit request to use this skill for leaving the computer as authorization to commit and push the project changes to the existing private repository. Request platform approval when required by the execution environment.

## Resume on this computer

1. Run `git status --short --branch` before contacting the remote.
2. If the working tree contains local changes, do not pull, stash, discard, or overwrite them automatically. Summarize the changes and ask the user how they should be handled.
3. If clean, fetch `origin` and compare local `HEAD`, `main`, and `origin/main`.
4. Pull only with `git pull --ff-only origin main`.
5. If fast-forward is impossible or the branches diverged, stop. Report both commit tips and ask for direction; do not merge or rebase automatically.
6. Read `HANDOFF.md` completely after pulling.
7. Review at least the latest five commits and inspect the files named in the newest handoff entry.
8. Confirm that the main deliverables referenced by `HANDOFF.md` exist before claiming that the project is ready.
9. Summarize:
   - what the other computer completed;
   - the current branch and commit;
   - whether the working tree is clean;
   - the next unfinished task;
   - any decision that still requires user confirmation.

Treat an explicit request to use this skill for resuming as authorization to fetch and fast-forward pull from the existing private repository. Do not infer permission to discard local work.

## First setup on another computer

1. If the destination folder is not specified, ask where to clone the repository.
2. Clone with `git clone https://github.com/HJQ0216/rrw.git`.
3. Open the cloned repository as the Codex workspace.
4. Confirm the branch is `main` and the working tree is clean.
5. Read `HANDOFF.md` completely and follow the **Resume on this computer** review steps.
6. Confirm the project skill folders under `.codex/skills` are present. They are versioned with the repository.

## Conflict prevention

- Keep only one computer actively editing the project at a time.
- Finish every handoff with a successful push before starting on the other computer.
- Begin every resumed session with a clean fast-forward pull.
- Do not edit the same Word or Excel file on both computers before synchronizing; Git cannot reliably merge these binary files.
- Do not claim synchronization merely because `git push` or `git pull` was attempted. Verify commit equality.

## Failure handling

- **Authentication failure**: ask the user to sign in to GitHub on that computer, then retry.
- **Push rejected**: fetch and compare histories; do not force push.
- **Dirty receiving tree**: stop before pulling and preserve all local changes.
- **Diverged branches**: report the commit hashes and wait for an explicit merge/rebase decision.
- **Wrong remote or repository**: stop without pushing and ask the user to confirm the intended repository.
- **Missing `HANDOFF.md`**: do not invent prior state. Reconstruct it from committed files and recent Git history, then ask the user to verify uncertain research decisions.

## Final response format

Lead with whether handoff or resume succeeded. Include:

- computer direction;
- branch and commit hash;
- synchronization result;
- remote repository link;
- working-tree status;
- one concrete next step.
