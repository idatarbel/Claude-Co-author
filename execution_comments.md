### #12. Post-move health check (Dropbox -> C:\GitHub)
- The move turned out to be a no-op for this repo. It has zero build tooling:
  no package.json, no node_modules, no requirements.txt, no venv, no
  __pycache__, no .clasp.json, no git hooks, submodules, or worktrees. The only
  Python file (`word-online/assets/generate_icons.py`) resolves its output
  directory from `os.path.dirname(os.path.abspath(__file__))`, so it is
  location-independent by construction.
- `.git/config` contains only the HTTPS origin URL. Git stores no absolute path
  to the working tree in a normal (non-worktree) clone, which is why moving the
  folder was safe. A repo using `git worktree` or a `.git` *file* pointing at a
  gitdir would have needed `git worktree repair`.
- Verified deploy by md5-comparing every served GitHub Pages asset against the
  local file. `Code.js` initially read as a mismatch (41304 local vs 40202
  live). That is CRLF-vs-LF, not corruption: the working copy is checked out
  with CRLF while the committed blob and the Pages-served bytes are LF, a
  1102-byte delta matching the line count. Comparing `git show
  HEAD:word-online/Code.js` to the live URL gives identical md5s. Worth
  remembering — comparing a Windows working file to a served blob will always
  false-positive on text files.
- Old Dropbox copy left in place deliberately. Confirmed clean (no uncommitted
  changes, no unpushed commits, no stashes) so deleting it loses nothing, but
  deletion is Dan's call.
