# Updating Your Fork from Parent Repository

This guide explains how to keep your fork (`omnidox/DesktopCommanderMCP`) up to date with the parent repository (`wonderwhy-er/DesktopCommanderMCP`) while preserving your custom Word document reading feature.

## Table of Contents
- [Understanding the Setup](#understanding-the-setup)
- [Method 1: Using GitHub Desktop (Recommended)](#method-1-using-github-desktop-recommended)
- [Method 2: Using Command Line](#method-2-using-command-line)
- [Checking for Available Updates](#checking-for-available-updates)
- [Troubleshooting](#troubleshooting)
- [Update Schedule Recommendations](#update-schedule-recommendations)

---

## Understanding the Setup

### Two Repositories:

**`origin`** = Your fork (`omnidox/DesktopCommanderMCP`)
- Your personal repository on GitHub
- Contains your custom Word document feature
- Where you push your changes

**`upstream`** = Parent repository (`wonderwhy-er/DesktopCommanderMCP`)
- The original Desktop Commander MCP repository
- Maintained by wonderwhy-er
- Where new features and updates come from

### Your Custom Features:
- **Word Document Support**: Ability to read `.docx` files using the mammoth library
- Located in: `package.json`, `src/tools/filesystem.ts`, `src/server.ts`

---

## Method 1: Using GitHub Desktop (Recommended)

### Step-by-Step Instructions:

#### Step 1: Fetch Updates
1. Open **GitHub Desktop**
2. Make sure you're on the **`main`** branch
3. Click **Repository** → **Fetch origin** (or click the "Fetch origin" button in the top bar)
   - This downloads information about new commits from both your fork and the parent

#### Step 2: Check if Updates Exist
1. Go to **Branch** → **Merge into current branch...**
2. Look for **`upstream/main`** in the list
3. If it shows commits below it, updates are available!
4. If it says "This branch is up to date", no updates needed ✅

#### Step 3: Merge Updates (if available)
1. With **`upstream/main`** selected, click **"Merge upstream/main into main"**
2. If there are conflicts (rare), GitHub Desktop will show them - resolve if needed
3. The merge will be applied to your local `main` branch

#### Step 4: Push to Your Fork
1. Click **"Push origin"** button at the top
2. This saves the updated code to your GitHub fork
3. Done! Your fork is now up to date ✅

### Visual Summary (GitHub Desktop):
```
Fetch origin → Branch Menu → Merge into current branch →
Select upstream/main → Merge → Push origin
```

---

## Method 2: Using Command Line

### Prerequisites:
Make sure you're in the repository directory:
```bash
cd C:\Users\omnid\GitHub\DesktopCommanderMCP
```

Or on Mac/Linux:
```bash
cd ~/path/to/DesktopCommanderMCP
```

### Step-by-Step Commands:

#### Step 1: Ensure You're on Main Branch
```bash
git checkout main
```

#### Step 2: Fetch Updates from Parent
```bash
git fetch upstream main
```
This downloads the latest commits from the parent repository.

#### Step 3: Check What's New (Optional but Recommended)
```bash
git log --oneline HEAD..upstream/main
```
- **No output** = You're up to date! ✅
- **Shows commits** = Updates are available

#### Step 4: Merge Updates
```bash
git merge upstream/main
```
This applies the parent's new commits to your main branch.

**If you see conflicts:**
- Git will tell you which files have conflicts
- Open those files and resolve the conflicts manually
- Then run: `git add .` and `git commit`

#### Step 5: Push to Your Fork
```bash
git push origin main
```
This saves the merged changes to your GitHub fork.

### Quick One-Liner (Advanced):
```bash
git checkout main && git fetch upstream main && git merge upstream/main && git push origin main
```

---

## Checking for Available Updates

### Using GitHub Desktop:
1. Click **Repository** → **Fetch origin**
2. Go to **Branch** → **Merge into current branch**
3. Select **upstream/main**
4. Look below - if commits are listed, updates are available

### Using Command Line:
```bash
# Fetch latest info
git fetch upstream main

# Check for new commits
git log --oneline HEAD..upstream/main
```

**Interpreting Results:**
- **No output** = Up to date! ✅
- **Commits listed** = Updates available! Run merge and push

**Example Output When Updates Exist:**
```
abc1234 Add new feature X
def5678 Fix bug in Y
ghi9012 Improve performance
```

---

## Troubleshooting

### Problem: "Merge Conflicts"

**What it means:** Your custom changes and parent's changes modified the same lines of code.

**Solution (GitHub Desktop):**
1. GitHub Desktop will show conflicted files
2. Click **"Open in [your editor]"**
3. Look for conflict markers:
   ```
   <<<<<<< HEAD
   Your code
   =======
   Parent's code
   >>>>>>> upstream/main
   ```
4. Choose which version to keep (or combine both)
5. Remove the conflict markers
6. Save the file
7. In GitHub Desktop, mark as resolved and commit

**Solution (Command Line):**
```bash
# See which files have conflicts
git status

# Edit the conflicted files manually
# Remove conflict markers and keep desired code

# Mark as resolved
git add .
git commit -m "Resolve merge conflicts"
git push origin main
```

### Problem: "Your branch is behind 'origin/main'"

**Solution:**
```bash
git pull origin main
git push origin main
```

### Problem: "Cannot fetch upstream"

**Solution:** Verify upstream is configured:
```bash
git remote -v
```

You should see:
```
origin    https://github.com/omnidox/DesktopCommanderMCP.git (fetch)
origin    https://github.com/omnidox/DesktopCommanderMCP.git (push)
upstream  https://github.com/wonderwhy-er/DesktopCommanderMCP.git (fetch)
upstream  https://github.com/wonderwhy-er/DesktopCommanderMCP.git (push)
```

If `upstream` is missing, add it:
```bash
git remote add upstream https://github.com/wonderwhy-er/DesktopCommanderMCP.git
```

---

## Update Schedule Recommendations

### How Often Should You Update?

**Weekly** (Recommended for active development)
- Check every week for new features and bug fixes
- Minimal conflicts since changes are small and frequent

**Bi-Weekly/Monthly** (For stable usage)
- Check every 2-4 weeks if everything is working well
- Good balance between staying current and avoiding disruption

**As Needed** (Reactive approach)
- Only update when:
  - A bug affects your workflow
  - A new feature you need is added
  - Security updates are announced

**Before Major Work** (Best practice)
- Always sync before starting a large project
- Reduces merge conflicts later

### Quick Check Script

Save this as `check-updates.sh` and run it weekly:

```bash
#!/bin/bash
echo "🔍 Checking for updates from parent repository..."
git fetch upstream main 2>&1 > /dev/null

NEW_COMMITS=$(git log --oneline HEAD..upstream/main 2>/dev/null | wc -l)

if [ "$NEW_COMMITS" -eq 0 ]; then
    echo "✅ You are up to date!"
else
    echo "📦 $NEW_COMMITS new commit(s) available!"
    echo ""
    git log --oneline HEAD..upstream/main | head -10
fi
```

**Usage:**
```bash
bash check-updates.sh
```

---

## Verification After Update

After updating, verify everything still works:

### 1. Check Your Custom Features
```bash
# Verify Word document support is still present
grep "mammoth" package.json
```

**Expected output:**
```json
"mammoth": "^1.8.0",
```

### 2. Check Commit History
```bash
git log --oneline -5
```

**You should see:**
- Your Word document commit: "Add Microsoft Word (.docx) document reading support"
- Parent's latest commits below it

### 3. Verify Build
```bash
npm install
npm run build
```

Should complete without errors.

### 4. Test Word Document Reading
Create a test:
```bash
# Start the server
npm start

# In Claude Desktop, try reading a .docx file
# It should extract text successfully
```

---

## Summary Cheat Sheet

### GitHub Desktop (Quick Reference):
```
1. Fetch origin
2. Branch → Merge into current branch
3. Select upstream/main
4. Click Merge
5. Push origin
```

### Command Line (Quick Reference):
```bash
git fetch upstream main
git merge upstream/main
git push origin main
```

### Check for Updates:
```bash
git fetch upstream main
git log --oneline HEAD..upstream/main
```

---

## Need Help?

If you encounter issues:
1. Check the [Troubleshooting](#troubleshooting) section above
2. Review git status: `git status`
3. Check what branch you're on: `git branch`
4. Verify remotes are configured: `git remote -v`

---

**Last Updated:** October 21, 2025
**Your Fork:** omnidox/DesktopCommanderMCP
**Parent Repo:** wonderwhy-er/DesktopCommanderMCP
