# Diagnostic Check for Desktop Commander MCP Fork

This document provides commands to verify that your fork is properly configured and functioning correctly.

## Table of Contents
- [Quick Health Check](#quick-health-check)
- [Detailed Diagnostics](#detailed-diagnostics)
- [Automated Diagnostic Script](#automated-diagnostic-script)
- [Troubleshooting Failed Checks](#troubleshooting-failed-checks)

---

## Quick Health Check

Run these commands to get a quick overview:

### 1. Check Current Branch
```bash
git branch --show-current
```
**Expected output:** `main`

### 2. Check Git Status
```bash
git status
```
**Expected output:** `On branch main` and `nothing to commit, working tree clean`

### 3. Check Remote Configuration
```bash
git remote -v
```
**Expected output:**
```
origin    https://github.com/omnidox/DesktopCommanderMCP.git (fetch)
origin    https://github.com/omnidox/DesktopCommanderMCP.git (push)
upstream  https://github.com/wonderwhy-er/DesktopCommanderMCP.git (fetch)
upstream  https://github.com/wonderwhy-er/DesktopCommanderMCP.git (push)
```

### 4. Check Recent Commits
```bash
git log --oneline -5
```
**Expected to see:**
- "Add comprehensive guide for updating fork from parent repository"
- "Add Microsoft Word (.docx) document reading support"
- Recent commits from parent repository

### 5. Verify Word Document Feature
```bash
grep "mammoth" package.json
```
**Expected output:** `"mammoth": "^1.8.0",`

---

## Detailed Diagnostics

### Git Configuration Checks

#### Check Branch Tracking
```bash
git branch -vv
```
**What to look for:**
- `main` branch should be tracking a remote branch
- Should show `[origin/...]` indicating proper tracking

#### Check Upstream Connection
```bash
git ls-remote upstream
```
**What to look for:**
- Should list references from the parent repository
- No errors about "repository not found"

#### Check Sync Status with Parent
```bash
git fetch upstream main && git log --oneline HEAD..upstream/main
```
**What to look for:**
- **No output** = You're up to date with parent ✅
- **Commits listed** = Updates available from parent

#### Check Sync Status with Your Fork
```bash
git fetch origin && git log --oneline origin/main..HEAD
```
**What to look for:**
- **No output** = Local matches GitHub ✅
- **Commits listed** = You have unpushed local commits

---

### Repository Structure Checks

#### Verify Key Files Exist
```bash
ls -la | grep -E "(package.json|README.md|UPDATING_FROM_PARENT.md)"
```

Or on Windows:
```bash
dir | findstr /I "package.json README.md UPDATING_FROM_PARENT.md"
```

**Expected files:**
- `package.json`
- `README.md`
- `UPDATING_FROM_PARENT.md`

#### Check Source Files
```bash
ls -la src/tools/filesystem.ts src/server.ts
```

Or on Windows:
```bash
dir src\tools\filesystem.ts src\server.ts
```

**Expected:**
- Both files should exist
- No "File Not Found" errors

#### Verify Documentation
```bash
ls -la document_info/word-document-implementation-summary.md
```

Or on Windows:
```bash
dir document_info\word-document-implementation-summary.md
```

**Expected:**
- File should exist

---

### Feature Verification Checks

#### 1. Check Mammoth Import in Code
```bash
grep -n "import mammoth" src/tools/filesystem.ts
```

Or on Windows:
```bash
findstr /N "import mammoth" src\tools\filesystem.ts
```

**Expected output:** `8:import mammoth from 'mammoth';`

#### 2. Check Word Document Detection Logic
```bash
grep -n "isWordDoc" src/tools/filesystem.ts
```

Or on Windows:
```bash
findstr /N "isWordDoc" src\tools\filesystem.ts
```

**Expected:**
- Should find multiple references
- Around lines 681, 692

#### 3. Check Tool Descriptions Updated
```bash
grep -n "Microsoft Word" src/server.ts
```

Or on Windows:
```bash
findstr /N "Microsoft Word" src\server.ts
```

**Expected:**
- Should find references to Word document support
- Around lines 256, 276

#### 4. Verify Package.json Dependencies
```bash
cat package.json | grep -A 10 '"dependencies"'
```

Or on Windows:
```bash
type package.json | findstr /A "dependencies" /C:"mammoth"
```

**Expected dependencies should include:**
- `@modelcontextprotocol/sdk`
- `@vscode/ripgrep`
- `mammoth`
- `zod`

---

### Build and Dependencies Checks

#### Check Node.js Version
```bash
node --version
```
**Expected:** `v18.0.0` or higher

#### Check NPM Version
```bash
npm --version
```
**Expected:** Any recent version (7.x or higher)

#### Verify Dependencies Installed
```bash
ls -la node_modules/mammoth
```

Or on Windows:
```bash
dir node_modules\mammoth
```

**Expected:**
- Directory should exist if `npm install` was run
- If not found, run: `npm install`

#### Test Build Process
```bash
npm run build
```

**Expected:**
- Should compile TypeScript files
- Create `dist/` directory
- No critical errors (type warnings are okay)

#### Check Build Output
```bash
ls -la dist/index.js dist/server.js dist/tools/filesystem.js
```

Or on Windows:
```bash
dir dist\index.js dist\server.js dist\tools\filesystem.js
```

**Expected:**
- All three files should exist after build
- If missing, run: `npm run build`

---

## Automated Diagnostic Script

Save this as `diagnostic.sh` (Mac/Linux) or `diagnostic.bat` (Windows):

### For Mac/Linux: `diagnostic.sh`

```bash
#!/bin/bash

echo "=========================================="
echo "Desktop Commander MCP Fork Diagnostics"
echo "=========================================="
echo ""

# Color codes
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Track pass/fail
CHECKS_PASSED=0
CHECKS_FAILED=0

check_result() {
    if [ $1 -eq 0 ]; then
        echo -e "${GREEN}✓ PASS${NC}: $2"
        ((CHECKS_PASSED++))
    else
        echo -e "${RED}✗ FAIL${NC}: $2"
        ((CHECKS_FAILED++))
    fi
}

echo "1. Git Configuration Checks"
echo "----------------------------"

# Check if on main branch
CURRENT_BRANCH=$(git branch --show-current)
if [ "$CURRENT_BRANCH" = "main" ]; then
    check_result 0 "On main branch"
else
    check_result 1 "Current branch is '$CURRENT_BRANCH', expected 'main'"
fi

# Check if origin remote exists
git remote | grep -q "origin"
check_result $? "Origin remote configured"

# Check if upstream remote exists
git remote | grep -q "upstream"
check_result $? "Upstream remote configured"

# Check working tree is clean
git diff-index --quiet HEAD --
check_result $? "Working tree is clean"

echo ""
echo "2. Repository Structure Checks"
echo "------------------------------"

# Check key files exist
[ -f "package.json" ]
check_result $? "package.json exists"

[ -f "UPDATING_FROM_PARENT.md" ]
check_result $? "UPDATING_FROM_PARENT.md exists"

[ -f "src/tools/filesystem.ts" ]
check_result $? "src/tools/filesystem.ts exists"

[ -f "src/server.ts" ]
check_result $? "src/server.ts exists"

[ -f "document_info/word-document-implementation-summary.md" ]
check_result $? "Word implementation docs exist"

echo ""
echo "3. Feature Verification Checks"
echo "-------------------------------"

# Check mammoth dependency
grep -q '"mammoth"' package.json
check_result $? "Mammoth dependency in package.json"

# Check mammoth import
grep -q "import mammoth" src/tools/filesystem.ts
check_result $? "Mammoth import in filesystem.ts"

# Check Word document detection
grep -q "isWordDoc" src/tools/filesystem.ts
check_result $? "Word document detection logic present"

# Check tool descriptions updated
grep -q "Microsoft Word" src/server.ts
check_result $? "Server tool descriptions mention Word support"

echo ""
echo "4. Sync Status Checks"
echo "---------------------"

# Check sync with upstream
git fetch upstream main 2>&1 > /dev/null
UPSTREAM_COMMITS=$(git log --oneline HEAD..upstream/main 2>/dev/null | wc -l)
if [ "$UPSTREAM_COMMITS" -eq 0 ]; then
    check_result 0 "Up to date with parent repository"
else
    echo -e "${YELLOW}⚠ INFO${NC}: $UPSTREAM_COMMITS new commit(s) available from parent"
fi

# Check sync with origin
git fetch origin 2>&1 > /dev/null
UNPUSHED_COMMITS=$(git log --oneline origin/main..HEAD 2>/dev/null | wc -l)
if [ "$UNPUSHED_COMMITS" -eq 0 ]; then
    check_result 0 "All commits pushed to GitHub"
else
    echo -e "${YELLOW}⚠ INFO${NC}: $UNPUSHED_COMMITS unpushed commit(s)"
fi

echo ""
echo "5. Dependencies and Build Checks"
echo "--------------------------------"

# Check Node.js version
NODE_VERSION=$(node --version 2>/dev/null)
if [ ! -z "$NODE_VERSION" ]; then
    check_result 0 "Node.js installed ($NODE_VERSION)"
else
    check_result 1 "Node.js not found"
fi

# Check if node_modules exists
[ -d "node_modules" ]
if [ $? -eq 0 ]; then
    check_result 0 "Dependencies installed (node_modules exists)"
else
    echo -e "${YELLOW}⚠ INFO${NC}: Run 'npm install' to install dependencies"
fi

# Check if mammoth module exists
[ -d "node_modules/mammoth" ]
check_result $? "Mammoth module installed"

# Check if dist directory exists
[ -d "dist" ]
if [ $? -eq 0 ]; then
    check_result 0 "Build output exists (dist directory)"
else
    echo -e "${YELLOW}⚠ INFO${NC}: Run 'npm run build' to build the project"
fi

echo ""
echo "=========================================="
echo "Summary"
echo "=========================================="
echo -e "${GREEN}Checks Passed: $CHECKS_PASSED${NC}"
if [ $CHECKS_FAILED -gt 0 ]; then
    echo -e "${RED}Checks Failed: $CHECKS_FAILED${NC}"
else
    echo -e "${GREEN}Checks Failed: 0${NC}"
fi
echo ""

if [ $CHECKS_FAILED -eq 0 ]; then
    echo -e "${GREEN}✓ All checks passed! Your fork is properly configured.${NC}"
    exit 0
else
    echo -e "${RED}✗ Some checks failed. See DIAGNOSTIC_CHECK.md for troubleshooting.${NC}"
    exit 1
fi
```

**To run:**
```bash
chmod +x diagnostic.sh
./diagnostic.sh
```

### For Windows: `diagnostic.bat`

```batch
@echo off
echo ==========================================
echo Desktop Commander MCP Fork Diagnostics
echo ==========================================
echo.

set CHECKS_PASSED=0
set CHECKS_FAILED=0

echo 1. Git Configuration Checks
echo ----------------------------

REM Check if on main branch
for /f %%i in ('git branch --show-current') do set CURRENT_BRANCH=%%i
if "%CURRENT_BRANCH%"=="main" (
    echo [PASS] On main branch
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Current branch is '%CURRENT_BRANCH%', expected 'main'
    set /a CHECKS_FAILED+=1
)

REM Check remotes
git remote | findstr "origin" >nul
if %errorlevel%==0 (
    echo [PASS] Origin remote configured
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Origin remote not found
    set /a CHECKS_FAILED+=1
)

git remote | findstr "upstream" >nul
if %errorlevel%==0 (
    echo [PASS] Upstream remote configured
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Upstream remote not found
    set /a CHECKS_FAILED+=1
)

echo.
echo 2. Repository Structure Checks
echo ------------------------------

if exist "package.json" (
    echo [PASS] package.json exists
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] package.json not found
    set /a CHECKS_FAILED+=1
)

if exist "UPDATING_FROM_PARENT.md" (
    echo [PASS] UPDATING_FROM_PARENT.md exists
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] UPDATING_FROM_PARENT.md not found
    set /a CHECKS_FAILED+=1
)

if exist "src\tools\filesystem.ts" (
    echo [PASS] src\tools\filesystem.ts exists
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] src\tools\filesystem.ts not found
    set /a CHECKS_FAILED+=1
)

if exist "src\server.ts" (
    echo [PASS] src\server.ts exists
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] src\server.ts not found
    set /a CHECKS_FAILED+=1
)

echo.
echo 3. Feature Verification Checks
echo -------------------------------

findstr /C:"mammoth" package.json >nul
if %errorlevel%==0 (
    echo [PASS] Mammoth dependency in package.json
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Mammoth dependency not found
    set /a CHECKS_FAILED+=1
)

findstr /C:"import mammoth" src\tools\filesystem.ts >nul
if %errorlevel%==0 (
    echo [PASS] Mammoth import in filesystem.ts
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Mammoth import not found
    set /a CHECKS_FAILED+=1
)

findstr /C:"isWordDoc" src\tools\filesystem.ts >nul
if %errorlevel%==0 (
    echo [PASS] Word document detection logic present
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Word document detection not found
    set /a CHECKS_FAILED+=1
)

findstr /C:"Microsoft Word" src\server.ts >nul
if %errorlevel%==0 (
    echo [PASS] Server tool descriptions mention Word support
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Word support not mentioned in server descriptions
    set /a CHECKS_FAILED+=1
)

echo.
echo 4. Dependencies and Build Checks
echo --------------------------------

node --version >nul 2>&1
if %errorlevel%==0 (
    for /f %%i in ('node --version') do echo [PASS] Node.js installed (%%i)
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Node.js not found
    set /a CHECKS_FAILED+=1
)

if exist "node_modules" (
    echo [PASS] Dependencies installed (node_modules exists)
    set /a CHECKS_PASSED+=1
) else (
    echo [INFO] Run 'npm install' to install dependencies
)

if exist "node_modules\mammoth" (
    echo [PASS] Mammoth module installed
    set /a CHECKS_PASSED+=1
) else (
    echo [FAIL] Mammoth module not found
    set /a CHECKS_FAILED+=1
)

if exist "dist" (
    echo [PASS] Build output exists (dist directory)
    set /a CHECKS_PASSED+=1
) else (
    echo [INFO] Run 'npm run build' to build the project
)

echo.
echo ==========================================
echo Summary
echo ==========================================
echo Checks Passed: %CHECKS_PASSED%
echo Checks Failed: %CHECKS_FAILED%
echo.

if %CHECKS_FAILED%==0 (
    echo [SUCCESS] All checks passed! Your fork is properly configured.
    exit /b 0
) else (
    echo [WARNING] Some checks failed. See DIAGNOSTIC_CHECK.md for troubleshooting.
    exit /b 1
)
```

**To run:**
```batch
diagnostic.bat
```

---

## Troubleshooting Failed Checks

### "Origin remote not configured"
```bash
git remote add origin https://github.com/omnidox/DesktopCommanderMCP.git
```

### "Upstream remote not configured"
```bash
git remote add upstream https://github.com/wonderwhy-er/DesktopCommanderMCP.git
```

### "Mammoth dependency not found"
```bash
npm install mammoth --save
```

### "Dependencies not installed"
```bash
npm install
```

### "Build output not found"
```bash
npm run build
```

### "Not on main branch"
```bash
git checkout main
```

### "Working tree not clean"
```bash
git status  # See what changed
git add .   # Stage changes
git commit -m "Your commit message"
```

### "Unpushed commits"
```bash
git push origin main
```

### "Behind parent repository"
See `UPDATING_FROM_PARENT.md` for full instructions, or quick update:
```bash
git fetch upstream main
git merge upstream/main
git push origin main
```

---

## Quick One-Line Diagnostic

For a super quick check, run this single command:

### Mac/Linux:
```bash
echo "Branch: $(git branch --show-current) | Status: $(git status --short | wc -l) changes | Mammoth: $(grep -c mammoth package.json) | Remotes: $(git remote | wc -l)"
```

### Windows:
```batch
git branch --show-current && git status --short && findstr "mammoth" package.json && git remote
```

**Expected output indicators:**
- Branch: `main`
- Status: `0 changes` (or empty)
- Mammoth: `1` (found)
- Remotes: `2` (origin and upstream)

---

## Regular Health Check Schedule

**Weekly:** Run the automated diagnostic script
**Before Updates:** Run sync status checks
**After Updates:** Run full diagnostics
**After Changes:** Run feature verification checks

---

**Last Updated:** October 21, 2025
**Repository:** omnidox/DesktopCommanderMCP
