# Automated Sync Setup

This document explains how the automated synchronization works between the main work-scripts repository and the e3-automation-tools public repository.

## Overview

The sync system automatically copies E3 automation files from the private work-scripts repository to the public e3-automation-tools repository whenever changes are made to relevant files.

## Architecture

### 1. Trigger Workflow (Main Repository)
**File**: `.github/workflows/trigger-e3-sync.yml` (in work-scripts repo)

This workflow monitors changes to E3-related files and triggers the sync:

**Monitored Files:**
- `lib/e3_wire_numbering.py` - Wire numbering library
- `lib/e3_device_designation.py` - Device designation library  
- `lib/e3_terminal_pin_names.py` - Terminal pin names library
- `apps/e3_NA_Standards.py` - NA Standards GUI application
- `apps/set_wire_numbers.py` - Legacy wire numbering script
- `README_*.md` - Documentation files
- `requirements.txt` - Dependencies

**Trigger Conditions:**
- Push to `master` or `main` branch
- Changes to any monitored file
- Sends repository dispatch event to public repo

### 2. Sync Workflow (Public Repository)
**File**: `.github/workflows/sync-from-main.yml` (in e3-automation-tools repo)

This workflow receives the trigger and performs the actual sync:

**Actions Performed:**
1. Checks out both repositories
2. Copies library modules to `lib/` directory
3. Copies GUI application to `gui/` directory  
4. Copies legacy scripts to `scripts/` directory
5. Updates documentation in `docs/` directory
6. Updates requirements.txt with relevant dependencies
7. Commits and pushes changes
8. Creates automatic release with timestamp

## File Mapping

| Source (work-scripts) | Destination (e3-automation-tools) | Purpose |
|----------------------|-----------------------------------|---------|
| `lib/e3_wire_numbering.py` | `lib/e3_wire_numbering.py` | Wire numbering library |
| `lib/e3_device_designation.py` | `lib/e3_device_designation.py` | Device designation library |
| `lib/e3_terminal_pin_names.py` | `lib/e3_terminal_pin_names.py` | Terminal pin names library |
| `apps/e3_NA_Standards.py` | `gui/e3_NA_Standards.py` | NA Standards GUI |
| `apps/set_wire_numbers.py` | `scripts/set_wire_numbers.py` | Legacy script (compatibility) |
| `README_wire_numbering.md` | `docs/wire_numbering.md` | Wire numbering documentation |
| `README_device_designation.md` | `docs/device_designation.md` | Device designation docs |
| `README_terminal_pin_names.md` | `docs/terminal_pin_names.md` | Terminal pin names docs |

## Repository Structure

### Public Repository (e3-automation-tools)
```
e3-automation-tools/
├── lib/                    # Library modules
│   ├── e3_wire_numbering.py
│   ├── e3_device_designation.py
│   └── e3_terminal_pin_names.py
├── gui/                    # GUI applications
│   └── e3_NA_Standards.py
├── scripts/                # Legacy scripts
│   └── set_wire_numbers.py
├── docs/                   # Documentation
│   ├── wire_numbering.md
│   ├── device_designation.md
│   ├── terminal_pin_names.md
│   ├── na_standards_gui.md
│   └── automated_sync.md
├── .github/workflows/      # Sync workflows
├── requirements.txt        # Dependencies
├── README.md              # Main documentation
└── LICENSE                # License file
```

## Setup Requirements

### GitHub Secrets
The following secrets must be configured in the repositories:

**Main Repository (work-scripts):**
- `E3_REPO_TOKEN` - Personal access token with write access to public repo

**Public Repository (e3-automation-tools):**
- `MAIN_REPO_TOKEN` - Personal access token with read access to private repo

### Token Permissions
The tokens need the following permissions:
- `repo` - Full repository access
- `workflow` - Workflow permissions
- `contents:write` - Content modification
- `metadata:read` - Repository metadata

## Manual Sync

If automatic sync fails or manual sync is needed:

1. **Trigger from Main Repository:**
   - Go to Actions tab in work-scripts repository
   - Find "Trigger E3 Tools Sync" workflow
   - Click "Run workflow" button

2. **Trigger from Public Repository:**
   - Go to Actions tab in e3-automation-tools repository  
   - Find "Sync from Main Repository" workflow
   - Click "Run workflow" button

## Troubleshooting

### Common Issues

**Sync not triggering:**
- Check that changed files are in the monitored paths list
- Verify GitHub secrets are properly configured
- Ensure tokens have sufficient permissions

**Sync fails:**
- Check workflow logs in both repositories
- Verify source files exist in main repository
- Check for file permission issues

**Missing files after sync:**
- Verify file paths in sync workflow match actual locations
- Check for typos in file names or paths
- Ensure source files are committed to main repository

### Monitoring

- Check the Actions tab in both repositories for workflow status
- Review workflow logs for detailed error information
- Monitor the releases page in public repository for successful syncs

## Maintenance

### Adding New Files
To add new files to the sync:

1. Update `trigger-e3-sync.yml` in main repository (paths section)
2. Update `sync-from-main.yml` in public repository (copy steps)
3. Update `setup_repository.py` if needed for initial setup

### Removing Files
To remove files from sync:

1. Remove from trigger workflow paths
2. Remove copy steps from sync workflow  
3. Manually delete from public repository if needed

The automated sync ensures the public repository stays current with E3 automation developments while maintaining separation between private and public codebases.
