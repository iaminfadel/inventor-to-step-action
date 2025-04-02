# Inventor to STEP Action

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub Workflow Status](https://img.shields.io/github/actions/workflow/status/iaminfadel/inventor-to-step-action/export-inventor.yml?label=Export%20Status)](https://github.com/iaminfadel/inventor-to-step-action/actions)

Custom Github Action workflow to export Inventor .ipt files to STEP files automatically when pushed to the repository.

## Requirements
This action requires Inventor to be installed on the GitHub runner.
GitHub's default Windows runners don't have Inventor installed, so you have to use a self-hosted runner with Inventor installed.

## Self-hosted Runner Setup
Since you'll need Inventor installed:
1. Set up a Windows computer with Inventor installed as a GitHub self-hosted runner
2. In your GitHub repository, go to `Settings > Actions > Runners > New self-hosted runner`
3. Follow the instructions to connect your computer

## How It Works

The action automatically:
1. Triggers when `.ipt` files are pushed to the `main` branch
2. Scans the repository for all `.ipt` files
3. Exports each file to STEP format in a `STEP_Exports` subdirectory
4. Commits and pushes the generated STEP files back to the repository

## Workflow Configuration

The workflow is pre-configured in `.github/workflows/export-inventor.yml`. No additional configuration is needed - just push your Inventor files!

Example workflow usage:
```yaml
name: Export Inventor Files
on:
  push:
    branches: [ main ]
    paths:
      - '**/*.ipt'

jobs:
  export-to-step:
    runs-on: self-hosted
    steps:
      - uses: actions/checkout@v3
      # ... rest of the workflow
```

## Outputs

- Generated STEP files are placed in a `STEP_Exports` directory next to the source `.ipt` files
- Files are automatically committed back to the repository
- Each STEP file maintains the same base name as its source `.ipt` file

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
