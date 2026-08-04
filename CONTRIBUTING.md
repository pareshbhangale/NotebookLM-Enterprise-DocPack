# Contributing

Thanks for helping improve NotebookLM Enterprise DocPack.

## Before opening a pull request

1. Keep source documents, generated output, credentials, and local manifests out of the repository.
2. Make focused changes and update the documentation when behavior or commands change.
3. Run the test suite:

   ```bash
   python3 -m unittest discover -s tests -v
   ```

4. Confirm the command-line interface still works:

   ```bash
   ./run.sh --help
   ```

## Reporting issues

Use GitHub Issues for reproducible defects, documentation improvements, and feature requests. Do not include confidential documents, source URLs that require access, or credentials in an issue.

For security-sensitive reports, follow [SECURITY.md](SECURITY.md).
