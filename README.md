# NotebookLM Enterprise DocPack

Cross-platform Python tooling to download, unpack, normalize, and package enterprise documentation into source-labeled, upload-ready knowledge packs for ChatGPT Projects, NotebookLM, and similar AI tools.

Documentation collections rarely arrive in one upload-friendly format. A corpus may include PDFs, spreadsheets, web pages, Office files, ZIP archives, and deeply nested document trees. Uploading those files directly can leave content hidden in archives, fragmented across formats, or too large to ingest reliably.

NotebookLM Enterprise DocPack turns that collection into consistently formatted DOCX packs that are easier to upload, search, and ground against.

```text
CSV URL manifest
  → downloads
  → safe ZIP extraction
  → Markdown conversion
  → source-labeled merged corpus
  → size-controlled DOCX packs
```

## What it does

- Reads a CSV manifest of source URLs.
- Downloads source files into `Output/raw`.
- Extracts top-level ZIP archives safely and processes their document trees.
- Converts supported documents to Markdown using [MarkItDown](https://github.com/microsoft/markitdown).
- Merges converted files with a `# FILE:` source marker for each document.
- Splits the corpus into source-labeled DOCX packs with configurable word limits.
- Keeps generated downloads and outputs out of Git by default.

Supported source types depend on MarkItDown and the installed optional converters, and commonly include PDFs, Excel workbooks, HTML, Word files, PowerPoint files, text, and documents inside ZIP archives.

## Why prepare documents first?

Preparing the corpus improves the likelihood that your AI workspace sees the relevant text:

- ZIP-contained and nested documentation is extracted instead of being treated as an opaque attachment.
- Different document types are normalized to one text-oriented representation before packaging.
- Every converted document retains a source boundary in the merged corpus and output packs.
- Large corpora are partitioned into predictable DOCX files instead of relying on a single oversized upload.
- The process is reproducible: update the manifest and rerun it when documentation changes.

This does not change the source material's authority, access restrictions, or licensing. Only download and upload content you are permitted to use.

## Quick start

Requirements: Python 3.10 or later and network access to the URLs you list.

The Python pipeline runs on macOS, Linux, and Windows. `setup.sh` and `run.sh` are Bash helpers for macOS/Linux; Windows users run `notebooklm_enterprise_docpack.py` directly with Python.

```bash
./setup.sh
cp Input/example_download_manifest.csv Input/manifest.csv
# Edit Input/manifest.csv to contain the documents you are authorized to use.
./run.sh
```

The generated packs are written to `Output/docx_packs`. Upload every `.docx` file in that folder to one ChatGPT Project, NotebookLM notebook, or comparable AI workspace.

For Windows, install dependencies and run the Python entry point directly:

```powershell
python -m pip install -r requirements.txt
Copy-Item Input/example_download_manifest.csv Input/manifest.csv
python notebooklm_enterprise_docpack.py
```

Read the full guide in [HOW_TO_USE.md](HOW_TO_USE.md).

## Public-repository notes

- Generated downloads and knowledge packs are excluded from version control. Do not commit documents, manifests, or other data you are not authorized to share.
- The included example manifest contains public URLs only. Create `Input/manifest.csv` locally for your own sources; it is ignored by Git.
- NotebookLM is a Google product. This independent project is not affiliated with or endorsed by Google.

## Manifest format

`Input/manifest.csv` must include a URL column named `url`, `link`, or `source_url`. The `kind` and `title` columns are optional labels.

```csv
kind,title,url
pdf,Employee handbook,https://docs.example.com/handbook.pdf
zip,Product documentation,https://docs.example.com/product-docs.zip
```

`Input/example_download_manifest.csv` is a retained example URL list. It contains links only; no downloaded source documents are bundled with this repository.

## Outputs

| Path | Purpose |
| --- | --- |
| `Output/raw` | Downloaded documents and extracted top-level ZIP contents. |
| `Output/markdown` | One Markdown conversion per source file. |
| `Output/knowledge_pack.md` | The merged, source-labeled Markdown corpus. |
| `Output/docx_packs` | The upload-ready DOCX parts. |
| `Output/pipeline_summary.txt` | Run paths and stage counts. |

## Use with an AI workspace

Upload the complete contents of `Output/docx_packs`, then start with a grounded request such as:

```text
Answer only from the uploaded sources. If the sources do not support an answer, say so. Cite the most relevant document section and keep the answer concise.
```

## Use with a Codex task

Open a Codex task in this project folder to ask questions directly against the converted files in `Output/markdown`.

```text
Use only the documentation in Output/markdown.

Question 1

Cite the source file paths and do not infer anything that is not documented.
```

```text
Search Output/markdown for Question 2 and compare it with Question 3.

Give me a concise analyst version, then a technical table-and-field version.

Cite the source file paths and do not infer anything that is not documented.
```

```text
Read the relevant files in Output/markdown and create a mapping of Question 1, Question 2, and Question 3, with documented examples.

Cite the source file paths and do not infer anything that is not documented.
```

Replace the placeholders with precise topics from your documentation corpus. See [HOW_TO_USE.md](HOW_TO_USE.md) for the complete workflow.

## Command-line options

```bash
./run.sh --help
```

The useful options are:

- `--skip-download`, `--skip-convert`, `--skip-merge`, and `--skip-docx` to reuse earlier stages.
- `--chunk-words`, `--pack-words`, and `--max-words` to control pack size.
- `--max-docx-parts` to impose a maximum number of packs; `0` disables that limit.
- `--manifest`, `--raw-dir`, `--md-dir`, `--master-md`, `--docx-dir`, and `--summary-file` to use alternate paths inside the project directory.

## Security and privacy

NotebookLM Enterprise DocPack blocks ZIP path traversal and confines command-line paths to the project directory. It does not authenticate to protected sites, remove sensitive information, or grant you redistribution rights. Review output before sending it to an external AI service.

## License

NotebookLM Enterprise DocPack is available under the [MIT License](LICENSE).

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines and [SECURITY.md](SECURITY.md) for vulnerability reporting.
