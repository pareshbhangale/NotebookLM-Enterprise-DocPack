# How to use NotebookLM Enterprise DocPack

NotebookLM Enterprise DocPack is a cross-platform Python tool that downloads a documentation corpus, converts it to Markdown, and creates DOCX packs for use in an AI knowledge workspace. The Bash helpers are for macOS/Linux; Windows uses the Python entry point directly.

## 1. Set up the environment

On macOS or Linux:

```bash
./setup.sh
./run.sh --help
```

The setup script creates `.venv`, installs the dependencies from `requirements.txt`, and verifies the key imports.

On Windows PowerShell:

```powershell
python --version
python -m pip install -r requirements.txt
python notebooklm_enterprise_docpack.py --help
```

## 2. Create your manifest

Copy the supplied example and edit it to list only sources you are authorized to download and process.

```bash
cp Input/example_download_manifest.csv Input/manifest.csv
```

The manifest is CSV with a required URL field. `kind` and `title` are optional labels.

```csv
kind,title,url
pdf,Policy manual,https://docs.example.com/policy-manual.pdf
xlsx,Data dictionary,https://docs.example.com/data-dictionary.xlsx
zip,Technical reference,https://docs.example.com/technical-reference.zip
```

Accepted URL field names are `url`, `link`, and `source_url`.

## 3. Run the complete pipeline

```bash
./run.sh
```

The run performs these stages:

1. Downloads each manifest URL into `Output/raw`.
2. Extracts top-level ZIP archives into `Output/raw/_unzipped`.
3. Converts each source file to Markdown in `Output/markdown`.
4. Merges Markdown into `Output/knowledge_pack.md`, preserving a `# FILE:` marker for each source.
5. Writes DOCX packs to `Output/docx_packs`.

Review `Output/pipeline_summary.txt` after the run. It reports the files and stage counts used by that invocation.

## 4. Upload the packs

Upload all `.docx` files in `Output/docx_packs` into one AI workspace. Keeping all parts together lets the workspace retrieve from the full corpus.

Useful initial prompt:

```text
Use only the uploaded sources. State when an answer is not supported, cite the relevant document or section, and distinguish facts from inferences.
```

For very broad questions, ask for a short source-grounded summary first and then investigate a narrower topic.

## 5. Ask questions in a Codex task

When a Codex task is opened in this project folder, it can read the converted Markdown directly. Use prompts that name the corpus, constrain the answer to documented material, and request source paths.

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

Replace each placeholder with a specific topic, metric, table, process, or policy from your corpus. `Output/markdown` is especially useful in Codex because each converted file remains individually addressable.

## 6. Process local files without downloading

Place local documents in `Output/raw`, then skip the download stage:

```bash
./run.sh --skip-download
```

Use `--raw-dir` if you want a different raw folder within the project.

## 7. Reuse earlier stages

Each stage clears generated files for itself and downstream stages before it runs. This prevents stale packs from being mixed with a new corpus. Before clearing anything, NotebookLM Enterprise DocPack verifies that any input required by a skipped stage exists.

Examples:

```bash
# Rebuild DOCX packs from the existing merged Markdown corpus.
./run.sh --skip-download --skip-convert --skip-merge

# Re-merge existing Markdown and rebuild the DOCX packs.
./run.sh --skip-download --skip-convert
```

Do not use skip flags for a stage whose required input does not already exist.

## 8. Tune pack size

The defaults target 375,000 source words per pack and warn when the generated DOCX exceeds 400,000 words. Adjust them when your target AI service has smaller source limits:

```bash
./run.sh --chunk-words 50000 --pack-words 150000 --max-words 175000 --max-docx-parts 20
```

- `--chunk-words` limits a single source fragment.
- `--pack-words` is the preferred target size for a DOCX pack.
- `--max-words` is the word-count warning threshold for a generated pack.
- `--max-docx-parts` caps the number of output packs; use `0` to disable the cap.

Keep `chunk-words` and `pack-words` at or below `max-words`.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| Download failure | Confirm the URL is reachable without an interactive login and that your network allows it. |
| Conversion failure | Check the source type and install dependencies again with `./setup.sh`. The pipeline logs the affected file and continues. |
| Empty or weak output | Inspect the corresponding Markdown file to see whether the source contains extractable text. Scanned PDFs may need OCR before processing. |
| Unsafe ZIP warning | The archive contains a path that could write outside the extraction folder; NotebookLM Enterprise DocPack intentionally skips it. |
| Too many packs | Raise `--max-docx-parts`, use `0`, or increase `--pack-words` without exceeding `--max-words`. |
| AI workspace misses a topic | Confirm every DOCX part was uploaded and ask a narrower, source-grounded question. |

## Data responsibility

Do not place credentials in the manifest. The tool does not handle login flows or protected-site authentication. Inspect generated content and ensure it is appropriate to upload to your chosen AI service.
