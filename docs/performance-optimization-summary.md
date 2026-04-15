# Performance Optimization Summary

## Goal

Reduce end-to-end conversion time for the Esfer@ Acta Extractor web app, with special attention to PDF-to-XLSX generation and batch processing, while preserving existing behavior.

## What Succeeded

- The Excel/export pipeline optimization succeeded and was the only performance change promoted to production.
- That work reduced workbook serialization churn, formatting overhead, redundant PDF reopening, and unnecessary ZIP compression.
- The production-ready optimization was merged into `main`, pushed to GitHub, validated on the beta subdomain, and deployed successfully to production.

## What Is Live

- `main` contains the deployed optimization.
- The production deployment is based on the merged Excel/export optimization path.
- The beta deployment pack under `deploy/perf/` was used successfully for isolated validation before production rollout.

## What Was Tested And Rejected

The later PDF-focused performance work did not produce a safe, meaningful deployable win on the real local corpus.

- `perf/pdf-extraction-optimization`
  - Added instrumentation and benchmark helpers for the PDF extraction path.
  - Tested `pdfplumber` settings changes, trailing-page early exit, structural prefilters, and deeper internal profiling.
  - Result: no candidate delivered a strong enough speedup with sufficient correctness confidence.

- `research/alternative-pdf-extractor`
  - Added a research-only harness to compare alternative table extractors against the current `pdfplumber` path.
  - Tested practical alternatives such as PyMuPDF and Camelot on the local corpus.
  - Result: no candidate was both fast enough and compatible enough to justify integration.

These branches are preserved as research/history branches and should not be merged at this time.

## Preserved Research Branches

- `perf/pdf-extraction-optimization`
- `research/alternative-pdf-extractor`

They are useful for:

- performance instrumentation
- benchmark helpers
- rejected experiment history
- future extraction research if the PDF path is revisited later

## Plausible Future Directions

If PDF extraction performance is revisited later, the most plausible next steps are:

- a narrowly scoped prototype around a materially different extraction engine
- a hybrid fallback design only if a candidate clearly wins on a well-defined subset of documents
- deeper research into the extraction engine internals only if backed by new corpus evidence

Until such evidence exists, the current production path should remain the Excel/export-optimized implementation in `main`.
