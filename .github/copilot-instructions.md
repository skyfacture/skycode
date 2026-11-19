# VBA macro development

- Never add an `Attribute VB_Name = ...` line to macro source.
- For Composer view-creation macros using `ZoomFitAll`, save
  `CameraSmoothTime`, set it to `0` before the first zoom, and restore it on
  completion, cancellation, and errors. Report restore failures.
