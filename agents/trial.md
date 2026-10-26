# Trial

Input: a small sample of the real data.
Output: a trial build + an early render (LibreOffice headless -> PDF/PNG
if available) to catch broken pages/slides before the full build.

Trial exists to fail fast: a broken template, a bad chart range, or a
timing XML error should be found here, not after the full document is
built.
